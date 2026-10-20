#!/usr/bin/env python3
"""闪办（SwiftDone，Web MVP）：闪速处理，纯净办公。打工人专属的效率工具箱。

四大场景（对应的顶部 Tab），全部为纯本地处理，不调用任何外部 API：
    1. PDF基础处理   —— PDF 压缩 / PDF 转 Word / PDF 合并 / 文档转 PDF
    2. PDF/图片改字  —— PDF 等长替换 + 图片擦除重绘（支持多行规则）
    3. 批量生成      —— Excel + 模板 PDF -> 批量 PDF，打包下载
    4. 去水印        —— 图片 / PDF / Word / PPT

运行：
    本目录已带好 .venv，直接：
        .venv/bin/python -m streamlit run app.py
    （首次在新机器上运行先执行：python3 -m venv .venv && .venv/bin/pip install -r requirements.txt）

说明：
    - 所有处理都在内存或临时目录里完成，结果读成字节交给下载按钮，请求结束立即清理；
    - 核心逻辑都在 compress.py / converter.py / pdf_replace.py 里，本文件只做界面与编排；
    - 上传内容为空、文件损坏等异常都转成中文提示，不会让整页报错；
    - 所有耗时按钮都走 `run_task()` 的 busy 锁，处理期间按钮置灰，防止连点打满 CPU。

商业化分级（“无痕限额”体验：普通界面不出现任何“免费版 vs Pro”的字样，
只有用户真正触到上限时才弹出升级提示）：

    - 基础额度            —— 单次最多 3 份文件、单文件最大 20MB；
    - Pro 额度（兑换码解锁）—— 单次最多 20 份文件、单文件最大 100MB；
    - 限额判断统一走 `limits_active()` / `split_files()`（侧边栏兑换码解锁后自动跳过）；
    - 触发拦截时才出现：st.error（额度说明）+ st.info（去左侧输入兑换码），
      并且不执行任何处理逻辑（不发请求、不落盘）。

第三方服务策略（2026-10 起，已彻底定型）：
    产品定位就是“纯本地运行、绝对安全、文件不上传、零成本”。全部图像 / 文档处理都用
    本地库完成（PyMuPDF、Pillow、reportlab、python-docx，图片改字与去水印都是纯像素
    算法，不需要 OpenCV 也能跑），**代码中不存在任何外部 API 调用**：没有 API Key、
    没有 HTTP 请求、没有 OCR / 语音转写等需要联网的能力。原先的“批量提取”“AI 写周报”
    已连同其 Mock 代码与额度逻辑一并删除，界面不保留任何“敬请期待”式占位。
    PDF 转 Word 时字体自动下载也为关闭状态（converter 的 auto_download_fonts=False），
    同时关闭了内嵌字体抽取（extract_embedded_fonts=False），缺字体只回退到本机字体，
    不会联网，也不会把用户文档里的字体子集写进本机缓存目录。
"""

from __future__ import annotations

import html as html_lib
import json
import re
import tempfile
import traceback
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, Sequence
from xml.etree import ElementTree as ET

import streamlit as st

import compress
import converter
import pdf_replace

import batch_fill
import image_edit
import pandas as pd
import pymupdf
import watermark_remove

APP_NAME = "闪办"
# Hero 主 Slogan（深灰、中等字重）与副标题（浅灰、最小号）分开存放，方便单独调整层级
APP_TAGLINE = "闪速处理，纯净办公。"
APP_SUBTITLE = "你的本地效率工具箱"
# 兼容旧引用：整句口号 = 主 Slogan + 副标题，用于 page_title / 文档说明等处
APP_CLAIM = f"{APP_TAGLINE}{APP_SUBTITLE}"
APP_SLOGAN = "极速纯净的效率工具箱"  # 浏览器标签页使用的短口号
APP_ICON = "⚡️"
# 闪电 SVG：实心 Lucide Zap 造型（多边形 + round join，边缘锐利不毛糙）。
# 尺寸与发光全部交给 CSS（.zjb-mark），所以这里不写 width/height/filter，
# 避免内联属性覆盖样式表、也避免和 52px 的规格打架。
APP_BOLT_SVG = (
    '<svg class="zjb-mark" viewBox="0 0 24 24" aria-hidden="true" focusable="false"'
    ' fill="currentColor" stroke="currentColor" stroke-width="1.5"'
    ' stroke-linecap="round" stroke-linejoin="round">'
    '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>'
    "</svg>"
)
APP_ALIAS = "SwiftDone"               # 英文名，用于页脚版权
APP_COPYRIGHT = "© 2024 闪办 (SwiftDone). 无广告，纯粹好用。"

# 全局安全承诺：所有上传组件下方都要出现
PRIVACY_NOTE = "🛡️ 闪办承诺：文件在内存中极速处理，即刻销毁，绝不保留。"

# ==========================================================================
# 额度分级（无痕限额：界面常态不出现“免费版 / Pro”对比字样，
#           只有触到上限的那一刻才弹 st.error + st.info）
# ==========================================================================
# 基础额度
BASE_MAX_FILES = 3                   # 单次最多处理 3 份文件
BASE_MAX_UPLOAD_MB = 20              # 单文件上限 20MB
BASE_MAX_UPLOAD_BYTES = BASE_MAX_UPLOAD_MB * 1024 * 1024
# Pro 额度（兑换码解锁）
PRO_MAX_FILES = 20                   # 单次最多 20 份文件
PRO_MAX_UPLOAD_MB = 100              # 单文件上限 100MB
# 兑换码：写死在源码里的测试码，正式发行时换成服务端下发的码表
PRO_CODES = ("VIP-TEST-001",)

# 邀请码验证（进入主界面之前先校验）
VALID_CODES = ["VIP-001", "VIP-002", "XHS-001", "WX-001", "FRIEND-01"]

# 拦截时的两条提示（一字不改：红色额度说明 + 蓝色解锁指引）
LIMIT_MESSAGE = (
    f"⚠️ 单次最多处理 {BASE_MAX_FILES} 份文件，且单文件不能超过 "
    f"{BASE_MAX_UPLOAD_MB}MB。升级 Pro 版即可解锁单次 {PRO_MAX_FILES} 份、"
    f"单文件 {PRO_MAX_UPLOAD_MB}MB 的极速处理特权。"
)
UNLOCK_MESSAGE = "👉 请前往左侧边栏，输入您的 Pro 兑换码即可解除限制。"

# 上传组件内 “N per file” 那行文案取自 st.file_uploader(max_upload_size=)，
# 不传则回落到 server.maxUploadSize=100（见 .streamlit/config.toml）→ 会显示成
# “100MB per file” 与下方 20MB 提示冲突。这里固定传基础额度，让前端显示 20MB；
# 它只影响前端提示与浏览器本地校验，服务端上限仍是 100MB，Pro 用户照样能传大文件。
UPLOADER_WIDGET_MAX_MB = BASE_MAX_UPLOAD_MB
# 所有上传框统一 label + 下方提示语（灰字 caption，全站一致）
UPLOADER_LABEL = f"上传文件（普通用户限{BASE_MAX_UPLOAD_MB}MB）"
UPLOADER_CAPTION = (
    f"💡 提示：普通用户单文件最大 {BASE_MAX_UPLOAD_MB}MB；Pro 用户最大 "
    f"{PRO_MAX_UPLOAD_MB}MB，单次最多 {PRO_MAX_FILES} 份。"
)
# 上传框内拖拽区的说明：Streamlit 前端写死 “200MB per file”，只能用 CSS 覆盖，
# 因此要按页给文案（PDF 页不可能传 Excel，批量生成页也不会只传 PDF）。
DRAGZONE_TIP_PDF = f"支持格式：PDF，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_IMAGE_PDF = f"支持格式：PDF/JPG/PNG，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_PDF_IMG = f"支持格式：PDF/PNG/JPG/WEBP，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_EXCEL = f"支持格式：Excel/CSV，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_OFFICE = f"支持格式：PNG/JPG/WEBP/PDF/Word/PPT，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_AUDIO = f"支持格式：MP3/WAV/M4A，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"
DRAGZONE_TIP_DOC = f"支持格式：Word/Excel/PPT/图片，单文件最大 {BASE_MAX_UPLOAD_MB}MB。"

# 侧边栏输入的 Pro 兑换码存放位置（每次判断都重新校验，清掉即恢复限制）
PRO_CODE_KEY = "pro_code"

# 邀请码验证通过标记（session key）
VERIFY_KEY = "verified"

COMPRESS_LEVELS = {
    "无损压缩（推荐，画质与版式不变）": False,
    "极致压缩（目标 < 1MB，仅压缩照片）": True,
}
COMPRESS_TARGET = "1MB"

MERGE_HELP = "可一次选择多份，按上传顺序合并"

# Google Fonts：西文用 Inter，中文回退到系统苹方 / 微软雅黑
FONT_LINKS = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
"""

CUSTOM_CSS = """
<style>
/* ============================================================
   0. 设计变量
   ============================================================ */
:root {
    --zjb-bg: #F9FAFB;
    --zjb-surface: #FFFFFF;
    --zjb-navy: #0B2447;
    --zjb-navy-hover: #123A72;
    --zjb-accent: #2563EB;
    --zjb-accent-soft: #EFF5FF;
    --zjb-border: #E4E8EF;
    --zjb-text: #111827;
    --zjb-muted: #6B7280;
    --zjb-faint: #9CA3AF;
    --zjb-radius: 16px;
    --zjb-shadow: 0 1px 2px rgba(16, 24, 40, .04), 0 8px 24px rgba(16, 24, 40, .05);
    --zjb-shadow-lift: 0 2px 4px rgba(16, 24, 40, .05), 0 16px 36px rgba(16, 24, 40, .09);
}
/* ============================================================
    0.5 移动端适配：768px 以下不换行横向滑动导航栏、压缩顶部留白
   ============================================================ */
@media (max-width: 768px) {
    /* ── 导航栏横向滑动（不换行、可滑动、无滚动条） ── */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        display: flex !important;
        flex-wrap: nowrap !important;
        overflow-x: auto !important;
        white-space: nowrap !important;
        -webkit-overflow-scrolling: touch;
        scrollbar-width: none;           /* Firefox */
        -ms-overflow-style: none;        /* IE / Edge */
        padding-bottom: 0px !important;
        gap: 10px !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-list"]::-webkit-scrollbar {
        display: none !important;
    }

    /* Tab 按钮不收缩，固定最小宽度保证可滑 */
    [data-testid="stTabs"] button[data-baseweb="tab"] {
        flex-shrink: 0 !important;
        white-space: nowrap !important;
        padding: 5px 10px !important;
    }

    /* 全局缩小边距，内容更紧凑 */
    .block-container,
    [data-testid="stMainBlockContainer"] {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
    }

    /* 缩小标题字号 + 间距 */
    h1 { font-size: 1.5rem !important; margin-bottom: 0.5rem !important; }
    h2 { font-size: 1.2rem !important; margin-top: 0.5rem !important; margin-bottom: 0.5rem !important; }
    h3 { font-size: 1rem !important; margin-bottom: 0.5rem !important; }

    /* 正文缩小 */
    p { font-size: 0.9rem !important; margin-bottom: 0.5rem !important; }

    /* Hero 区收紧上下内边距，让上传按钮早一点露出来 */
    .zjb-hero {
        padding-top: 0.6rem !important;
        padding-bottom: 0.8rem !important;
    }

    /* 品牌名字号缩小 */
    .zjb-hero .zjb-name {
        font-size: clamp(28px, 8vw, 32px) !important;
    }

    /* 主 Slogan 缩小 */
    .zjb-hero .zjb-slogan {
        font-size: 16px !important;
        margin-top: 8px !important;
    }

    /* 副标题缩小 */
    .zjb-hero .zjb-tagline {
        font-size: 12px !important;
        margin-top: 4px !important;
    }

    /* 信任标签胶囊缩小 */
    .zjb-hero .zjb-badge {
        margin-bottom: 18px !important;
        font-size: 11px !important;
        padding: 5px 10px !important;
    }

    /* 闪电 SVG 图标缩小 */
    .zjb-hero .zjb-mark {
        width: 32px !important;
        height: 32px !important;
    }
}

/* ============================================================
   1. 全局：背景、居中容器、隐藏原生 Header / 菜单 / 页脚
   ============================================================ */
html, body, .stApp, .stMarkdown, button, input, textarea {
    font-family: "Inter", -apple-system, BlinkMacSystemFont, "PingFang SC",
                 "Hiragino Sans GB", "Microsoft YaHei", "Segoe UI", sans-serif;
    -webkit-font-smoothing: antialiased;
}

.stApp {background: var(--zjb-bg); color: var(--zjb-text);}

/* 顶部原生 Header 完全隐藏（菜单、Deploy、工具栏、装饰条） */
header[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
#MainMenu, footer {display: none !important;}

/* 页面内容：最大 900px 并水平居中，两侧自动留白。
   顶部 80px：原生 Header 是 position:fixed 的覆盖层，不出现在文档流里，
   所以这 80px 是纯视觉留白，保证胶囊标签完整可见、不被顶栏压住。 */
.block-container,
[data-testid="stMainBlockContainer"] {
    max-width: 900px !important;
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding: 80px 1.5rem 4rem 1.5rem !important;
}

/* 兜底：万一 Streamlit 版本把 header 换成占位块，也压掉它的高度 */
header[data-testid="stHeader"] {
    height: 0 !important;
    min-height: 0 !important;
    background: transparent !important;
}

/* 竖排区块统一节奏 */
[data-testid="stVerticalBlock"] {gap: .85rem;}

/* ============================================================
   2. Hero 区：胶囊信任标签 + 品牌名 + Slogan + 副标题
   ------------------------------------------------------------
   间距节奏（从上到下，官方规格值）：
     标签 ──margin-top: 32px──> Logo ──margin-top: 24px──> Slogan ──margin-top: 12px──> 副标题
     → 间距逐级递减（32 → 24 → 12），视线自然从上往下收，
       这是排版里最经典的「漏斗式留白」，比等距更有呼吸感。
   窄屏处理：px 值在手机上会显得过大，用 clamp() 给下限，
   保证 320px 宽也放得下，桌面端严格等于规格值。
   ============================================================ */
.zjb-hero {
    --hero-ink: #0f172a;          /* 品牌名：深藏青/近黑，最重 */
    --hero-strong: #334155;       /* Slogan：深灰 */
    --hero-sub: #94a3b8;          /* 副标题：浅灰，最弱 */
    --hero-badge-bg: #f4f4f5;     /* 胶囊背景：极浅灰 */
    --hero-badge-fg: #71717a;     /* 胶囊文字：中灰 */

    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    /* 上下留白给足，让首屏不贴边 */
    padding: clamp(2rem, 5vw, 3rem) 0 clamp(3rem, 6vw, 4rem) 0;
}

/* 入场动画：只动 opacity / transform，走合成层，不引起重排 */
@keyframes zjb-fade-up {
    from {opacity: 0; transform: translate3d(0, 12px, 0);}
    to   {opacity: 1; transform: translate3d(0, 0, 0);}
}

/* Hero 内所有段落统一清零默认边距——间距一律由各自 margin-top 显式控制，
   避免 Streamlit / 浏览器默认样式叠加导致「凑在一起」。
   注意：这里只清零 block-start，否则会把下面各元素自己的 margin-top 也抹掉 */
.zjb-hero p {
    margin-block-start: 0;
    margin-block-end: 0;
    line-height: 1.5;
}

/* ---------- 1) 信任标签：真正的胶囊 ---------- */
.zjb-hero .zjb-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    margin: 0 0 40px 0;                   /* 规格：与下方 Logo 留 40px */
    padding: 8px 16px;                    /* 规格：8px 16px */
    border: 1px solid rgba(24, 24, 27, .06);
    border-radius: 999px;                 /* 完整胶囊 */
    background: var(--hero-badge-bg);
    color: var(--hero-badge-fg) !important;   /* 规格：#71717a */
    font-size: 14px;                      /* 规格：14px */
    font-weight: 500;
    line-height: 20px;
    letter-spacing: .01em;
    white-space: nowrap;
    align-self: center;
    animation: zjb-fade-up .6s cubic-bezier(.22, .61, .36, 1) both;
}

/* 绿点：用元素画圆而非 ● 字符（● 在不同系统字宽不一致，容易歪） */
.zjb-hero .zjb-badge .dot {
    width: 7px; height: 7px;
    flex: 0 0 7px;
    border-radius: 50%;
    background: #22c55e;
    box-shadow: 0 0 0 3px rgba(34, 197, 94, .14);
}

/* ---------- 2) 品牌名：整页视觉核心 ---------- */
/* 注意：Streamlit 会把 <h1> 的子内容再包一层 <span>（st-emotion-cache-0，display:block），
   所以 flex 必须作用在「直接包住图标与文字的那个元素」上，否则 row 只会对 wrapper 生效，
   图标与文字会在 wrapper 内部按 block 上下堆叠。
   这里的 > span:not([data-testid]) 精确命中内容 wrapper，同时排除 Streamlit 注入的
   操作区 <span data-testid="stHeaderActionElements">。 */
.zjb-hero .zjb-name,
.zjb-hero .zjb-name > span:not([data-testid]) {
    display: flex;
    flex-direction: row;
    align-items: center;                  /* 图标与文字垂直居中对齐 */
    justify-content: center;
    gap: 12px;
}

.zjb-hero .zjb-name {
    /* 上方间距由胶囊的 margin-bottom: 40px 提供，这里清零避免叠加成 72px */
    margin: 0;
    padding: 0;
    color: var(--hero-ink) !important;        /* 规格：#0f172a */
    font-size: clamp(48px, 8.6vw, 64px);      /* 规格：64px，窄屏降到 48px */
    font-weight: 800;                         /* 规格：800 */
    line-height: 1.1;
    letter-spacing: -.02em;                   /* 大字号收紧字距，避免松散 */
    position: relative;                       /* 给 Streamlit 的操作区做定位基准 */
    animation: zjb-fade-up .6s cubic-bezier(.22, .61, .36, 1) .08s both;
}

/* Streamlit 会给每个标题注入一个隐藏操作区（16×16），让它脱离 flex 流，
   否则它会作为一个 flex item 占位、把「图标+文字」整体推向左侧 */
.zjb-hero .zjb-name > [data-testid="stHeaderActionElements"] {
    position: absolute;
    right: 0;
    top: 50%;
    transform: translateY(-50%);
    display: none;                            /* 无锚点/复制按钮时本就不需要 */
}

/* 文字部分：锁死行内排版，避免被 Streamlit 的段落样式干扰 */
.zjb-hero .zjb-name .zjb-word {
    display: inline-block;
    color: inherit !important;
    white-space: nowrap;
}

/* 闪电图标：Lucide Zap 造型（实心多边形 + 圆角连接），尺寸固定 52px */
.zjb-hero .zjb-mark {
    width: 52px;
    height: 52px;
    flex: 0 0 auto;
    display: block;
    color: #f59e0b;
    /* 克制的发光：紧贴边缘的 8px 柔光，不做大面积模糊 */
    filter: drop-shadow(0 0 8px rgba(245, 158, 11, .4));
}

/* ---------- 3) 主 Slogan ---------- */
.zjb-hero .zjb-slogan {
    margin: clamp(18px, 4vw, 24px) 0 0 0;     /* 规格：mt-24px */
    color: var(--hero-strong) !important;      /* 规格：#334155 */
    font-size: clamp(20px, 4.4vw, 24px);       /* 规格：24px */
    font-weight: 500;                          /* 规格：500 */
    line-height: 1.45;
    letter-spacing: 1px;                       /* 规格：1px，更舒展 */
    animation: zjb-fade-up .6s cubic-bezier(.22, .61, .36, 1) .16s both;
}

/* ---------- 4) 副标题 ---------- */
.zjb-hero .zjb-tagline {
    margin: 12px 0 0 0;                        /* 规格：mt-12px */
    color: var(--hero-sub) !important;         /* 规格：#94a3b8 */
    font-size: 15px;                           /* 规格：15px */
    font-weight: 400;                          /* 规格：400 */
    line-height: 1.6;
    letter-spacing: .01em;
    animation: zjb-fade-up .6s cubic-bezier(.22, .61, .36, 1) .24s both;
}

/* 尊重系统「减少动态效果」：关掉所有动画，直接显示最终状态 */
@media (prefers-reduced-motion: reduce) {
    .zjb-hero .zjb-badge,
    .zjb-hero .zjb-name,
    .zjb-hero .zjb-slogan,
    .zjb-hero .zjb-tagline,
    .zjb-hero .zjb-mark {
        animation: none !important;
        opacity: 1;
        transform: none;
    }
}

/* ============================================================
   3. Tab 标签栏：居中 + 选中项深蓝下划线
   ============================================================ */
/* 当前 DOM：div[role="tablist"] > div[data-testid="stTab"]
   旧版本：button[data-baseweb="tab"] —— 两套选择器都写上，避免版本差异 */
[data-testid="stTabs"] {width: 100%;}

[data-testid="stTabs"] [role="tablist"],
.stTabs [data-baseweb="tab-list"] {
    display: flex !important;
    flex-wrap: wrap !important;
    justify-content: center !important;
    gap: .3rem !important;
    padding: 0 !important;
    border-bottom: 1px solid var(--zjb-border) !important;
    background: transparent !important;
}

/* 单个标签：灰字、无底色、2px 透明底线预占位，避免选中时跳动 */
[data-testid="stTabs"] [data-testid="stTab"],
.stTabs button[data-baseweb="tab"] {
    display: flex !important;
    flex: 0 0 auto !important;
    align-items: center !important;
    justify-content: center !important;
    height: 3.1rem !important;
    padding: 0 1.05rem !important;
    /* 悬停/点击后的底色统一为全透明，只靠文字颜色表达状态 */
    background: transparent !important;
    border: 0 !important;
    border-bottom: 2px solid transparent !important;
    border-radius: 0 !important;
    color: var(--zjb-muted) !important;
    font-size: .95rem !important;
    font-weight: 500 !important;
    letter-spacing: .01em;
    cursor: pointer;
    user-select: none;
    transition: color .18s ease, border-color .18s ease;
}
[data-testid="stTabs"] [data-testid="stTab"] p,
.stTabs button[data-baseweb="tab"] p {
    font-size: .95rem !important;
    font-weight: 500 !important;
    color: inherit !important;
    white-space: nowrap;
    margin: 0 !important;
}
/* 悬停、键盘聚焦、点击后（react-aria 会保留 data-hovered / data-focused）都保持无底色；
   用 !important + 高特异性同时压住组件内置的 &[data-hovered] 规则 */
[data-testid="stTabs"] [data-testid="stTab"]:hover,
[data-testid="stTabs"] [data-testid="stTab"][data-hovered],
[data-testid="stTabs"] [data-testid="stTab"][data-focused],
[data-testid="stTabs"] [data-testid="stTab"][data-hovered][data-selected],
[data-testid="stTabs"] [data-testid="stTab"][data-focused][data-selected],
.stTabs button[data-baseweb="tab"]:hover,
.stTabs button[data-baseweb="tab"][aria-selected="true"]:hover {
    color: var(--zjb-navy) !important;
    background: transparent !important;
    box-shadow: none !important;
}
[data-testid="stTabs"] [data-testid="stTab"]:focus,
[data-testid="stTabs"] [data-testid="stTab"]:focus-visible {
    outline: none !important;
    box-shadow: none !important;
}

/* 选中态：深蓝文字 + 深蓝下划线 */
[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"],
.stTabs button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--zjb-navy) !important;
    background: transparent !important;
    border-bottom: 2px solid var(--zjb-navy) !important;
}
[data-testid="stTabs"] [data-testid="stTab"][aria-selected="true"] p,
.stTabs button[data-baseweb="tab"][aria-selected="true"] p {
    color: var(--zjb-navy) !important;
    font-weight: 700 !important;
}
/* 去掉内置滑动高亮条，我们用自己的下划线 */
.react-aria-SelectionIndicator,
[data-testid="stTabs"] .react-aria-SelectionIndicator,
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] {display: none !important;}

/* 内容面板与标签栏的间距 */
[data-testid="stTabs"] [data-testid="stTabPanel"] {padding: 1.4rem 0 0 0 !important;}


/* ============================================================
   4. 卡片容器：白底 + 浅边框 + 圆角 + 轻阴影
   ============================================================ */
[class*="st-key-card-"] {
    background: var(--zjb-surface) !important;
    border: 1px solid var(--zjb-border) !important;
    border-radius: var(--zjb-radius) !important;
    box-shadow: var(--zjb-shadow);
    padding: 1.6rem 1.7rem 1.5rem 1.7rem !important;
    margin-bottom: 1.15rem;
    transition: box-shadow .2s ease, border-color .2s ease;
}
[class*="st-key-card-"]:hover {
    box-shadow: var(--zjb-shadow-lift);
    border-color: #D8DEEA !important;
}
[class*="st-key-card-"] [data-testid="stVerticalBlock"] {gap: .7rem;}

/* 卡片内的分栏在窄屏自动堆叠 */
@media (max-width: 640px) {
    /* Hero 区的字号/间距已全部由 clamp() 自适应，这里不再重复覆盖 */
    /* PDF基础处理有 4 个二级标签，窄屏收紧内边距，避免换行过密 */
    [data-testid="stTabs"] [data-testid="stTab"] {padding: 0 .62rem !important;}
    [data-testid="stTabs"] [data-testid="stTab"] p {font-size: .86rem !important;}
}

/* ============================================================
   5. 魔法上传框：白底 + 浅蓝虚线 + 悬停微蓝 + 拖拽高亮
   ============================================================ */
/* 外层容器：去掉多余边距，让它撑满卡片 */
[data-testid="stFileUploader"] {width: 100%;}
[data-testid="stFileUploader"] > section,
[data-testid="stFileUploaderDropzone"] {
    background: #FFFFFF !important;
    border: 1.5px dashed #B6CBE8 !important;      /* 浅蓝虚线 */
    border-radius: 14px !important;
    padding: 2.1rem 1.4rem !important;           /* 更大的呼吸感 */
    min-height: 164px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: .75rem !important;
    cursor: pointer;
    transition: background .2s ease, border-color .2s ease, box-shadow .2s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
    background: #F7FAFF !important;              /* 悬停微蓝 */
    border-color: #6E9BE0 !important;
    box-shadow: 0 0 0 4px rgba(37, 99, 235, .06);
}
/* 拖拽文件经过时：实一点的高亮 */
[data-testid="stFileUploaderDropzone"]:focus-within {
    border-color: var(--zjb-accent) !important;
    box-shadow: 0 0 0 4px rgba(37, 99, 235, .10);
}
/* 上传框内的说明文字 */
[data-testid="stFileUploaderDropzoneInstructions"] > div > span,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: var(--zjb-muted) !important;
    font-size: .84rem !important;
    text-align: center;
}
/* ------------------------------------------------------------------
   拖拽区文案：隐藏前端硬编码的 “200MB per file”
   Streamlit 这句是前端硬编码的 `${size} per file • ${类型}`（参数改不了），
   这里把 “数值 + per file” 视觉隐藏，再用 ::after 补文案；
   ::after 的 content 由 _apply_dragzone_tip() 按当前页面动态注入
   （各页支持格式不同，不能在本文件里写死，否则 Excel 上传框会显示成 PDF）。
   真正生效的限制另有 server.maxUploadSize=100（见 .streamlit/config.toml）
   与后端 split_files() 校验（普通用户 20MB / 3 份，兑换码用户 100MB / 20 份）。
   ------------------------------------------------------------------ */
[data-testid="stFileUploaderDropzoneInstructions"] span > span:first-child {
    font-size: 0 !important;      /* 隐藏原本的 “200MB per file” */
    letter-spacing: -1em;
}
[data-testid="stFileUploaderDropzoneInstructions"] span > span:first-child::after {
    content: var(--zjb-dragzone-tip, "支持格式：PDF/JPG/PNG，单文件最大 20MB。");
    font-size: .84rem !important;
    letter-spacing: normal;
    color: var(--zjb-muted);
}
/* 万一内置结构变化，兜底：整块说明区不出现 per file 英文 */
[data-testid="stFileUploaderDropzoneInstructions"] {text-align: center;}

/* 上传框内的“Browse files”按钮：淡蓝胶囊 */
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploaderDropzone"] button:hover {
    background: var(--zjb-accent-soft) !important;
    color: var(--zjb-accent) !important;
    border: 1px solid #D3E3FB !important;
    border-radius: 999px !important;
    padding: .45rem 1.15rem !important;
    font-size: .86rem !important;
    font-weight: 600 !important;
    box-shadow: none !important;
}
[data-testid="stFileUploaderDropzone"] button:hover {
    background: #E3EDFE !important;
    border-color: #BFD6F8 !important;
}
/* 已上传文件的列表块 */
[data-testid="stFileUploader"] ul,
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"] {
    background: #FBFCFE !important;
    border: 1px solid var(--zjb-border) !important;
    border-radius: 12px !important;
}
[data-testid="stFileUploader"] label,
[data-testid="stWidgetLabel"] p {
    color: #374151 !important;
    font-size: .88rem !important;
    font-weight: 600 !important;
}

/* ============================================================
   6. 按钮：圆角矩形、主色、轻阴影、悬停抬起 + 变色
   ============================================================ */
.stButton > button,
.stDownloadButton > button,
[data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-primary"] {
    border-radius: 12px !important;
    font-size: .92rem !important;
    font-weight: 600 !important;
    letter-spacing: .01em;
    padding: .62rem 1.5rem !important;
    min-height: 2.75rem;
    transition: transform .16s ease, box-shadow .2s ease,
                background .2s ease, border-color .2s ease, color .2s ease;
}

/* 主按钮（type="primary"） */
.stButton > button[kind="primary"],
[data-testid="stBaseButton-primary"] {
    background: var(--zjb-navy) !important;
    color: #FFFFFF !important;
    border: 1px solid var(--zjb-navy) !important;
    box-shadow: 0 1px 2px rgba(11, 36, 71, .18), 0 6px 16px rgba(11, 36, 71, .16);
}
.stButton > button[kind="primary"]:hover,
[data-testid="stBaseButton-primary"]:hover {
    background: var(--zjb-navy-hover) !important;
    border-color: var(--zjb-navy-hover) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px) scale(1.015);
    box-shadow: 0 10px 22px rgba(11, 36, 71, .2);
}

/* ============================================================
   7. 复选框 / 单选框：整齐居中、深蓝选中
   ============================================================ */
[data-testid="stCheckbox"],
[data-testid="stRadio"] {
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
    justify-content: flex-start !important;
    width: 100% !important;
}
[data-testid="stCheckbox"] > [data-testid="stRadioGroup"],
[data-testid="stRadio"] > [data-testid="stRadioGroup"] {width: 100% !important;}
/* 复选框方框：圆角 + 深蓝选中态（新旧两套 DOM 都覆盖；当前界面未用到 st.checkbox） */
[data-testid="stCheckbox"] [data-baseweb="checkbox"] div[role="checkbox"],
[data-testid="stCheckbox"] [data-testid="stCheckboxRootElement"] {
    border-radius: 6px !important;
    border-color: #C7D0DD !important;
}
[data-testid="stCheckbox"] [data-baseweb="checkbox"] div[role="checkbox"][aria-checked="true"],
[data-testid="stCheckbox"] [data-selected] [data-testid="stCheckboxRootElement"] {
    background-color: var(--zjb-navy) !important;
    border-color: var(--zjb-navy) !important;
}
/* 单选项（压缩模式）横向居中排布：新版 DOM 是 [data-testid="stRadioGroup"] > div > label */
[data-testid="stRadio"] [role="radiogroup"],
[data-testid="stRadioGroup"] {
    display: flex !important;
    flex-wrap: wrap !important;
    justify-content: center !important;
    gap: .5rem 1.4rem !important;
}
[data-testid="stRadioGroup"] > div {
    display: flex !important;
    align-items: center !important;
    flex: 1 1 auto !important;
}
[data-testid="stRadioGroup"] > div > label {max-width: 100%;}
/* 单选/复选的标题单独占一行，选项组整体居中 */
[data-testid="stRadio"] > label[data-testid="stWidgetLabel"],
[data-testid="stCheckbox"] > label[data-testid="stWidgetLabel"] {
    display: block !important;
    width: 100% !important;
    margin: 0 0 .5rem 0 !important;
}
[data-testid="stRadioOption"] {
    display: inline-flex !important;
    align-items: center !important;
    gap: .55rem !important;
    cursor: pointer;
    white-space: nowrap;
}
[data-testid="stRadioOption"] p {
    color: #374151 !important;
    font-size: .9rem !important;
    line-height: 1.4 !important;
    margin: 0 !important;
}
[data-testid="stRadioOption"][data-selected] p {color: var(--zjb-navy) !important; font-weight: 600 !important;}
/* 单选圆环：未选灰色描边，选中深蓝圆环 + 白点 */
[data-testid="stRadioOption"] div:has(> div) {border-color: #C7D0DD !important;}
[data-testid="stRadioOption"][data-selected] [class]:has(> div) {
    background-color: transparent !important;
    border-color: var(--zjb-navy) !important;
    border-width: 6px !important;
}
[data-testid="stRadioOption"][data-selected] [class]:has(> div) > div {background-color: transparent !important;}
/* 卡片里的分栏内容垂直居中，避免左右高度不齐 */
[class*="st-key-card-"] [data-testid="stHorizontalBlock"] {align-items: flex-start;}

/* ============================================================
   8. 输入框 / 文本域 / 提示文字
   ============================================================ */
/* Streamlit 的输入框是“外层 div（画底/边框）+ 内部无样式 input”，
   两套 DOM（data-baseweb 旧版 / data-testid 新版）都覆盖 */
.stTextInput [data-baseweb="input"],
.stTextInput [data-testid="stTextInputRootElement"],
.stTextArea [data-baseweb="base-input"],
.stTextArea [data-testid="stTextAreaRootElement"],
div[data-baseweb="input"],
div[data-baseweb="base-input"] {
    background: #FFFFFF !important;
    border: 1px solid #DCE2EC !important;
    border-radius: 11px !important;
    box-shadow: inset 0 1px 1px rgba(16, 24, 40, .02) !important;
    transition: border-color .18s ease, box-shadow .18s ease;
    overflow: hidden;
}
.stTextInput [data-baseweb="input"]:focus-within,
.stTextInput [data-testid="stTextInputRootElement"]:focus-within,
.stTextArea [data-baseweb="base-input"]:focus-within,
.stTextArea [data-testid="stTextAreaRootElement"]:focus-within {
    border-color: var(--zjb-accent) !important;
    box-shadow: 0 0 0 4px rgba(37, 99, 235, .10) !important;
}
.stTextInput input,
.stTextArea textarea {
    background: transparent !important;
    border: 0 !important;
    border-radius: 11px !important;
    padding: .72rem .95rem !important;
    font-size: .93rem !important;
    color: var(--zjb-text) !important;
    box-shadow: none !important;
    outline: none !important;
}
.stTextInput input::placeholder,
.stTextArea textarea::placeholder {color: #A7B0BF !important;}

/* 灰字说明（st.caption）：统一小号浅灰 */
[data-testid="stCaptionContainer"] p,
.stCaption p {
    color: var(--zjb-faint) !important;
    font-size: .8rem !important;
    line-height: 1.6 !important;
}

/* 安全承诺小字：居中的浅灰胶囊 */
.zjb-privacy {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: .35rem;
    margin: .1rem 0 .35rem 0;
    color: #9AA3B2;
    font-size: .76rem;
    line-height: 1.6;
    text-align: center;
}

/* ============================================================
   8.1 改字规则表：浅色底 + 深色字，禁用右上角浮动工具条
   ============================================================ */
/* 改字规则区：白底浅框，一行一对输入框，任何时候都是浅色主题 */
.zjb-rules {
    background: #FFFFFF;
    border: 1px solid var(--zjb-border);
    border-radius: 14px;
    padding: .9rem 1rem .6rem 1rem;
    margin: .2rem 0 .5rem 0;
}
.zjb-rules [data-testid="stWidgetLabel"] p {
    color: #4B5563 !important;
    font-size: .8rem !important;
    font-weight: 600 !important;
}
.zjb-rule-arrow {
    text-align: center;
    color: var(--zjb-accent);
    font-size: 1.05rem;
    font-weight: 700;
    line-height: 2.75rem;
}
/* 行尾删除按钮：红字 + 透明底，hover 出现浅红底色（图标本身是 🗑️） */
.zjb-rules [data-testid="stColumn"]:last-child button {
    background: transparent !important;
    border: 1px solid transparent !important;
    color: #DC2626 !important;
    font-size: 1.05rem !important;
    padding: .1rem .3rem !important;
    min-height: 2.4rem;
    box-shadow: none !important;
}
.zjb-rules [data-testid="stColumn"]:last-child button:hover {
    background: #FEE2E2 !important;
    border-color: #FECACA !important;
    color: #B91C1C !important;
}
/* “💡 怎么用”：浅蓝底一行说明，比长段落更好扫读 */
.zjb-howto {
    background: #F5F9FF;
    border: 1px solid #DCE9FB;
    border-radius: 10px;
    padding: .6rem .85rem;
    margin: .3rem 0 .1rem 0;
    color: #41506B;
    font-size: .86rem;
    line-height: 1.65;
}
.zjb-howto strong { color: var(--zjb-accent); }

/* “＋ 添加一条替换规则”：轻量文字按钮 */
[data-testid="stButton"] button[kind="tertiary"] {
    color: var(--zjb-accent) !important;
    font-weight: 600 !important;
}

/* 兜底：若后续重新引入 st.data_editor / st.dataframe，
   统一改成浅色（白底深字），并隐藏右上角浮动工具条（统计 / 搜索 / 全屏 / 下载） */
[data-testid="stDataFrame"] {
    background: #FFFFFF !important;
    border: 1px solid var(--zjb-border) !important;
    border-radius: 12px !important;
    overflow: hidden;
}
[data-testid="stDataFrame"] [data-testid="stDataFrameGlideDataEditor"],
[data-testid="stDataFrame"] canvas {
    background: #FFFFFF !important;
    color: var(--zjb-text) !important;
}
/* 工具条：新版是 stElementToolbar，旧版挂在 element 的 overlay 容器里 */
[data-testid="stDataFrame"] [data-testid="stElementToolbar"],
[data-testid="stElementToolbar"],
[data-testid="stDataFrame"] .gdg-toolbar,
[data-testid="stDataFrame"] [class*="toolbar"] {
    display: none !important;
    visibility: hidden !important;
    opacity: 0 !important;
    pointer-events: none !important;
}
[data-testid="stDataFrame"] {padding-right: 0 !important;}

/* ============================================================
   9. 页面大标题
   ============================================================ */

/* 页面大标题：极简，无序号标签 */
.zjb-page-head {
    margin: .1rem 0 1.1rem 0;
}
.zjb-page-head h2 {
    margin: 0 0 .4rem 0;
    color: var(--zjb-text);
    font-size: 1.5rem; font-weight: 700; line-height: 1.35;
    letter-spacing: .01em;
}
.zjb-page-head p {
    margin: 0;
    max-width: 46rem;
    color: var(--zjb-muted); font-size: .92rem; line-height: 1.75;
}

/* 步骤卡片顶部：标题 + 一句话说明 */
.zjb-step-head {
    display: flex; align-items: flex-start; gap: .7rem;
    margin: 0 0 .55rem 0;
}
.zjb-step-head .zjb-step-text {min-width: 0;}
.zjb-step-head .zjb-step-text strong {
    display: block; color: var(--zjb-text); font-size: .98rem; font-weight: 700;
    line-height: 1.5;
}
.zjb-step-head .zjb-step-text span {
    display: block; color: var(--zjb-muted); font-size: .8rem; line-height: 1.6;
}

/* 提示字号小字：正文级黑灰，保证任何底色上都能看清 */
.zjb-tip-text {
    color: #374151;
    font-size: .86rem;
    line-height: 1.7;
    margin: -.35rem 0 .1rem .15rem;
    font-weight: 400;
}
@media (prefers-color-scheme: dark) {
    .zjb-tip-text {color: #E5E7EB;}
}

/* 提示框（st.info / st.warning / st.success / st.error）：
   全局 .stApp 的文字色会漏进组件内部，浅色模式下黄底配浅字「看不清」。
   这里在浅色模式把警示框文字钉成深色；深色模式保留组件默认的浅色字，避免反过来糊住 */
@media (prefers-color-scheme: light) {
    [data-testid="stAlert"],
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span,
    [data-testid="stAlert"] div,
    [data-testid="stAlert"] li,
    [data-testid="stAlertContainer"],
    [data-testid="stAlertContainer"] p,
    [data-testid="stAlertContainer"] span,
    [data-testid="stAlertContainer"] div {
        color: #1F2937 !important;
        font-weight: 500;
        line-height: 1.7;
    }
    [data-testid="stAlert"] a,
    [data-testid="stAlertContainer"] a {color: #1D4ED8 !important; text-decoration: underline;}
}

/* 预览图：细边框 + 圆角，避免图片直接“糊”在卡片上 */
.zjb-preview img {
    border: 1px solid var(--zjb-border) !important;
    border-radius: 12px !important;
    box-shadow: 0 6px 18px rgba(11, 36, 71, .06);
}

/* 页脚 */
.zjb-footer {
    margin-top: 2.4rem; padding-top: 1.6rem;
    border-top: 1px solid var(--zjb-border);
    text-align: center;
    color: var(--zjb-faint); font-size: .78rem; line-height: 1.9;
}
.zjb-footer strong {color: #6B7280; font-weight: 600;}
/* 侧边栏品牌：闪电 + 闪办 + 英文名 */
.zjb-side-brand {
    display: flex; align-items: baseline; gap: .4rem;
    padding: 0 0 .55rem 0; margin-bottom: .55rem;
    border-bottom: 1px solid var(--zjb-border);
}
.zjb-side-brand .mark {font-size: 1.05rem; line-height: 1;}
.zjb-side-brand .word {
    color: var(--zjb-navy); font-size: 1.06rem; font-weight: 700; letter-spacing: .08em;
}
.zjb-side-brand .alias {
    color: var(--zjb-faint); font-size: .7rem; letter-spacing: .08em;
}

/* 页脚品牌行：闪电 + 闪办，比版权行略深一点 */
.zjb-footer .zjb-footer-brand {
    display: inline-block; margin-bottom: .15rem;
    color: #6B7280; font-weight: 600; letter-spacing: .06em;
}

.stButton > button[kind="primary"]:active,
[data-testid="stBaseButton-primary"]:active {
    transform: translateY(0) scale(.995);
    box-shadow: 0 1px 2px rgba(11, 36, 71, .2);
}
.stButton > button[kind="primary"]:focus,
[data-testid="stBaseButton-primary"]:focus {
    outline: none !important;
    box-shadow: 0 0 0 4px rgba(37, 99, 235, .18) !important;
}

/* 次要按钮（如上传框里的按钮已单独处理，这里管页面上的普通按钮） */
.stButton > button:not([kind="primary"]) {
    background: #FFFFFF !important;
    color: var(--zjb-navy) !important;
    border: 1px solid #D5DCE7 !important;
    box-shadow: 0 1px 2px rgba(16, 24, 40, .05);
}
.stButton > button:not([kind="primary"]):hover {
    background: #F4F7FC !important;
    border-color: #B9C6DA !important;
    color: var(--zjb-navy) !important;
    transform: translateY(-1px);
    box-shadow: 0 6px 14px rgba(16, 24, 40, .08);
}

/* 下载按钮：深蓝描边胶囊，和主按钮区分但同一家族 */
.stDownloadButton > button {
    background: #FFFFFF !important;
    color: var(--zjb-navy) !important;
    border: 1.5px solid var(--zjb-navy) !important;
    box-shadow: 0 1px 2px rgba(11, 36, 71, .08), 0 6px 16px rgba(11, 36, 71, .08);
}
.stDownloadButton > button:hover {
    background: var(--zjb-navy) !important;
    color: #FFFFFF !important;
    transform: translateY(-1px) scale(1.015);
    box-shadow: 0 10px 22px rgba(11, 36, 71, .2);
}

</style>
"""


# --------------------------------------------------------------------------
# 纯逻辑：不依赖 Streamlit，方便直接测试
# --------------------------------------------------------------------------
@dataclass
class Outcome:
    """一次处理的统一结果，界面只认这个结构。"""

    ok: bool
    message: str
    data: bytes | None = None
    filename: str = ""
    mime: str = ""
    log: list[str] | None = None


def display_name(raw: str | None, fallback: str = "文件") -> str:
    """只保留文件名部分，避免上传名里带路径（如 ../x.pdf）。"""
    name = Path(raw or "").name.strip()
    return name or fallback


def with_suffix(name: str, suffix: str, *, new_ext: str | None = None) -> str:
    """报告.pdf + '-压缩' -> 报告-压缩.pdf；new_ext='.docx' 时换扩展名。"""
    path = Path(name)
    extension = new_ext or path.suffix or ".pdf"
    return f"{path.stem or '文件'}{suffix}{extension}"


def compress_one(data: bytes, filename: str, *, aggressive: bool = False) -> Outcome:
    """压缩一份 PDF：无损优先；选中“极致压缩”时才对照片图片按 1MB 目标降级。"""
    log: list[str] = []

    def note(message: str) -> None:
        log.append(message)

    try:
        with tempfile.TemporaryDirectory(prefix="zjb_compress_") as tmp:
            folder = Path(tmp)
            source = folder / display_name(filename, "upload.pdf")
            source.write_bytes(data)
            result = compress.compress_pdf(
                source,
                out_dir=folder,
                target_size=COMPRESS_TARGET if aggressive else None,
                overwrite=True,
                callback=note,
            )
            if not result.ok or result.output is None:
                return Outcome(False, result.message or "压缩失败", log=log)
            payload = result.output.read_bytes()
    except Exception as exc:  # noqa: BLE001 - 任何底层异常都转成界面提示
        log.append(traceback.format_exc())
        return Outcome(False, f"压缩失败：{exc}", log=log)

    return Outcome(
        True,
        result.describe(),
        data=payload,
        filename=with_suffix(display_name(filename, "upload.pdf"), "-压缩"),
        mime="application/pdf",
        log=log,
    )


def to_word(data: bytes, filename: str) -> Outcome:
    """PDF 转 Word（.docx），走 converter 的 pdf2docx + PyMuPDF 回退。"""
    log: list[str] = []

    def note(message: str) -> None:
        log.append(message)

    try:
        with tempfile.TemporaryDirectory(prefix="zjb_word_") as tmp:
            folder = Path(tmp)
            source = folder / display_name(filename, "upload.pdf")
            source.write_bytes(data)
            result = converter.convert_pdf(
                source,
                out_dir=folder,
                overwrite=True,
                callback=note,
                auto_download_fonts=False,
                extract_embedded_fonts=False,
            )
            if not result.ok or result.output is None:
                return Outcome(False, result.message or "转换失败", log=log)
            payload = result.output.read_bytes()
            engine = {"pdf2docx": "版式还原", "pymupdf-text": "文字模式"}.get(
                result.engine, result.engine or "默认引擎"
            )
    except Exception as exc:  # noqa: BLE001
        log.append(traceback.format_exc())
        return Outcome(False, f"转换失败：{exc}", log=log)

    return Outcome(
        True,
        f"转换完成（{engine}），共 {result.page_count} 页",
        data=payload,
        filename=with_suffix(display_name(filename, "upload.pdf"), "", new_ext=".docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        log=log,
    )


def merge_pdfs(files: list[tuple[str, bytes]]) -> Outcome:
    """按上传顺序合并多份 PDF：pypdf 优先（不改内部结构），缺失时用 PyMuPDF。"""
    log: list[str] = []
    readers: list[object] = []
    buffers: list[bytes] = []

    def add(message: str) -> None:
        log.append(message)

    import pymupdf

    for name, payload in files:
        try:
            reader = pymupdf.open(stream=payload, filetype="pdf")
        except Exception as exc:  # noqa: BLE001 - 单个坏文件不影响其它文件
            add(f"✘ {name}：无法读取（{exc}）")
            continue
        if reader.needs_pass:
            add(f"✘ {name}：已加密，请先解密")
            reader.close()
            continue
        readers.append(reader)
        buffers.append(payload)
        add(f"✔ {name}：{reader.page_count} 页")

    try:
        if not readers:
            return Outcome(False, "没有可合并的 PDF（文件可能已加密或已损坏）。", log=log)

        total_pages = sum(reader.page_count for reader in readers)  # type: ignore[attr-defined]
        if len(buffers) < 2:
            add("提示：至少需要 2 份 PDF 才有合并效果，已原样输出所选文件。")

        # insert_pdf 只搬运页面对象，不改写页面内容流，字体与版式保持原样。
        merged = pymupdf.open()
        try:
            for part in readers:
                merged.insert_pdf(part)  # type: ignore[arg-type]
            payload = merged.tobytes(garbage=3, clean=True)
        finally:
            merged.close()
        if not payload.startswith(b"%PDF"):
            raise RuntimeError("合并结果为空")
    except Exception as exc:  # noqa: BLE001
        add(traceback.format_exc())
        return Outcome(False, f"合并失败：{exc}", log=log)
    finally:
        for reader in readers:
            try:
                reader.close()  # type: ignore[attr-defined]
            except Exception:  # noqa: BLE001
                pass

    add("→ 合并结果.pdf")
    return Outcome(
        True,
        f"已合并 {len(buffers)} 份 PDF，共 {total_pages} 页",
        data=payload,
        filename="合并结果.pdf",
        mime="application/pdf",
        log=log,
    )


# --------------------------------------------------------------------------
# 文档转 PDF：Word / Excel / 图片 -> PDF（纯逻辑，不依赖 Streamlit）
# --------------------------------------------------------------------------
DOC_TO_PDF_TYPES = ["docx", "xlsx", "png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"]
DOC_TO_PDF_HELP = "支持 .docx / .xlsx / 图片（PNG · JPG · WEBP · BMP · TIFF）"

# 图片扩展名 -> 展示用格式名
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
IMAGE_FORMAT_NAMES = {
    ".png": "PNG",
    ".jpg": "JPG",
    ".jpeg": "JPEG",
    ".webp": "WEBP",
    ".bmp": "BMP",
    ".tif": "TIFF",
    ".tiff": "TIFF",
}

# A4 竖版与页边距（pt）：正文宽度决定排版可用的横向空间
DOC_PAGE_MARGIN_X = 56
DOC_PAGE_MARGIN_Y = 56
DOC_PAGE_WIDTH = 595
DOC_PAGE_HEIGHT = 842
DOC_CONTENT_WIDTH = DOC_PAGE_WIDTH - DOC_PAGE_MARGIN_X * 2   # 483pt

# w:br 的 type 属性（分页符）
_W_BR = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br"
_W_TYPE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type"


def html_escape(text: str) -> str:
    """转义 HTML 特殊字符（用于不带引号的 style 属性与正文）。"""
    return html_lib.escape(text, quote=False)


def docx_to_html(data: bytes) -> tuple[str, int]:
    """把 .docx 转成 HTML 片段，返回 (html, 块数量)。

    逐块读取：保留标题字号、加粗 / 斜体 / 下划线、段落对齐、表格结构；
    分页符（w:br type="page"）转成强制分页。
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = Document(BytesIO(data))
    align_map = {
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }

    def runs_html(paragraph: Any) -> str:
        parts: list[str] = []
        for run in paragraph.runs:
            text = html_escape(run.text)
            if not text:
                continue
            style = []
            if run.bold:
                style.append("font-weight:700")
            if run.italic:
                style.append("font-style:italic")
            if run.underline:
                style.append("text-decoration:underline")
            parts.append(f'<span style="{";".join(style)}">{text}</span>' if style else text)
        return "".join(parts)

    def paragraph_html(paragraph: Any) -> str:
        text = runs_html(paragraph)
        name = (paragraph.style.name or "").lower() if paragraph.style else ""
        if name.startswith("heading"):
            level = "".join(ch for ch in name if ch.isdigit()) or "1"
            size = {"1": 22, "2": 18, "3": 15}.get(level, 13)
            return f'<p style="font-size:{size}px;font-weight:700;margin:0 0 8px 0;">{text}</p>'
        align = align_map.get(paragraph.alignment)
        align_css = f"text-align:{align};" if align else ""
        return f'<p style="{align_css}margin:0 0 6px 0;">{text}&nbsp;</p>'

    def table_html(table: Any) -> str:
        rows: list[str] = []
        for row in table.rows:
            cells = "".join(
                '<td style="border:1px solid #9AA3B2;padding:4px 6px;">'
                f'{"<br>".join(html_escape(one.text) for one in cell.paragraphs)}</td>'
                for cell in row.cells
            )
            rows.append(f"<tr>{cells}</tr>")
        return (
            '<table style="border-collapse:collapse;width:100%;font-size:11px;'
            f'margin:0 0 8px 0;">{"".join(rows)}</table>'
        )

    blocks: list[str] = []
    for child in document.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            blocks.append(paragraph_html(Paragraph(child, document)))
            if any(br.get(_W_TYPE) == "page" for br in child.iter(_W_BR)):
                blocks.append('<div style="page-break-after:always;"></div>')
        elif tag == "tbl":
            blocks.append(table_html(Table(child, document)))

    if not blocks:
        blocks.append('<p style="margin:0;">（文档为空）</p>')
    return "".join(blocks), len(blocks)


def cell_text(value: Any) -> str:
    """Excel 单元格转文本：整数化的浮点去掉 .0，其余原样。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    """读取 xlsx 的共享字符串表（富文本取纯文字）。"""
    try:
        raw = archive.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(raw)
    return [
        "".join(node.text or "" for node in item.iter() if node.tag.endswith("}t"))
        for item in root
    ]


def xlsx_row_cells(row: Any, shared: list[str]) -> list[str]:
    """把一行 <row> 里的单元格转成文本列表（空单元格保留为 ''）。"""
    values: dict[int, str] = {}
    column = 0
    for cell in (node for node in row if node.tag.endswith("}c")):
        ref = cell.get("r") or ""
        letters = "".join(ch for ch in ref if ch.isalpha())
        if letters:  # 依据单元格引用还原列号，避免空单元格错位
            index = 0
            for ch in letters:
                index = index * 26 + (ord(ch.upper()) - 64)
            column = index - 1
        kind = cell.get("t")
        raw = next((n.text or "" for n in cell if n.tag.endswith("}v")), "")
        if kind == "s" and raw.isdigit() and int(raw) < len(shared):
            text = shared[int(raw)]
        elif kind == "inlineStr":
            text = "".join(n.text or "" for n in cell.iter() if n.tag.endswith("}t"))
        else:
            text = raw
        values[column] = text
        column += 1
    if not values:
        return []
    width = max(values) + 1
    return [values.get(index, "") for index in range(width)]


def sheet_table_html(title: str, rows: list[str]) -> str:
    """把一个工作表包成「标题 + 表格 + 分页」，保持阅读顺序清晰。"""
    body = "".join(rows) or '<tr><td style="color:#8A93A3;">（空工作表）</td></tr>'
    return (
        f'<p style="font-weight:700;margin:0 0 6px 0;">{html_escape(title)}</p>'
        '<table style="border-collapse:collapse;width:100%;font-size:10px;'
        f'margin:0 0 6px 0;">{body}</table>'
        '<div style="page-break-after:always;"></div>'
    )


def xlsx_to_html(data: bytes) -> tuple[str, int]:
    """把 .xlsx 转成 HTML 表格（按工作表分页），返回 (html, 工作表数)。

    xlsx 本质是 zip：直接读 workbooks.xml 拿到工作表名与顺序，再按 sheetN.xml
    逐行还原单元格文字。不引入 openpyxl，部署环境无需新增依赖。
    """
    try:
        archive = zipfile.ZipFile(BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ValueError("不是有效的 .xlsx 文件（可能已损坏或实为 .xls）") from exc

    with archive:
        shared = shared_strings(archive)
        names = archive.namelist()
        ordered: list[tuple[str, str]] = []
        if "xl/workbook.xml" in names:
            book = ET.fromstring(archive.read("xl/workbook.xml"))
            for node in book.iter():
                if not node.tag.endswith("}sheet"):
                    continue
                rel_id = next(
                    (v for key, v in node.attrib.items() if key.endswith("}id")), ""
                )
                ordered.append((node.get("name") or "Sheet", rel_id))

        rels: dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in names:
            rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            for node in rel_root:
                target = node.get("Target") or ""
                if target.startswith("/"):
                    target = target[1:]
                elif not target.startswith("xl/"):
                    target = f"xl/{target}"
                rels[node.get("Id") or ""] = target

        fallback_sheets = sorted(
            name for name in names if name.startswith("xl/worksheets/sheet")
        )
        targets: list[tuple[str, str]] = []
        for index, (title, rel_id) in enumerate(ordered):
            path = rels.get(rel_id, "")
            if path not in names:
                path = fallback_sheets[index] if index < len(fallback_sheets) else ""
            if path:
                targets.append((title, path))
        if not targets:
            targets = [(f"Sheet{i + 1}", path) for i, path in enumerate(fallback_sheets)]
        if not targets:
            raise ValueError("Excel 里没有找到可读取的工作表")

        sections: list[str] = []
        for title, path in targets:
            root = ET.fromstring(archive.read(path))
            rows: list[str] = []
            for row in (node for node in root.iter() if node.tag.endswith("}row")):
                cells = xlsx_row_cells(row, shared)
                if not cells or not any(cell.strip() for cell in cells):
                    continue  # 跳过整行空白
                rows.append(
                    "<tr>"
                    + "".join(
                        '<td style="border:1px solid #9AA3B2;padding:4px 6px;">'
                        f"{html_escape(cell)}</td>"
                        for cell in cells
                    )
                    + "</tr>"
                )
            sections.append(sheet_table_html(title, rows))

    return "".join(sections), len(sections)


def image_to_pdf(data: bytes, suffix: str) -> bytes:
    """单张图片转 PDF：页面按图片比例自适应，限制在 A4 大小内。"""
    try:
        pixmap = pymupdf.Pixmap(data)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(
            f"这张图没读出来（{IMAGE_FORMAT_NAMES.get(suffix, suffix)}）：{exc}"
        ) from exc
    if pixmap.colorspace is not None and pixmap.colorspace.n > 3:
        pixmap = pymupdf.Pixmap(pymupdf.csRGB, pixmap)
    png_bytes = pixmap.tobytes("png")

    # 96dpi -> pt，宽度先限制在 A4 之内，再按比例算高度
    width = min(max(pixmap.width * 72 / 96, 72), DOC_PAGE_WIDTH)
    height = width * pixmap.height / max(pixmap.width, 1)
    if height > DOC_PAGE_HEIGHT:  # 长图不生成超长页
        scale = DOC_PAGE_HEIGHT / height
        width, height = width * scale, DOC_PAGE_HEIGHT

    document = pymupdf.open()
    page = document.new_page(width=width, height=height)
    page.insert_image(pymupdf.Rect(0, 0, width, height), stream=png_bytes)
    output = document.tobytes(garbage=3, deflate=True)
    document.close()
    return output


def html_to_pdf(fragment: str, font_size: int = 12) -> bytes:
    """HTML 片段 -> 多页 A4 PDF。用 PyMuPDF Story 流式排版，超出自动翻页。"""
    template = pymupdf.open()
    template.new_page(width=DOC_PAGE_WIDTH, height=DOC_PAGE_HEIGHT)  # 用 A4 作页面模板
    sink = BytesIO()
    writer = pymupdf.DocumentWriter(sink)
    try:
        story = pymupdf.Story(
            html=(
                '<body style="font-family:sans-serif;'
                f'font-size:{font_size}px;line-height:1.65;color:#111827;">'
                f"{fragment}</body>"
            )
        )
        content = pymupdf.Rect(
            DOC_PAGE_MARGIN_X,
            DOC_PAGE_MARGIN_Y,
            DOC_PAGE_WIDTH - DOC_PAGE_MARGIN_X,
            DOC_PAGE_HEIGHT - DOC_PAGE_MARGIN_Y,
        )
        more = 1
        pages = 0
        while more:
            device = writer.begin_page(pymupdf.paper_rect("a4"))
            more, _filled = story.place(content)
            story.draw(device)
            writer.end_page()
            pages += 1
            if pages > 500:  # 防御：异常内容不至于无限翻页
                raise RuntimeError("文档页数过多（超过 500 页），请拆分后再转换")
        writer.close()
        output = sink.getvalue()
    finally:
        template.close()
    return output


def document_to_pdf(data: bytes, filename: str) -> Outcome:
    """Word / Excel / 图片 -> PDF。纯 Python 实现，不依赖 Office 或 LibreOffice。"""
    log: list[str] = []
    name = display_name(filename, "upload")
    suffix = Path(name).suffix.lower()
    if not data:
        return Outcome(False, "文件内容为空，请重新上传。", log=log)

    stem = name[: -len(suffix)] if suffix and name.endswith(suffix) else Path(name).stem
    stem = stem or "转换结果"

    try:
        if suffix in IMAGE_EXTENSIONS:
            pdf_bytes = image_to_pdf(data, suffix)
            log.append(f"✔ {name}：{IMAGE_FORMAT_NAMES.get(suffix, suffix)} 图片已嵌入 PDF")
            detail = f"{IMAGE_FORMAT_NAMES.get(suffix, suffix)} 图片已转换为 PDF"
        elif suffix == ".docx":
            fragment, count = docx_to_html(data)
            log.append(f"✔ {name}：读出 {count} 段文字 / 表格")
            pdf_bytes = html_to_pdf(fragment, font_size=12)
            detail = "Word 文档已转换为 PDF"
        elif suffix == ".xlsx":
            fragment, sheets = xlsx_to_html(data)
            log.append(f"✔ {name}：共 {sheets} 个工作表")
            pdf_bytes = html_to_pdf(fragment, font_size=10)
            detail = f"Excel 表格已转换为 PDF（{sheets} 个工作表）"
        elif suffix in (".doc", ".xls"):
            return Outcome(
                False,
                f"旧版 {suffix} 格式暂不支持，请另存为 {suffix}x 后再上传。",
                log=log,
            )
        else:
            return Outcome(
                False,
                f"暂不支持的文件类型：{suffix or '无扩展名'}。{DOC_TO_PDF_HELP}",
                log=log,
            )
    except Exception as exc:  # noqa: BLE001 - 任何底层异常都转成中文提示
        log.append(traceback.format_exc())
        return Outcome(False, f"转换失败：{exc}", log=log)

    if not pdf_bytes.startswith(b"%PDF"):
        return Outcome(False, "转换结果为空，请检查文件是否损坏。", log=log)

    log.append(f"→ {stem}.pdf（{compress.human_size(len(pdf_bytes))}）")
    return Outcome(
        True,
        detail,
        data=pdf_bytes,
        filename=f"{stem}.pdf",
        mime="application/pdf",
        log=log,
    )


def split_rule_lines(text: str) -> list[str]:
    """把多行文本拆成规则行：去掉空行与首尾空白，保留顺序。"""
    return [line.strip() for line in (text or "").splitlines() if line.strip()]


@dataclass
class Hit:
    """命中位置：页码（1 起）、该页命中总数、已完成几处。"""

    page: int
    total: int
    done: int

    @property
    def keep_last(self) -> int:
        """还要保留在该页后面的命中数，交给引擎只改一处。"""
        return max(0, self.total - self.done - 1)


def locate_next_hit(data: bytes, rule: "pdf_replace.ReplaceRule") -> tuple[Hit, str]:
    """找到下一次要替换的位置，返回 (命中, 错误提示)。

    按页用 PyMuPDF 搜索原文；同一页里已经替换了几处（搜索新文字的数量）就算作
    已完成，交给引擎保留这些位置之后的所有命中，从而“一次只改一处”，并从文档
    开头往后依次推进。
    """
    allowed = rule.page_list()  # None = 全部页
    scope = "全部页" if allowed is None else "、".join(str(page) for page in allowed) + " 页"
    try:
        with pymupdf.open(stream=data, filetype="pdf") as document:
            if document.needs_pass:
                return Hit(0, 0, 0), "该 PDF 已加密，暂不支持修改。"
            for page_no in range(1, document.page_count + 1):
                if allowed is not None and page_no not in allowed:
                    continue
                page = document[page_no - 1]
                hits = page.search_for(rule.old)
                if not hits:
                    continue
                done = len(page.search_for(rule.new)) if rule.new != rule.old else 0
                if done >= len(hits):
                    continue  # 这一页的命中都改完了
                return Hit(page_no, len(hits), done), ""
    except Exception as exc:  # noqa: BLE001
        return Hit(0, 0, 0), f"定位文字失败：{exc}"
    return Hit(0, 0, 0), f"在{scope}中没有找到需要替换的「{rule.old}」。"


def build_replace_rules(
    old_text: str,
    new_text: str,
    pages: str = "1",
) -> tuple[list[pdf_replace.ReplaceRule], str]:
    """把两个多行输入框配成替换规则列表，返回 (规则, 错误提示)。

    第 N 行旧文字 对应 第 N 行新文字；行数不一致时返回错误说明。
    每一组都必须等长，否则会改变排版。
    """
    olds, news = split_rule_lines(old_text), split_rule_lines(new_text)
    if not olds or not news:
        return [], "旧文字和新文字都不能为空。"
    if len(olds) != len(news):
        return [], (
            f"行数不一致：旧文字 {len(olds)} 行，新文字 {len(news)} 行。"
            "请按行一一对应填写（一行一条替换规则）。"
        )

    rules: list[pdf_replace.ReplaceRule] = []
    mismatched: list[str] = []
    for index, (old, new) in enumerate(zip(olds, news), start=1):
        if len(old) != len(new):
            mismatched.append(f"第 {index} 行 {len(old)}→{len(new)} 字（{old} → {new}）")
            continue
        rules.append(pdf_replace.ReplaceRule(old, new, pages))
    if mismatched:
        return [], (
            "必须等长替换，以下规则长度不同会导致错位：" + "；".join(mismatched)
            + "。请调整后再试。"
        )
    return rules, ""


def replace_text(
    data: bytes,
    filename: str,
    old_text: str,
    new_text: str,
    pages: str = "1",
    *,
    replace_all: bool = False,
) -> Outcome:
    """PDF/图片改字（支持多行多组规则）。

    replace_all=True 时把每个匹配位置都替换掉；否则每组规则只替换文档里第一次出现的位置。
    """
    pages = (pages or "").strip()
    name = display_name(filename, "upload.pdf")
    log: list[str] = []

    def note(message: str) -> None:
        log.append(message)

    rules, problem = build_replace_rules(old_text, new_text, pages)
    if problem:
        return Outcome(False, problem, log=log)
    try:
        for rule in rules:
            rule.page_list()  # 提前校验页码写法（1 / 1,3 / 2-5）
    except pdf_replace.RuleError as exc:
        return Outcome(False, str(exc), log=log)

    try:
        if replace_all:
            # 全部替换：一次调用，底层会把每个内容流里的命中全部替换掉
            with tempfile.TemporaryDirectory(prefix="zjb_replace_all_") as tmp:
                folder = Path(tmp)
                source = folder / name
                source.write_bytes(data)
                destination = folder / with_suffix(name, "_已修改")
                result = pdf_replace.replace_pdf_text(
                    source,
                    rules,
                    output_pdf_path=destination,
                    match_size=True,
                    overwrite=True,
                    log=note,
                )
                if not result.ok or result.output_pdf_path is None:
                    return Outcome(False, result.message or "替换失败", log=log)
                payload = Path(result.output_pdf_path).read_bytes()
            total = sum(repl.count for repl in (result.detail.replacements if result.detail else []))
            return Outcome(
                True,
                f"{result.describe()}（{len(rules)} 组规则，全文替换共 {total or '全部'} 处）",
                data=payload,
                filename=with_suffix(name, "_已修改"),
                mime="application/pdf",
                log=log,
            )

        # 只改一处：逐条规则串行处理，每条只提交一处命中，其余保持原样
        with tempfile.TemporaryDirectory(prefix="zjb_replace_one_") as tmp:
            folder = Path(tmp)
            current = data
            counts: list[int] = []
            for index, rule in enumerate(rules, start=1):
                hit, problem = locate_next_hit(current, rule)
                if problem:
                    return Outcome(False, problem, log=log)
                note(
                    f"[{index}/{len(rules)}] {rule.old} → {rule.new}"
                    f"（第 {hit.page} 页第 {hit.done + 1}/{hit.total} 处）"
                )
                step = folder / f"step{index}"
                step.mkdir()
                source = step / name
                source.write_bytes(current)
                destination = step / with_suffix(name, "_已修改")
                result = pdf_replace.replace_pdf_text(
                    source,
                    [pdf_replace.ReplaceRule(rule.old, rule.new, str(hit.page))],
                    output_pdf_path=destination,
                    match_size=True,
                    overwrite=True,
                    log=note,
                    keep_last=hit.keep_last,
                )
                if not result.ok or result.output_pdf_path is None:
                    return Outcome(False, result.message or "替换失败", log=log)
                current = Path(result.output_pdf_path).read_bytes()
                counts.append(1)

        return Outcome(
            True,
            f"已替换 {len(counts)} 处（{len(rules)} 组规则，每组仅替换第一次出现的位置）",
            data=current,
            filename=with_suffix(name, "_已修改"),
            mime="application/pdf",
            log=log,
        )
    except Exception as exc:  # noqa: BLE001
        log.append(traceback.format_exc())
        return Outcome(False, f"替换失败：{exc}", log=log)


# --------------------------------------------------------------------------
# 商业化分级：后端校验（前端提示只是展示，真正拦截在这里）
# --------------------------------------------------------------------------
# 商业化分级：后端校验（前端提示只是展示，真正拦截在这里）
# --------------------------------------------------------------------------
def _file_size(file: Any) -> int:
    """取上传文件字节数：优先 file.size，缺失时退回读 getvalue()。"""
    size = getattr(file, "size", None)
    if isinstance(size, int) and size >= 0:
        return size
    try:
        return len(file.getvalue())
    except Exception:  # noqa: BLE001 - 取不到大小按 0 处理，交给后续解析报错
        return 0


def limits_active() -> bool:
    """当前会话是否仍受基础额度限制（输入有效 Pro 兑换码后返回 False）。

    每次调用时重新校验 session 里的兑换码，而不是记一个“已解锁”布尔值：
    这样用户清掉兑换码后限制会立刻恢复，也不会被前端伪造成已解锁状态。
    """
    return st.session_state.get(PRO_CODE_KEY) not in PRO_CODES


def current_file_limit() -> int:
    """当前单次可处理的文件份数上限（基础 3 份 / Pro 20 份）。"""
    return BASE_MAX_FILES if limits_active() else PRO_MAX_FILES


def current_size_bytes() -> int:
    """当前单文件大小上限（基础 20MB / Pro 100MB）。"""
    return BASE_MAX_UPLOAD_BYTES if limits_active() else PRO_MAX_UPLOAD_MB * 1024 * 1024


def show_upgrade(error: str = LIMIT_MESSAGE, info: str = UNLOCK_MESSAGE) -> None:
    """触到额度上限时的统一提示：红色额度说明 + 蓝色解锁指引。

    只在真正越界时调用，普通浏览路径完全不出现，保证界面干净。
    """
    st.error(error)
    st.info(info)


def check_files(files: Sequence[Any]) -> str:
    """返回越界原因（""=全部合规）。

    份数或单文件大小任一超出基础额度即视为越界；Pro 兑换码生效时永远返回 ""。
    """
    items = [f for f in (files or []) if f is not None]
    if not items or not limits_active():
        return ""
    reasons: list[str] = []
    if len(items) > BASE_MAX_FILES:
        reasons.append(f"本次选择了 {len(items)} 份文件，已超出单次处理上限。")
    over = [f for f in items if _file_size(f) > BASE_MAX_UPLOAD_BYTES]
    if over:
        detail = "、".join(
            f"{display_name(getattr(f, 'name', ''), '文件')}"
            f"（{_file_size(f) / 1024 / 1024:.1f}MB）"
            for f in over
        )
        reasons.append(f"超出单文件大小的文件：{detail}。")
    return "\n\n".join(reasons)


def split_files(
    files: Sequence[Any],
    *,
    strict: bool = False,
    over_limit_message: str = "",
) -> list[Any] | None:
    """额度校验唯一入口：份数 + 单文件大小，越界时弹升级提示。

    参数：
        files:  本次要处理的文件（单个文件也会被包成列表）。
        strict: 上传框里用 True —— 越界时立即 st.stop()，上传框以下的处理逻辑不渲染；
                点击处理按钮时用 False —— 只返回 None，由调用方决定后续（通常什么都不做）。
        over_limit_message: 份数超限时替换红色提示正文（个别页面可用更贴近操作的措辞）；
                单文件超大时仍用统一的 LIMIT_MESSAGE。
    返回：
        合规文件列表（原样返回，可直接交给压缩 / 转换等处理函数）；
        越界时返回 None —— 与“空列表（还没上传文件）”区分开。
    """
    items = [f for f in (files or []) if f is not None]
    if not items:
        return []
    reason = check_files(items)
    if not reason:
        return items

    too_many = limits_active() and len(items) > BASE_MAX_FILES
    if over_limit_message and too_many:
        show_upgrade(over_limit_message)
    else:
        show_upgrade()
    st.caption(reason)
    st.caption("本次操作已中止，未上传、未处理、未保存任何内容。")
    if strict:
        st.stop()
    return None


def enforce_upload_limits(files: Sequence[Any], *, strict: bool = True) -> bool:
    """兼容旧调用点：返回是否合规（True = 可以继续处理）。"""
    return split_files(files, strict=strict) is not None


# --------------------------------------------------------------------------
# 界面公共组件
# --------------------------------------------------------------------------
def privacy_note() -> None:
    """全局安全承诺：放在每个上传组件下方。"""
    st.markdown(f'<div class="zjb-privacy">{PRIVACY_NOTE}</div>', unsafe_allow_html=True)


def _apply_pro_code(code: str) -> bool:
    """校验兑换码：有效则写入 session 并返回 True。"""
    normalized = (code or "").strip().upper()
    if normalized in PRO_CODES:
        st.session_state[PRO_CODE_KEY] = normalized
        return True
    return False


def render_sidebar() -> None:
    """左侧边栏：Pro 兑换码入口。

    平时收起、不占用主界面，只有触到额度上限的用户才会被提示过来输入；
    输入有效兑换码后，所有份数 / 大小限制立即失效，界面也不出现任何报错。
    """
    with st.sidebar:
        # 侧边栏品牌：闪电 + 闪办（点击不了，只做识别）
        st.markdown(
            '<div class="zjb-side-brand">'
            f'<span class="mark">{APP_ICON}</span>'
            f'<span class="word">{APP_NAME}</span>'
            '<span class="alias">SwiftDone</span>'
            "</div>",
            unsafe_allow_html=True,
        )
        st.markdown("#### 🎟️ Pro 兑换码")
        if not limits_active():
            st.success(f"已解锁 Pro 权益（{st.session_state.get(PRO_CODE_KEY)}）")
            st.caption(
                f"单次最多 {PRO_MAX_FILES} 份 · 单文件最大 {PRO_MAX_UPLOAD_MB}MB · "
                "全部功能本地处理。"
            )
            if st.button("退出 Pro（恢复默认额度）", key="pro_signout_btn", use_container_width=True):
                st.session_state.pop(PRO_CODE_KEY, None)
                st.rerun()
            return

        code = st.text_input(
            "输入兑换码",
            key="pro_code_input",
            placeholder="例如 VIP-TEST-001",
            help="输入后点击「解锁」，即刻生效。",
        )
        if st.button("解锁 Pro 权益", key="pro_unlock_btn", type="primary", use_container_width=True):
            if _apply_pro_code(code):
                st.rerun()
            else:
                st.warning("兑换码无效，请核对后再试。")
        st.caption("还没有兑换码？可联系客服获取。")


def page_head(title: str, desc: str = "") -> None:
    """页面大标题 + 一句话副标题（无序号、无小标签，排版保持极简）。"""
    st.markdown(
        f'<div class="zjb-page-head"><h2>{title}</h2>'
        f'<p>{desc}</p></div>',
        unsafe_allow_html=True,
    )


def card(key: str) -> Any:
    """白色卡片容器：白底 + 浅边框 + 圆角 + 轻阴影（样式由 CSS 按 st-key-card-* 命中）。"""
    return st.container(key=f"card-{key}")


# --------------------------------------------------------------------------
# 防连点：本地处理大文件会长时间占用 CPU，必须挡住重复提交
# --------------------------------------------------------------------------
BUSY_KEY = "zjb_busy"          # 处理任务 id（同一时刻只允许一个任务运行）


def is_busy() -> bool:
    """当前是否有任务正在处理（用于把其他按钮一起置灰）。"""
    return bool(st.session_state.get(BUSY_KEY))


def any_busy(tag: str) -> bool:
    """别的按钮正在处理时返回 True —— 用来联动置灰，避免并行任务把 CPU 打满。

    tag 是当前按钮自己的任务 id，所以“自己忙”不会把自己置灰（否则按钮会闪成
    disabled 却又可被再次点击，判断变得混乱）；is_busy() 才是全局判断。
    """
    current = st.session_state.get(BUSY_KEY)
    return bool(current) and current != tag


def run_task(tag: str, worker: Callable[[], Any]) -> Any:
    """在 busy 锁内执行一次处理任务，防连点、防重复提交。

    用法：`if st.button(..., disabled=any_busy("compress")): run_task("compress", do_it)`

    要点：
        1. 进入前先抢锁（抢不到说明已有任务在跑，直接返回 None，什么都不做）；
        2. 抢到锁后立刻占位渲染一个“正在处理…”提示，并**同步**写回 session_state：
           浏览器在这帧收到 disabled=True 之前，后续点击都会被这里挡住，
           不会重复触发 CPU 密集的文件处理；
        3. 处理完成后清锁（finally），按钮自动恢复可点、结果正常展示；
        4. 任何异常都保证清锁，不会把界面永久卡在 disabled 状态。

    注意：这里刻意不使用 st.rerun() —— 它在测试运行器（AppTest）里会中止本轮脚本，
    导致任务和结果都丢失；用占位提示同样能在浏览器端即时反馈“正在处理”。
    """
    if is_busy():
        return None
    st.session_state[BUSY_KEY] = tag
    st.caption("⏳ 正在处理，请稍候…（处理完成前按钮会保持不可用，避免重复提交）")
    try:
        return worker()
    finally:
        st.session_state.pop(BUSY_KEY, None)


def step_head(title: str, desc: str = "") -> None:
    """卡片顶部的“标题 + 说明”，用于三步式页面（不显示序号）。"""
    st.markdown(
        f'<div class="zjb-step-head">'
        f'<span class="zjb-step-text"><strong>{title}</strong>'
        f'<span>{desc}</span></span></div>',
        unsafe_allow_html=True,
    )


def preview_image(data: bytes, caption: str = "") -> None:
    """展示预览图（PNG 字节），带细边框。"""
    if not data:
        return
    st.markdown('<div class="zjb-preview">', unsafe_allow_html=True)
    st.image(data, caption=caption or None, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


def parse_editor_rules(frame: Any) -> tuple[list[tuple[str, str]], str]:
    """把改字规则的二维数据解析成 [(旧文字, 新文字), ...]。

    空行（两列都空）直接忽略；只有旧文字、没有新文字视为「删除这段文字」。
    返回 (规则列表, 错误说明)，有错误时规则为空。
    """
    rules: list[tuple[str, str]] = []
    if frame is None or len(frame) == 0:
        return [], "至少要填一条：左边写原来的字，右边写新字。"
    for number, row in enumerate(frame.itertuples(index=False), start=1):
        old = str(getattr(row, "旧文字", "") or "").strip()
        new = str(getattr(row, "新文字", "") or "").strip()
        if not old and not new:
            continue
        if not old:
            return [], f"第 {number} 行只填了新内容，请补上左边「原内容」。"
        rules.append((old, new))
    if not rules:
        return [], "至少要填一条：左边写原来的字，右边写新字。"
    return rules, ""


def render_header() -> None:
    """首屏 Hero：信任标签 + 品牌名 + 主 Slogan + 副标题（不放任何额度对比字样）。

    结构说明：
      <header>   首屏语义容器，内含唯一的 <h1>（整页 SEO / 读屏焦点）
      <p>        品牌名以外的三行都是段落，用 <p> 比 <div> 语义更准确；
                 品牌名放在 <h1> 里，图标用 aria-hidden 避免读屏念出「闪电」
      class 命名  沿用项目已有的 zjb-* 前缀，避免与 Streamlit 内部类名撞车
    """
    st.markdown(
        '<header class="zjb-hero">'
        '<p class="zjb-badge"><span class="dot" aria-hidden="true"></span>'
        "本地处理 · 无需注册 · 文件不留存</p>"
        "<h1 class=\"zjb-name\">"
        f"{APP_BOLT_SVG}"
        f'<span class="zjb-word">{APP_NAME}</span>'
        "</h1>"
        f'<p class="zjb-slogan">{APP_TAGLINE}</p>'
        f'<p class="zjb-tagline">{APP_SUBTITLE}</p>'
        "</header>",
        unsafe_allow_html=True,
    )


def show_result(outcome: Outcome, key: str) -> None:
    """统一收尾：成功 -> 指标 + 下载按钮；失败 -> 错误 + 可展开日志。"""
    if not outcome.ok:
        st.error(outcome.message)
        if outcome.log:
            with st.expander("查看详细日志"):
                st.code("\n".join(outcome.log), language="text")
        return

    if outcome.data is None:
        st.warning(outcome.message)
        return

    st.success(outcome.message)
    left, right = st.columns([2, 1], vertical_alignment="center")
    with left:
        st.caption(f"输出文件：{outcome.filename} · {compress.human_size(len(outcome.data))}")
    with right:
        st.download_button(
            "⬇ 下载处理结果",
            data=outcome.data,
            file_name=outcome.filename,
            mime=outcome.mime or "application/octet-stream",
            type="primary",
            key=f"download_{key}",
            use_container_width=True,
        )
    if outcome.log:
        with st.expander("处理日志"):
            st.code("\n".join(outcome.log), language="text")


@st.cache_data(show_spinner=False)
def _sample_bytes(kind: str) -> bytes:
    """示例文件字节（缓存一次，避免每次 rerun 重新生成）。

    batch_fill 里对应的是 sample_table_xlsx() / sample_template_pdf() 两个真实生成函数，
    分别产出标准 .xlsx 与 .pdf；这里做一层兜底，任何异常都不再抛出到页面：
        1. 生成函数缺失 -> 返回 b""，界面显示中文提示；
        2. Excel 生成异常 -> 退回带 BOM 的 CSV 兜底（同样能下载）；
        3. 格式校验失败（不是 zip / 不是 PDF）-> 同样视为生成失败。
    """
    try:
        if kind == "excel":
            payload = batch_fill.sample_table_xlsx()
            if payload and zipfile.is_zipfile(BytesIO(payload)):
                return payload
            # 极端兜底：pandas / openpyxl 全挂时，至少给一份能打开的 CSV
            return batch_fill.sample_table_csv()
        payload = batch_fill.sample_template_pdf()
        return payload if payload.startswith(b"%PDF") else b""
    except Exception:  # noqa: BLE001 - 示例下载失败不能影响整页
        return b""


def sample_download(kind: str, filename: str, mime: str, key: str) -> None:
    """“下载示例文件”按钮：有真实文件才渲染按钮，生成失败时给出中文提示而不是空按钮。"""
    payload = _sample_bytes(kind)
    label = "📥 下载 Excel 示例" if kind == "excel" else "📥 下载 PDF 模板示例"
    if not payload:
        st.button(
            label,
            key=f"{key}_disabled",
            disabled=True,
            use_container_width=True,
            help="示例文件生成失败，请直接按页面提示自行准备文件。",
        )
        st.caption("⚠️ 示例文件暂时给不了，照着下面两行自己建一个也很快：第一行写表头，下面每行写一个人的信息。")
        return
    st.download_button(
        label,
        data=payload,
        file_name=filename,
        mime=mime,
        key=key,
        use_container_width=True,
    )


def coming_soon(feature: str) -> Outcome:
    """占位功能：页面可正常切换，点击后给出明确的「开发中」提示。"""
    return Outcome(
        False,
        f"{feature} 正在开发中，敬请期待。其余功能不受影响，可继续使用。",
        log=[f"{feature}：正在开发中"],
    )


def _strip_quota_phrases(text: str) -> str:
    """去掉文案里的大小 / 份数额度表述，只保留格式说明。

    UPLOADER_CAPTION 已统一给出“普通 20MB / Pro 100MB / 20 份”，
    各页 caption_tip 里再写一遍 “单文件最大 20MB” 就成了两种口径并列。
    例：'支持格式：PDF，单文件最大 20MB。' -> '支持格式：PDF'
    """
    cleaned = re.sub(
        rf"[，,；;\s]*(单文件最大|免费版限|单次最多)\s*\d+\s*(MB|份)",
        "",
        text or "",
        flags=re.IGNORECASE,
    )
    return cleaned.strip().rstrip("。.；;，, ")


def _apply_dragzone_tip(tip: str) -> None:
    """给当前页面后续渲染的上传框设置拖拽区说明（CSS 变量，空串则沿用默认）。

    必须在 file_uploader() 之后调用：Streamlit 只把 style 插入到后续元素之前，
    CSS 变量按 DOM 顺序继承，所以“先控件、后注入”才能命中它自己的拖拽区。
    """
    if not tip:
        return
    text = json.dumps(tip, ensure_ascii=False)
    st.markdown(
        f'<style>:root {{ --zjb-dragzone-tip: {text}; }}</style>',
        unsafe_allow_html=True,
    )


def _render_uploader(
    label: str,
    key: str,
    types: list[str],
    *,
    multiple: bool = False,
    help_text: str = "",
    caption_tip: str = "",
    dragzone_tip: str = "",
    over_limit_message: str = "",
) -> list[object]:
    """上传框统一实现：统一 label + 统一提示 + 额度校验。

    - label 不采用调用方文案，全站统一为 “上传文件（普通用户限20MB）”，
      避免各页各写一套（“上传需要压缩的 PDF”等）导致口径不一致；
    - 格式说明合进同一条灰字 caption，末尾附隐私承诺；
    - max_upload_size 固定传 20MB，把组件内那行 “100MB per file” 改成 20MB；
    - 只有真的越界（超过 3 份或 20MB）时，split_files() 才弹升级提示，
      并且页面从上传框往下都不渲染（st.stop），不会执行任何处理逻辑。

    dragzone_tip：拖拽区内那行说明的后半段（支持格式）。Streamlit 把
    “N per file” 写死在前端，但数值可经 max_upload_size 控制；格式部分仍按页注入。
    """
    uploaded = st.file_uploader(
        UPLOADER_LABEL,
        type=types,
        accept_multiple_files=multiple,
        key=key,
        help=help_text or None,
        max_upload_size=UPLOADER_WIDGET_MAX_MB,
    )
    _apply_dragzone_tip(dragzone_tip)
    # caption 只保留“支持格式”：大小/份数口径由 UPLOADER_CAPTION 统一给出，
    # 这里把调用方文案里的 “单文件最大 20MB / 限 3 份” 去掉，避免同一段提示重复两遍。
    fmt = _strip_quota_phrases(caption_tip or "")
    parts = [UPLOADER_CAPTION.rstrip("。.；;，, ")]
    if fmt:
        parts.append(fmt)
    parts.append(PRIVACY_NOTE)
    st.caption("；".join(parts))
    if uploaded is None:
        return []
    # 注意：AppTest 会传出 UploadedFile 的 list 子类，必须按 list 处理，
    # 不能把整个列表当成“一个文件”包成 [list]，否则会被误判成超大文件。
    files = list(uploaded) if isinstance(uploaded, list) else [uploaded]
    approved = split_files(files, strict=True, over_limit_message=over_limit_message)
    return approved or []


def pdf_uploader(label: str, key: str, *, multiple: bool = False) -> list[object]:
    """PDF 上传框，返回合规文件列表（越界时提示并中止）。"""
    return _render_uploader(
        label,
        key,
        ["pdf"],
        multiple=multiple,
        help_text="可多选" if multiple else "仅支持 PDF",
        caption_tip=f"支持格式：PDF，单文件最大 {BASE_MAX_UPLOAD_MB}MB。",
        dragzone_tip=DRAGZONE_TIP_PDF,
    )


def uploader(
    label: str,
    key: str,
    types: list[str],
    *,
    multiple: bool = False,
    help_text: str = "",
    caption_tip: str = "",
    dragzone_tip: str = "",
    over_limit_message: str = "",
) -> list[object]:
    """通用上传框，返回合规文件列表（越界时提示并中止）。"""
    return _render_uploader(
        label,
        key,
        types,
        multiple=multiple,
        help_text=help_text,
        caption_tip=caption_tip,
        dragzone_tip=dragzone_tip,
        over_limit_message=over_limit_message,
    )


# --------------------------------------------------------------------------
# 板块一：PDF基础处理（免费）
# --------------------------------------------------------------------------
def render_basic() -> None:
    page_head(
        "PDF基础处理",
        "极速格式转换，免安装无广告。文件处理完即刻销毁，保护您的隐私。",
    )

    (
        tab_compress,
        tab_word,
        tab_merge,
        tab_pdf,
        tab_excel,
        tab_ppt,
        tab_split,
    ) = st.tabs(
        [
            "PDF 压缩",
            "PDF 转 Word",
            "PDF 合并",
            "文档转 PDF",
            "PDF 转 Excel",
            "PDF 转 PPT",
            "PDF 拆分",
        ]
    )

    # ---- PDF 压缩 ----
    with tab_compress:
        with card("compress"):
            files = pdf_uploader("上传需要压缩的 PDF", "basic_compress_upload")
            mode = st.radio(
                "压缩模式",
                list(COMPRESS_LEVELS),
                horizontal=True,
                key="basic_compress_mode",
            )
            if st.button(
                "开始压缩",
                type="primary",
                key="basic_compress_btn",
                use_container_width=True,
                disabled=any_busy("compress"),
            ):
                if not files:
                    st.warning("请先上传一个 PDF 文件。")
                else:
                    file = files[0]

                    def _do_compress() -> None:
                        with st.spinner("正在无损压缩…"):
                            outcome = compress_one(
                                file.getvalue(),
                                file.name,
                                aggressive=COMPRESS_LEVELS[mode],
                            )
                        show_result(outcome, "basic_compress")

                    run_task("compress", _do_compress)

    # ---- PDF 转 Word ----
    with tab_word:
        with card("word"):
            files = pdf_uploader("上传需要转换的 PDF", "basic_word_upload")
            st.caption("转出来的 Word 和原来长得一样；如果是照片或扫描件，暂时认不出上面的字。")
            if st.button(
                "开始转换",
                type="primary",
                key="basic_word_btn",
                use_container_width=True,
                disabled=any_busy("word"),
            ):
                if not files:
                    st.warning("请先上传一个 PDF 文件。")
                else:
                    file = files[0]

                    def _do_word() -> None:
                        with st.spinner("正在按原样转成 Word…"):
                            outcome = to_word(file.getvalue(), file.name)
                        show_result(outcome, "basic_word")

                    run_task("word", _do_word)

    # ---- PDF 转 Excel（占位） ----
    with tab_excel:
        with card("to_excel"):
            files = pdf_uploader("上传需要转换的 PDF", "basic_to_excel_upload")
            st.caption("把 PDF 里的表格变成 Excel；照片或扫描件暂时还认不出来。")
            if st.button(
                "转换为 Excel", type="primary", key="basic_to_excel_btn", use_container_width=True
            ):
                if not files:
                    st.warning("请先上传一个 PDF 文件。")
                else:
                    show_result(coming_soon("PDF 转 Excel"), "basic_to_excel")

    # ---- PDF 转 PPT（占位） ----
    with tab_ppt:
        with card("to_ppt"):
            files = pdf_uploader("上传需要转换的 PDF", "basic_to_ppt_upload")
            st.caption("按页拆成 PPT 幻灯片，一页 PDF 对应一页幻灯片。")
            if st.button(
                "转换为 PPT", type="primary", key="basic_to_ppt_btn", use_container_width=True
            ):
                if not files:
                    st.warning("请先上传一个 PDF 文件。")
                else:
                    show_result(coming_soon("PDF 转 PPT"), "basic_to_ppt")

    # ---- PDF 拆分（占位） ----
    with tab_split:
        with card("split"):
            files = pdf_uploader("上传需要拆分的 PDF", "basic_split_upload")
            st.text_input(
                "拆分方式",
                value="每页拆成一个文件",
                key="basic_split_mode",
                help="也可以填写页码范围，例如 1-3,4-6，每个逗号分隔的区间输出一份。",
            )
            st.caption("输出为 ZIP 压缩包，解压后得到拆分好的 PDF。")
            if st.button(
                "开始拆分", type="primary", key="basic_split_btn", use_container_width=True
            ):
                if not files:
                    st.warning("请先上传一个 PDF 文件。")
                else:
                    show_result(coming_soon("PDF 拆分"), "basic_split")

    # ---- PDF 合并 ----
    with tab_merge:
        with card("merge"):
            files = pdf_uploader("按顺序选择多份 PDF", "basic_merge_upload", multiple=True)
            st.caption(MERGE_HELP + "；输出文件名为「合并结果.pdf」。")
            if st.button(
                "开始合并",
                type="primary",
                key="basic_merge_btn",
                use_container_width=True,
                disabled=any_busy("merge"),
            ):
                if not files:
                    st.warning("请先上传至少一份 PDF 文件。")
                else:

                    def _do_merge() -> None:
                        with st.spinner("正在合并…"):
                            outcome = merge_pdfs([(f.name, f.getvalue()) for f in files])
                        show_result(outcome, "basic_merge")

                    run_task("merge", _do_merge)

    # ---- 文档转 PDF ----
    with tab_pdf:
        with card("to_pdf"):
            files = uploader(
                "上传 Word / Excel / 图片",
                "basic_to_pdf_upload",
                DOC_TO_PDF_TYPES,
                help_text=DOC_TO_PDF_HELP,
                dragzone_tip=DRAGZONE_TIP_DOC,
            )
            st.caption(DOC_TO_PDF_HELP + "；一次转换一份，输出标准 A4 PDF。")
            if st.button(
                "转换为 PDF",
                type="primary",
                key="basic_to_pdf_btn",
                use_container_width=True,
                disabled=any_busy("to_pdf"),
            ):
                if not files:
                    st.warning("请先上传一个 Word、Excel 或图片文件。")
                else:
                    file = files[0]

                    def _do_to_pdf() -> None:
                        size_mb = len(file.getvalue()) / 1024 / 1024
                        with st.spinner("正在排版并生成 PDF…"):
                            outcome = document_to_pdf(file.getvalue(), file.name)
                        if outcome.ok:
                            outcome.message = f"{outcome.message}（源文件 {size_mb:.1f}MB）"
                        show_result(outcome, "basic_to_pdf")

                    run_task("to_pdf", _do_to_pdf)


# --------------------------------------------------------------------------
# 板块二：PDF/图片改字（PDF 等长替换 / 图片擦除重绘）
# --------------------------------------------------------------------------
IMAGE_TYPES = ["pdf", "jpg", "jpeg", "png"]
IMAGE_TIP = (
    "💡 小提示：字底下的背景最好是纯色。如果背景有花纹或图案，"
    "改完可能有一点点色差，先看预览再下载。"
)
# 改字规则：默认 2 行示范格式，用户可按需增删单行（最多 MAX_RULE_ROWS 行）
RULE_DEFAULTS: list[tuple[str, str]] = [
    ("2024", "2025"),
    ("旧公司名称", "新公司名称"),
]
MAX_RULE_ROWS = 5
RULE_LIMIT_MESSAGE = f"最多只能加 {MAX_RULE_ROWS} 条规则，再多建议分两次改。"
RULE_SEED_KEY = "replace_rule_seed"   # 自增序号：保证新增行的 key 永不与旧行重复
RULE_STATE_KEY = "replace_rule_rows"  # 当前存在的行序号列表（可增可删）


def rules_to_frame(rules: Sequence[tuple[str, str]]) -> pd.DataFrame:
    """把 [(旧文字, 新文字), ...] 转成和 parse_editor_rules 对接的 DataFrame。"""
    return pd.DataFrame(
        [{"旧文字": old, "新文字": new} for old, new in rules],
        columns=["旧文字", "新文字"],
    )


def _rule_rows() -> list[int]:
    """当前要渲染的规则行 id 列表（session_state 持有，可增可删）。

    用自增 id 而不是“行数”作为 key：删除中间某行后，剩余行的输入值
    仍跟着自己的 key 走，不会串行；新增行也不会复用已删除的 key。
    """
    ids = st.session_state.get(RULE_STATE_KEY)
    if not isinstance(ids, list) or not ids:
        ids = list(range(len(RULE_DEFAULTS)))
        st.session_state[RULE_STATE_KEY] = ids
        st.session_state[RULE_SEED_KEY] = len(ids)
    return ids


def _add_rule_row() -> None:
    """追加一条空白规则（受 MAX_RULE_ROWS 限制）。"""
    ids = _rule_rows()
    if len(ids) >= MAX_RULE_ROWS:
        return
    seed = int(st.session_state.get(RULE_SEED_KEY, len(ids)))
    st.session_state[RULE_SEED_KEY] = seed + 1
    st.session_state[RULE_STATE_KEY] = [*ids, seed]


def _remove_rule_row(row_id: int) -> None:
    """删除指定行；至少保留 1 行，避免界面变成空的。"""
    ids = _rule_rows()
    remaining = [i for i in ids if i != row_id]
    if not remaining:
        remaining = [row_id]
    st.session_state[RULE_STATE_KEY] = remaining


def render_replace() -> None:
    page_head(
        "PDF/图片改字",
        "把文件里的字换掉：左边写原来的字，右边写新字，点一下就好。"
        "排版会自动对齐，改完可以先预览。",
    )

    with card("replace"):
        files = uploader(
            "上传需要修改的 PDF 或图片",
            "replace_upload",
            IMAGE_TYPES,
            help_text="支持 .pdf / .jpg / .jpeg / .png，单个文件处理",
            dragzone_tip=DRAGZONE_TIP_IMAGE_PDF,
        )
        file = files[0] if files else None
        is_image = bool(file and not file.name.lower().endswith(".pdf"))

        # 改字规则表：这里刻意不用 st.data_editor——新版 Streamlit 的单元格是 canvas 绘制，
        # 主题参数（theme="light"）在 1.66 之前并不存在，浅色 CSS 变量也覆盖不到画布像素，
        # 深色环境下整块表格会变成黑底，右上角还会浮出带统计图/搜索的工具栏挡住文字。
        # 改成「每行一对输入框 + 行尾删除按钮」：纯 HTML 输入控件，白底深字永远一致。
        st.markdown('<div class="zjb-rules">', unsafe_allow_html=True)
        st.caption("照着填就行：左边写要被换掉的字，右边写你想换成的新字。")
        rows: list[tuple[str, str]] = []
        for index, row_id in enumerate(_rule_rows()):
            old_default, new_default = (
                RULE_DEFAULTS[index] if index < len(RULE_DEFAULTS) else ("", "")
            )
            col_old, col_arrow, col_new, col_del = st.columns(
                [5, 0.7, 5, 1], vertical_alignment="center"
            )
            with col_old:
                old_value = st.text_input(
                    "原内容（想改掉的字）",
                    value=old_default,
                    key=f"replace_old_{row_id}",
                    placeholder="例如：2024年",
                    label_visibility="visible" if index == 0 else "collapsed",
                )
            with col_arrow:
                st.markdown('<div class="zjb-rule-arrow">→</div>', unsafe_allow_html=True)
            with col_new:
                new_value = st.text_input(
                    "新内容（替换成的字）",
                    value=new_default,
                    key=f"replace_new_{row_id}",
                    placeholder="例如：2025年",
                    label_visibility="visible" if index == 0 else "collapsed",
                )
            with col_del:
                if st.button(
                    "🗑️",
                    key=f"replace_del_{row_id}",
                    help="删掉这一条",
                    type="tertiary",
                ):
                    _remove_rule_row(row_id)
                    st.rerun()
            rows.append((old_value, new_value))
        st.markdown("</div>", unsafe_allow_html=True)

        left, right = st.columns([1, 1], vertical_alignment="center")
        with left:
            if st.button(
                "＋ 添加一条替换规则",
                key="replace_rule_add",
                type="tertiary",
                disabled=len(_rule_rows()) >= MAX_RULE_ROWS,
            ):
                _add_rule_row()
                st.rerun()
        with right:
            if len(_rule_rows()) >= MAX_RULE_ROWS:
                st.caption(RULE_LIMIT_MESSAGE)

        rules_frame = rules_to_frame(rows)
        rules, problem = parse_editor_rules(rules_frame)
        filled = len(rules)
        if filled:
            st.caption(f"已填 {filled} 条替换规则。")

        replace_all = st.checkbox(
            "✅ 把全篇的这个字都改掉",
            value=True,
            key="replace_all",
            help="勾上：全文每一处都改；取消：只改第一处。",
        )

        region: tuple[float, float, float, float] | None = None
        if is_image and file is not None:
            try:
                info = image_edit.inspect_image(file.getvalue())
            except image_edit.ImageEditError as exc:
                st.error(str(exc))
                return
            preview_image(info.preview_png, f"{info.width} × {info.height} 像素 · 原图")
            st.caption(IMAGE_TIP)
        else:
            st.text_input(
                "修改哪一页？（不填代表改全部页）",
                value="",
                key="replace_pages",
                placeholder="不填就改全部页；只改某几页可填 1,3 或 2-5",
            )

        if problem:
            st.caption(f"⚠️ {problem}")
        elif replace_all:
            st.caption("✅ 全篇出现的地方都会改。")
        else:
            st.caption("只改每处第一次出现的位置。")

        if is_image and file is not None:
            st.caption(f"ℹ️ {info.note}")

        st.markdown(
            '<div class="zjb-howto">💡 <strong>怎么用：</strong>'
            "左边填原话，右边填新话。如果右边留空，就是直接删掉这几个字。</div>",
            unsafe_allow_html=True,
        )

        if st.button(
            "开始改字",
            type="primary",
            key="replace_btn",
            use_container_width=True,
            disabled=any_busy("replace"),
        ):
            if not file:
                st.warning("请先上传一个 PDF 或图片文件。")
            elif problem:
                st.error(problem)
            elif is_image:

                def _do_replace_image() -> None:
                    with st.spinner("正在擦除旧文字并重绘新文字…"):
                        result = image_edit.replace_image_text(
                            file.getvalue(), region, rules, replace_all=replace_all
                        )
                    if not result.ok:
                        st.error(result.message)
                    else:
                        st.success(result.message)
                        for warning in result.warnings:
                            st.caption(f"ℹ️ {warning}")
                        preview_image(result.preview_png, "修改后预览")
                        st.download_button(
                            "⬇ 下载修改后的图片",
                            data=result.output or b"",
                            file_name=with_suffix(display_name(file.name, "图片.png"),
                                                  "_已修改", new_ext=".png"),
                            mime="image/png",
                            type="primary",
                            key="replace_img_download",
                            use_container_width=True,
                        )

                run_task("replace", _do_replace_image)
            else:

                def _do_replace_pdf() -> None:
                    outcome = replace_text(
                        file.getvalue(),
                        file.name,
                        "\n".join(old for old, _ in rules),
                        "\n".join(new for _, new in rules),
                        st.session_state.get("replace_pages", ""),
                        replace_all=replace_all,
                    )
                    show_result(outcome, "replace")

                run_task("replace", _do_replace_pdf)


# --------------------------------------------------------------------------
# 板块三：批量生成（Excel + 模板 PDF -> ZIP）
# --------------------------------------------------------------------------
BATCH_PLACEHOLDER_HINT = "把要替换的地方写成 {{姓名}} 这种格式，不懂就先下示例看看。"


@st.dialog("生成前确认")
def _batch_dialog() -> None:
    """未预览就直接批量生成时的确认弹窗。"""
    st.markdown("**建议先试看第一份，确认没问题再全部生成，是否继续？**")
    st.caption("直接生成会一次产出多份文件，先试看能提前发现填错位置、文字太长挤出去的问题。")
    left, right = st.columns(2)
    with left:
        if st.button("先试看一份", key="batch_dialog_cancel", use_container_width=True):
            st.session_state["batch_skip_preview"] = False
            st.session_state["batch_ask"] = False
            st.rerun()
    with right:
        if st.button("不用了，直接生成", type="primary", key="batch_dialog_ok", use_container_width=True):
            st.session_state["batch_skip_preview"] = True
            st.session_state["batch_ask"] = False
            st.rerun()


def render_batch() -> None:
    page_head(
        "批量生成",
        "1 份名单 + 1 份模板，一键生成几百份文件。合同、Offer、对账单不用再一份份改。",
    )

    with card("batch"):
        col_left, col_mid, col_right = st.columns(3, gap="large")

        # ---- 第 1 步：名单 ----
        with col_left:
            step_head("第 1 步：准备名单", "每一行代表一个人/一份文件")
            excel_files = _render_uploader(
                "上传 Excel（支持 .xlsx / .csv）",
                "batch_excel",
                ["xlsx", "xlsm", "csv", "xls"],
                caption_tip="支持 .xlsx / .csv",
                dragzone_tip=DRAGZONE_TIP_EXCEL,
            )
            excel = excel_files[0] if excel_files else None
            sample_download(
                "excel",
                "批量生成示例数据.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "batch_sample_excel",
            )

        # ---- 第 2 步：模板 ----
        with col_mid:
            step_head("第 2 步：准备模板", "把要改的地方留成 {{姓名}} 这样的空位")
            template_files = _render_uploader(
                "上传 PDF 模板",
                "batch_template",
                ["pdf"],
                caption_tip="支持 PDF",
                dragzone_tip=DRAGZONE_TIP_PDF,
            )
            template = template_files[0] if template_files else None
            sample_download(
                "pdf",
                "批量生成模板示例.pdf",
                "application/pdf",
                "batch_sample_pdf",
            )
            st.caption(BATCH_PLACEHOLDER_HINT)

        # ---- 第 3 步：生成 ----
        with col_right:
            step_head("第 3 步：一键生成", "先试看一份，没问题再全部打包")
            do_preview = st.button(
                "🔍 试看第一份效果",
                key="batch_preview_btn",
                use_container_width=True,
                disabled=any_busy("batch_preview"),
            )
            do_generate = st.button(
                "⚡ 批量生成全部打包下载",
                type="primary",
                key="batch_btn",
                use_container_width=True,
                disabled=any_busy("batch"),
            )

        # ---- 读取输入 ----
        table: batch_fill.TableData | None = None
        table_error = ""
        if excel is not None:
            try:
                table = batch_fill.read_table(excel.getvalue(), excel.name)
            except batch_fill.BatchError as exc:
                table_error = str(exc)

        template_bytes = template.getvalue() if template is not None else b""

        if table_error:
            st.error(table_error)

        if do_preview or do_generate:
            if table is None:
                st.warning("请先上传名单（Excel 或 CSV）。")
                return
            if not template_bytes:
                st.warning("请先上传 PDF 模板。")
                return

        if do_preview:

            def _do_preview() -> None:
                with st.spinner("正在生成第一份给你看…"):
                    report = batch_fill.prepare_preview(table, template_bytes)
                if not report.ok:
                    st.error(report.message)
                else:
                    st.success(report.message)
                    for warning in report.warnings:
                        st.caption(f"ℹ️ {warning}")
                    preview_image(report.preview_png, report.preview_name)
                    st.caption("识别到的空位：" + "、".join(report.fields))

                    st.session_state["batch_previewed"] = True

            run_task("batch_preview", _do_preview)

        if do_generate:
            # 份数校验：数据表行数超过当前额度时不执行生成，直接弹升级提示
            row_limit = current_file_limit()
            if table is not None and len(table.rows) > row_limit:
                show_upgrade()
                st.caption(
                    f"你的名单有 {len(table.rows)} 行（要生成 {len(table.rows)} 份文件），"
                    f"超过单次能做的 {row_limit} 份了。"
                )
                return
            if not st.session_state.get("batch_previewed") and not st.session_state.get(
                "batch_skip_preview"
            ):
                if hasattr(st, "dialog"):
                    # 弹窗会在这轮渲染结束后弹出；用户选择“继续生成”后置位 skip 标记，
                    # 下一轮 rerun 时会带着「批量生成全部」按钮的状态重新执行。
                    _batch_dialog()
                else:
                    st.warning("建议先试看第一份，没问题再全部生成。")
                return

            def _do_batch() -> None:
                with st.spinner("正在一份份生成并打包…"):
                    result = batch_fill.build_zip(table, template_bytes, limit=row_limit)
                st.session_state["batch_skip_preview"] = False
                if not result.ok:
                    st.error(result.message)
                    return
                st.success(result.message)
                st.download_button(
                    "⬇ 立刻下载全部文件（ZIP）",
                    data=result.output or b"",
                    file_name=result.filename,
                    mime="application/zip",
                    type="primary",
                    key="batch_download",
                    use_container_width=True,
                )
                with st.expander("生成明细"):
                    st.code("\n".join(result.log) or "（无）", language="text")

            run_task("batch", _do_batch)


WATERMARK_TYPES = ["png", "jpg", "jpeg", "webp", "pdf", "docx", "pptx"]
WATERMARK_HELP = f"支持格式：PNG/JPG/WEBP/PDF/Word/PPT，单文件最大 {BASE_MAX_UPLOAD_MB}MB"
WATERMARK_TIP = "💡 提示：当前支持去除纯白底色的文字/印章水印，以及文档的页眉页脚。"
WATERMARK_TIP_SUB = "复杂背景的图片水印效果可能受限。"


def _watermark_outcomes(files: list[Any], mode: str) -> list[tuple[Any, Any]]:
    """逐份处理，返回 [(上传文件, WatermarkResult), ...]。"""
    results = []
    for file in files:
        name = display_name(getattr(file, "name", ""), "文件")
        results.append(
            (file, watermark_remove.remove_watermark(file.getvalue(), name, mode))
        )
    return results


def _zip_bytes(items: list[tuple[str, bytes]]) -> bytes:
    """把多份结果打包成 ZIP（同名自动加序号），供一次下载。"""
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        used: dict[str, int] = {}
        for name, payload in items:
            if name in used:
                used[name] += 1
                stem, dot, ext = name.rpartition(".")
                name = f"{stem}({used[name]}){dot}{ext}"
            else:
                used[name] = 1
            archive.writestr(name, payload)
    return buffer.getvalue()


def render_watermark() -> None:
    page_head(
        "一键消除水印",
        "传上来点一下就好，正文一个字都不动。处理完文件立刻销毁，不留底。",
    )

    with card("watermark"):
        files = uploader(
            "上传带水印的文件",
            "watermark_upload",
            WATERMARK_TYPES,
            multiple=True,
            dragzone_tip=DRAGZONE_TIP_OFFICE,
            caption_tip=WATERMARK_HELP,
        )
        # 用 st.info 承载提示：旧 st.warning 的黄底在深色主题下会变成浅黄字，看不清。
        st.info(WATERMARK_TIP)
        st.markdown(
            f'<div class="zjb-tip-text">{WATERMARK_TIP_SUB}</div>',
            unsafe_allow_html=True,
        )

        # 默认按图片列出全部位置；上传了 Word / PPT 时自动收敛到文档能用的几项，
        # 避免出现“选了‘裁掉顶部’但文件是 docx”这种对不上的选项。
        uploaded = list(files or [])
        doc_only = bool(uploaded) and all(
            str(getattr(f, "name", "")).rsplit(".", 1)[-1].lower() in ("docx", "pptx")
            for f in uploaded
        )
        choices = list(watermark_remove.DOC_MODES if doc_only else watermark_remove.IMAGE_MODES)
        current = st.session_state.get("watermark_mode", "auto")
        if current not in choices:
            st.session_state["watermark_mode"] = "auto"
        mode = st.selectbox(
            "水印在哪？",
            choices,
            format_func=watermark_remove.region_label,
            key="watermark_mode",
            help="不知道怎么选就先试第一个「自动识别并去除」；"
            "水印压在某一块就选那一块，处理范围完全可预期。",
        )

        if st.button(
            "⚡️ 一键闪速去水印",
            type="primary",
            key="watermark_btn",
            use_container_width=True,
            disabled=any_busy("watermark"),
        ):
            if not uploaded:
                st.warning("请先上传至少一份要处理的文件。")
            else:

                def _do_watermark() -> None:
                    with st.spinner("正在擦掉水印…"):
                        outcomes = _watermark_outcomes(uploaded, mode)
                    succeeded = [(f, r) for f, r in outcomes if r.ok]
                    failed = [(f, r) for f, r in outcomes if not r.ok]

                    if succeeded:
                        st.success(f"搞定！{len(succeeded)} 份文件已经去好水印，可以下载了。")
                        for _, result in succeeded:
                            preview_image(result.preview_png, result.filename or "处理结果预览")
                        payouts = [(result.filename, result.output) for _, result in succeeded]
                        if len(payouts) == 1:
                            st.download_button(
                                "📥 下载无水印文件",
                                data=payouts[0][1],
                                file_name=payouts[0][0],
                                mime=succeeded[0][1].mime,
                                type="primary",
                                key="watermark_download",
                                use_container_width=True,
                            )
                        else:
                            st.download_button(
                                "📥 下载无水印文件（全部打包 ZIP）",
                                data=_zip_bytes(payouts),
                                file_name="去水印结果.zip",
                                mime="application/zip",
                                type="primary",
                                key="watermark_download_zip",
                                use_container_width=True,
                            )
                        with st.expander("处理明细"):
                            st.code(
                                "\n".join(
                                    f"{result.filename} · {result.message}" for _, result in succeeded
                                ),
                                language="text",
                            )

                    for file, result in failed:
                        st.error(
                            f"{display_name(getattr(file, 'name', ''), '文件')}：{result.message}"
                        )

                run_task("watermark", _do_watermark)


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(
        page_title=f"{APP_NAME} - {APP_SLOGAN}",
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    # 字体 + 主题样式
    st.markdown(FONT_LINKS, unsafe_allow_html=True)
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # ── 内测邀请码验证守卫 ──
    if not st.session_state.get(VERIFY_KEY):
        st.markdown(
            f"<h1 style='text-align:center;margin-top:60px;'>⚡️ {APP_NAME}</h1>"
            f"<p style='text-align:center;color:#6b7280;'>{APP_CLAIM}</p>",
            unsafe_allow_html=True,
        )
        code = st.text_input("请输入内测邀请码", placeholder="VIP-001 / XHS-001 …", key="invite_input")
        if st.button("进入", key="invite_btn", use_container_width=True):
            normalized = (code or "").strip().upper()
            if normalized in VALID_CODES:
                st.session_state[VERIFY_KEY] = True
                st.rerun()
            else:
                st.error("邀请码无效，请前往闲鱼/小红书私信获取")
        st.stop()

    render_sidebar()
    render_header()

    # 四大纯本地板块（全部在浏览器与服务之间完成，不调用任何外部 API）
    tab_basic, tab_replace, tab_batch, tab_watermark = st.tabs(
        [
            "PDF基础处理",
            "PDF/图片改字",
            "批量生成",
            "去水印",
        ]
    )
    with tab_basic:
        render_basic()
    with tab_replace:
        render_replace()
    with tab_batch:
        render_batch()
    with tab_watermark:
        render_watermark()

    st.markdown(
        '<div class="zjb-footer">'
        f'<span class="zjb-footer-brand">{APP_ICON} {APP_NAME}</span>'
        f"<br>{APP_COPYRIGHT}"
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
