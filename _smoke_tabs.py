#!/usr/bin/env python3
"""导航 Tab、页面文案与去水印页的回归测试（纯本地版：共 4 个顶层 Tab）。

运行：.venv/bin/python _smoke_tabs.py
"""
from __future__ import annotations

import io
import zipfile

from streamlit.testing.v1 import AppTest

EXPECTED_TABS = ["PDF基础处理", "PDF/图片改字", "批量生成", "去水印"]
BASIC_TABS = [
    "PDF 压缩",
    "PDF 转 Word",
    "PDF 合并",
    "文档转 PDF",
    "PDF 转 Excel",
    "PDF 转 PPT",
    "PDF 拆分",
]
EXPECTED_HEADS = {
    "PDF基础处理": "极速格式转换，免安装无广告。文件处理完即刻销毁，保护您的隐私。",
    "PDF/图片改字": "把文件里的字换掉：左边写原来的字，右边写新字，点一下就好。排版会自动对齐，改完可以先预览。",
    "批量生成": "1 份名单 + 1 份模板，一键生成几百份文件。合同、Offer、对账单不用再一份份改。",
    "一键消除水印": "传上来点一下就好，正文一个字都不动。处理完文件立刻销毁，不留底。",
}
# 已彻底移除的板块：界面上不允许再出现任何入口（含“敬请期待”式占位）
REMOVED_LABELS = ["批量提取", "AI写周报", "AI 帮你写周报", "录音转写", "敬请期待"]


def find(app: AppTest, kind: str, key: str):
    for item in app.get(kind):
        if item.key == key:
            return item
    raise AssertionError(f"找不到 {kind}[{key}]")


def check_tabs() -> AppTest:
    app = AppTest.from_file("app.py", default_timeout=180).run()
    assert not app.exception, app.exception

    tabs = [block.label for block in app.get("tabs")]
    # 顶层 4 个 Tab 的顺序（二级 Tab 会紧跟在自己的父 Tab 之后出现）
    top = [label for label in tabs if label not in BASIC_TABS]
    assert top == EXPECTED_TABS, top
    print("✔ 导航 Tab：", " / ".join(top))

    basic = [label for label in tabs if label in BASIC_TABS]
    assert basic == BASIC_TABS, basic
    print("✔ PDF基础处理 二级 Tab：", " / ".join(basic))

    markup = "\n".join(item.value for item in app.markdown if "zjb-page-head" in item.value)
    for title, desc in EXPECTED_HEADS.items():
        assert title in markup, f"缺少标题 {title}"
        assert desc in markup, f"{title} 副标题不符"
    print("✔ 四个页面大标题与副标题全部就位")

    # 已删除的板块：任何 UI 元素 / 文案 / Tab 上都不能再出现
    for removed in REMOVED_LABELS:
        assert not any(removed in label for label in tabs), f"Tab 仍残留 {removed}"
        assert removed not in markup, f"页面文案仍残留 {removed}"
    all_text = " ".join(
        str(getattr(item, "value", "") or "") + str(getattr(item, "label", "") or "")
        for kind in ("markdown", "caption", "button", "info", "warning", "title", "header")
        for item in app.get(kind)
    )
    for removed in REMOVED_LABELS:
        assert removed not in all_text, f"界面仍残留 {removed} 入口"
    print("✔ 批量提取 / AI 写周报 / 录音转写等入口已彻底清除（无占位残留）")

    # Hero 首屏：4 个元素的文案与语义标签必须齐备，顺序不能乱
    # 注意用 class="zjb-hero" 精确匹配，否则会连全局 <style> 块一起选中
    hero = "\n".join(item.value for item in app.markdown if 'class="zjb-hero"' in item.value)
    assert hero.startswith("<header"), hero[:80]
    order = [hero.index(token) for token in ("zjb-badge", "zjb-name", "zjb-slogan", "zjb-tagline")]
    assert order == sorted(order), order
    for text in (
        "本地处理 · 无需注册 · 文件不留存",
        "闪办",
        "闪速处理，纯净办公。",
        "你的本地效率工具箱",
    ):
        assert text in hero, f"Hero 缺少文案：{text}"
    # 闪电必须是实心、锐利的 SVG（不是 emoji，也不再用模糊发光）
    assert "<svg" in hero and 'class="zjb-mark"' in hero, "闪电图标应为 SVG"
    assert "⚡" not in hero, "Hero 里不应再出现 emoji 闪电"
    assert 'fill="currentColor"' in hero, "闪电应为实心（currentColor 填充）"
    assert "polygon" in hero, "应使用 Lucide Zap 的 polygon 造型"
    assert "linearGradient" not in hero, "已改为纯色 #F59E0B，不再需要渐变"
    assert "<h1" in hero, "品牌名必须是 h1"
    assert hero.count("<h1") == 1, "整页只能有一个 h1"
    print("✔ Hero 首屏：标签 / 品牌名 / Slogan / 副标题齐全且顺序正确")

    # 已彻底改名 / 无序号标签残留
    for old in ("批量工作台", "批量提效工作台", "基础处理站", "精准无痕修改", "轻量去水印"):
        assert old not in markup, f"仍残留旧名称 {old}"
    assert "zjb-eyebrow" not in markup and "01 ·" not in markup
    print("✔ 旧名称与序号标签已清除")
    return app


def check_busy_lock(app: AppTest) -> None:
    """防连点：所有核心处理按钮都必须接上 busy 锁。"""
    import app as app_module

    # 1. 源码层面：每个核心按钮的参数块里都要有 disabled=any_busy(...)
    core_buttons = [
        "basic_compress_btn",
        "basic_word_btn",
        "basic_merge_btn",
        "basic_to_pdf_btn",
        "replace_btn",
        "batch_preview_btn",
        "batch_btn",
        "watermark_btn",
    ]
    with open(app_module.__file__, encoding="utf-8") as handle:
        code = handle.read()
    for key in core_buttons:
        index = code.index(f'key="{key}"')
        block = code[index : index + 260]
        assert "disabled=any_busy(" in block, f"{key} 未接防连点锁"
    print(f"✔ {len(core_buttons)} 个核心处理按钮均已接入防连点锁")

    # 2. 运行层面：别的任务在跑时，其它处理按钮必须置灰
    app.session_state[app_module.BUSY_KEY] = "compress"
    app.run()
    assert find(app, "button", "replace_btn").disabled is True, "有任务在跑时改字按钮应置灰"
    assert find(app, "button", "watermark_btn").disabled is True, "有任务在跑时去水印按钮应置灰"
    # 自己那个任务保持可点（重复点击由 run_task 内部去重），避免按钮状态自相矛盾
    assert find(app, "button", "basic_compress_btn").disabled is False

    app.session_state[app_module.BUSY_KEY] = None
    app.run()
    assert find(app, "button", "replace_btn").disabled is False, "任务结束后按钮应恢复可点"
    print("✔ 联动置灰：有任务运行时其它按钮 disabled，结束后自动恢复")


def check_watermark(app: AppTest) -> None:
    types = [t.lstrip(".") for t in find(app, "file_uploader", "watermark_upload").proto.type]
    assert types == ["png", "jpg", "jpeg", "webp", "pdf", "docx", "pptx"], types
    print("✔ 去水印支持格式：", " / ".join(types))

    # 位置下拉：图片场景应列出全部 7 种傻瓜式选项
    box = find(app, "selectbox", "watermark_mode")
    labels = list(box.options)
    assert labels[:4] == [
        "自动识别并去除（适合纯白底）",
        "在页面顶部（删除页眉区域）",
        "在页面底部（删除页脚区域）",
        "全屏底纹遮盖（强行覆盖）",
    ], labels
    assert len(labels) == 7, labels
    print("✔ 去水印位置选项：", " / ".join(labels))

    # 提示框文案必须走 info（黄底 warning 在暗色下看不清），且逐字匹配
    infos = [m.value or "" for m in app.get("info")]
    assert any(
        "当前支持去除纯白底色的文字/印章水印，以及文档的页眉页脚" in text for text in infos
    ), infos
    assert not any("水印上的字" in (t.label or "") for t in app.text_input), "旧的关键字输入框应已删除"
    print("✔ 提示语改为 info + 黑色说明文字，关键字输入框已删除")

    # Word / PPT 必须是真实的 zip 包结构，否则会被判定为“文件已损坏”
    def office_bytes(main_part: str) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                f'<Override PartName="/{main_part}" ContentType="application/octet-stream"/>'
                "</Types>",
            )
            archive.writestr(main_part, b"<root/>")
        return buffer.getvalue()

    find(app, "file_uploader", "watermark_upload").set_value(
        [
            ("模板演示.pptx", office_bytes("ppt/presentation.xml"), "application/vnd.ms-powerpoint"),
            ("报价单.docx", office_bytes("word/document.xml"), "application/msword"),
        ]
    ).run()
    find(app, "button", "watermark_btn").click().run()
    assert not app.exception, app.exception
    assert any("搞定" in s.value and "可以下载了" in s.value for s in app.success), [
        s.value for s in app.success
    ]
    print("✔ 去水印：Word / PPT 上传入口与占位处理")


if __name__ == "__main__":
    app = check_tabs()
    check_busy_lock(app)
    check_watermark(app)
    print("\nALL TAB CHECKS PASSED")
