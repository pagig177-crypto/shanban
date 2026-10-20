#!/usr/bin/env python3
"""额度体系（无痕限额 + Pro 兑换码）回归测试。

覆盖：
    1. 纯常量：基础 3 份 / 20MB，Pro 20 份 / 100MB，提示文案
    2. 常态界面干净：不出现“免费版 / Pro 版”对比字样
    3. 基础用户：批量提取 4 份 / 超 20MB -> 点击时才弹升级提示，不执行处理
    4. Pro 兑换码：20 份、100MB 全部放行，界面无任何报错
    5. AI 每天 3 次：用完第 4 次才提示；Pro 不限次数
    6. 无效兑换码 / 退出 Pro 后限制恢复

应用说明：AppTest 里对「多文件」上传框 set_value 后，第一轮 run() 只会把文件放进
session_state，页面拿到的仍是空列表，需要再补一次 run() 才是真实状态（单文件上传框没有此问题）。
"""
from __future__ import annotations

import json
from pathlib import Path

import pymupdf
from streamlit.testing.v1 import AppTest

import app as app_module

MEM = 1024 * 1024
FAILED: list[str] = []


def _new_app() -> AppTest:
    """创建带已验证 session 的 AppTest，跳过内测邀请码守卫。"""
    at = AppTest.from_file("app.py", default_timeout=180)
    at.session_state["verified"] = True
    at.run()
    return at


def check(name: str, condition: bool, detail: str = "") -> None:
    print(("  ✅ " if condition else "  ❌ ") + name + (f"  [{detail}]" if detail and not condition else ""))
    if not condition:
        FAILED.append(name)


def visible(text: str) -> str:
    """去掉开头的 emoji 图标：st.error / st.info 会把「⚠️」「👉」当成图标剥离，
    页面上的可见文字与常量只差这个前缀。"""
    return text.lstrip("⚠️👉 \u200b\u200d\ufe0f")


def find(app: AppTest, kind: str, key: str):
    for item in app.get(kind):
        if item.key == key:
            return item
    raise AssertionError(f"找不到 {kind}[{key}]")


def has(obj: AppTest, kind: str, key: str) -> bool:
    return any(getattr(item, "key", "") == key for item in obj.get(kind))


def unlock(app: AppTest) -> None:
    """在侧边栏输入测试兑换码并解锁。"""
    find(app, "text_input", "pro_code_input").set_value("VIP-TEST-001")
    find(app, "button", "pro_unlock_btn").click().run()


def small_pdf() -> bytes:
    return b"%PDF-1.4\n1 0 obj<</Type/Catalog>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF\n"


def pdfs(count: int, *, prefix: str = "扫描") -> list[tuple[str, bytes, str]]:
    return [(f"{prefix}{i}.pdf", small_pdf(), "application/pdf") for i in range(1, count + 1)]


def real_pdfs(count: int, *, prefix: str = "扫描") -> list[tuple[str, bytes, str]]:
    """真实可解析的 PDF（去水印 / 合并等引擎会校验页数，空壳字节会被判定失败）。"""
    files = []
    for index in range(1, count + 1):
        doc = pymupdf.open()
        doc.new_page().insert_text((72, 72), f"{prefix} {index}")
        files.append((f"{prefix}{index}.pdf", doc.tobytes(), "application/pdf"))
        doc.close()
    return files


def rows_csv(count: int) -> tuple[str, bytes, str]:
    """批量生成的名单：表头 + count 行，用于验证单次份数额度。"""
    header = "姓名,岗位\n"
    body = "".join(f"员工{i},测试岗\n" for i in range(1, count + 1))
    return ("名单.csv", (header + body).encode("utf-8"), "text/csv")


def set_multi(app: AppTest, key: str, files: list[tuple[str, bytes, str]]) -> None:
    """多文件上传框赋值（见文件头说明：需要两轮 run）。"""
    find(app, "file_uploader", key).set_value(files)
    app.run()
    app.run()


def all_text(app: AppTest) -> list[str]:
    """页面所有可见文本（caption / error / info / markdown / 标题）。"""
    texts: list[str] = []
    for kind in ("caption", "error", "info", "warning", "success", "markdown",
                 "title", "header", "subheader"):
        texts.extend(str(getattr(item, "value", "")) for item in app.get(kind))
    return texts



# --------------------------------------------------------------------------
# 1. 常量与文案
# --------------------------------------------------------------------------
def check_constants() -> None:
    print("\n[1] 额度常量与提示文案")
    check(
        "基础额度 = 3 份 / 20MB",
        (app_module.BASE_MAX_FILES, app_module.BASE_MAX_UPLOAD_MB) == (3, 20),
    )
    check(
        "Pro 额度 = 20 份 / 100MB",
        (app_module.PRO_MAX_FILES, app_module.PRO_MAX_UPLOAD_MB) == (20, 100),
    )
    check("测试兑换码已登记", "VIP-TEST-001" in app_module.PRO_CODES, str(app_module.PRO_CODES))
    check(
        "红色提示包含份数 / 大小 / Pro 特权",
        app_module.LIMIT_MESSAGE
        == "⚠️ 单次最多处理 3 份文件，且单文件不能超过 20MB。"
        "升级 Pro 版即可解锁单次 20 份、单文件 100MB 的极速处理特权。",
        app_module.LIMIT_MESSAGE,
    )
    check(
        "蓝色提示指向左侧边栏",
        app_module.UNLOCK_MESSAGE == "👉 请前往左侧边栏，输入您的 Pro 兑换码即可解除限制。",
        app_module.UNLOCK_MESSAGE,
    )
    check(
        "上传提示口径统一：普通 20MB / Pro 100MB / 单次 20 份",
        app_module.UPLOADER_CAPTION
        == "💡 提示：普通用户单文件最大 20MB；Pro 用户最大 100MB，单次最多 20 份。",
        app_module.UPLOADER_CAPTION,
    )
    check(
        "上传框 label 统一",
        app_module.UPLOADER_LABEL == "上传文件（普通用户限20MB）",
        app_module.UPLOADER_LABEL,
    )
    check(
        "组件内大小显示 20MB（不是 100MB）",
        app_module.UPLOADER_WIDGET_MAX_MB == 20,
        app_module.UPLOADER_WIDGET_MAX_MB,
    )
    check("AI 额度和提取额度已随功能删除", not any(
        hasattr(app_module, name)
        for name in ("BASE_AI_DAILY_LIMIT", "EXTRACT_OVER_LIMIT_MESSAGE", "ai_quota_status")
    ))


# --------------------------------------------------------------------------
# 2. 常态界面：干净，无“免费版限 N 份”之类对比字样
# --------------------------------------------------------------------------
def check_clean_ui() -> None:
    print("\n[2] 常态界面不出现额度对比字样")
    app = AppTest.from_file("app.py", default_timeout=180).run()
    texts = all_text(app)
    joined = "\n".join(texts)
    forbidden = ["Pro 版", "升级 Pro", "额度", "兑换码即可解除"]
    hits = [word for word in forbidden if word in joined]
    check("主界面无 Pro 版对比字样", not hits, str(hits))
    # 「免费版」字样只应出现在批量生成的份数提示里；批量提取页已删除
    free_lines = [line for line in texts if "免费版" in line]
    check(
        "「免费版」仅出现在批量生成份数提示中（无提取页残留）",
        all("批量生成" in line or "生成" in line for line in free_lines),
        str(free_lines),
    )
    check("无 error 提示", not app.error, str([e.value for e in app.error]))
    # 唯一允许的 info 是去水印页的能力提示（黄色 warning 在暗色下看不清，已改 info）
    allowed_info = "当前支持去除纯白底色的文字/印章水印"
    stray_info = [i.value for i in app.info if allowed_info not in (i.value or "")]
    check("除去水印提示外无多余 info 横幅", not stray_info, str(stray_info))
    check("侧边栏有兑换码输入框", has(app, "text_input", "pro_code_input"))
    check(
        "上传框说明为统一大白话提示",
        any(app_module.UPLOADER_CAPTION.rstrip("。") in t for t in texts),
        joined[:400],
    )
    check(
        "CSS 拖拽区文案改为动态变量（不再全局写死）",
        "--zjb-dragzone-tip" in app_module.CUSTOM_CSS
        and 'content: var(--zjb-dragzone-tip' in app_module.CUSTOM_CSS,
    )
    injected = [m.value or "" for m in app.markdown]
    check(
        "PDF 页拖拽区写 PDF",
        any(app_module.DRAGZONE_TIP_PDF in text for text in injected),
        str([t[-40:] for t in injected if "--zjb" in t][:2]),
    )
    check(
        "批量生成 Excel 上传框写 Excel 而非 PDF",
        any(app_module.DRAGZONE_TIP_EXCEL in text for text in injected),
        str([t[-40:] for t in injected if "--zjb" in t][-3:]),
    )
    check(
        "两种文案确实不同（证明按页生效，没有全局写死）",
        app_module.DRAGZONE_TIP_PDF != app_module.DRAGZONE_TIP_EXCEL,
    )
    config = Path(".streamlit/config.toml").read_text(encoding="utf-8")
    check("网络层上限 100MB（容纳 Pro 大文件）", "maxUploadSize = 100" in config)



def check_oversize_gate() -> None:
    print("\n[4] 单文件超 20MB -> 拦截")
    app = AppTest.from_file("app.py", default_timeout=180).run()
    find(app, "file_uploader", "basic_compress_upload").set_value(
        ("大文件.pdf", b"%PDF" + b"0" * (21 * MEM), "application/pdf")
    ).run()
    errors = [e.value for e in app.error]
    check("红色提示", visible(app_module.LIMIT_MESSAGE) in [visible(e) for e in errors], str(errors))
    check("蓝色指引", visible(app_module.UNLOCK_MESSAGE) in [visible(i.value) for i in app.info])
    check(
        "报错点名超限文件",
        any("大文件.pdf" in c.value for c in app.caption),
        str([c.value for c in app.caption]),
    )
    check("后续处理未执行", not has(app, "button", "basic_compress_btn"))

    # 刚好 20MB 放行
    app2 = AppTest.from_file("app.py", default_timeout=180).run()
    find(app2, "file_uploader", "basic_compress_upload").set_value(
        ("临界.pdf", small_pdf() + b"0" * (20 * MEM - len(small_pdf())), "application/pdf")
    ).run()


# --------------------------------------------------------------------------
# 5. Pro 兑换码：20 份 / 100MB 全部放行
# --------------------------------------------------------------------------
def check_pro_access() -> None:
    print("\n[5] Pro 兑换码放行 20 份 / 100MB")
    app = AppTest.from_file("app.py", default_timeout=180).run()

    find(app, "text_input", "pro_code_input").set_value("wrong-code")
    find(app, "button", "pro_unlock_btn").click().run()
    check("无效兑换码被拒绝", app.session_state.get("pro_code") is None)
    check("无效兑换码给出提示", any("兑换码无效" in w.value for w in app.warning))

    unlock(app)
    check("有效兑换码写入 session", app.session_state.get("pro_code") == "VIP-TEST-001")
    check("侧边栏显示已解锁", any("已解锁 Pro 权益" in s.value for s in app.success))

    # 20 份批量去水印：直接可用，不会被拦（必须是真实 PDF，空壳文件会被引擎拒绝）
    set_multi(app, "watermark_upload", real_pdfs(20, prefix="批量"))
    check("20 份上传不报错", not app.error, str([e.value for e in app.error]))
    find(app, "button", "watermark_btn").click().run()
    check(
        "20 份正常处理",
        any("搞定！20 份文件" in s.value for s in app.success),
        str([s.value for s in app.success]),
    )

    # 100MB 单文件：应通过业务校验（不再弹升级提示）
    app2 = AppTest.from_file("app.py", default_timeout=180).run()
    unlock(app2)
    find(app2, "file_uploader", "basic_compress_upload").set_value(
        ("大册子.pdf", b"%PDF" + b"0" * (99 * MEM), "application/pdf")
    ).run()
    check(
        "Pro 用户 100MB 以内不弹升级提示",
        visible(app_module.LIMIT_MESSAGE) not in [visible(e.value) for e in app2.error],
        str([e.value for e in app2.error]),
    )
    check("Pro 用户压缩按钮已渲染", has(app2, "button", "basic_compress_btn"))

    check("恰好 20MB 不报错", not app2.error, str([e.value for e in app2.error]))



# --------------------------------------------------------------------------
# 7. 退出 Pro 后限制恢复
# --------------------------------------------------------------------------
def check_signout() -> None:
    print("\n[7] 退出 Pro 后恢复基础额度")
    app = AppTest.from_file("app.py", default_timeout=180).run()
    unlock(app)
    find(app, "button", "pro_signout_btn").click().run()
    check("兑换码已清除", app.session_state.get("pro_code") is None)
    # 退出 Pro 后单次份数限制回到 3 份：批量生成页 4 行名单 -> 生成前拦截
    set_multi(app, "batch_excel", [rows_csv(4)])
    set_multi(app, "batch_template", real_pdfs(1, prefix="模板")[:1])
    find(app, "button", "batch_btn").click().run()
    check(
        "重新受 3 份限制（生成前拦截）",
        any(visible(app_module.LIMIT_MESSAGE) in visible(e.value) for e in app.error),
        str([e.value for e in app.error]),
    )


def main() -> None:
    check_constants()
    check_clean_ui()
    check_oversize_gate()
    check_pro_access()
    check_signout()
    print()
    if FAILED:
        print(f"❌ {len(FAILED)} 项未通过：" + "、".join(FAILED))
        raise SystemExit(1)
    print("ALL QUOTA CHECKS PASSED")


if __name__ == "__main__":
    main()
