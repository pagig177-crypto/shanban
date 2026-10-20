#!/usr/bin/env python3
"""去水印位置模式的回归测试（原 _smoke_extract_fields.py：批量提取已下线）。

覆盖：
    1. 自动识别模式：顶部 / 底部水印被擦除，正文保留
    2. 顶部 / 底部条带模式：只处理所选区域
    3. 裁切模式：尺寸真的变小
    4. 页面控件与文案：位置下拉存在、旧输入框与旧品牌已清除

运行：.venv/bin/python _smoke_watermark.py
"""
from __future__ import annotations

import pymupdf
from streamlit.testing.v1 import AppTest

import watermark_remove

FAILED: list[str] = []


def check(name: str, condition: bool, detail: str = "") -> None:
    print(("  ✅ " if condition else "  ❌ ") + name + (f"  [{detail}]" if detail and not condition else ""))
    if not condition:
        FAILED.append(name)


def has(app: AppTest, kind: str, key: str) -> bool:
    return any(getattr(item, "key", "") == key for item in app.get(kind))

import app as app_module
# --------------------------------------------------------------------------
# 去水印区域
# --------------------------------------------------------------------------
def corner_watermark_png() -> bytes:
    """白底 + 顶部浅灰“水印” + 中间深灰“正文” + 底部浅灰“页脚水印”。"""
    document = pymupdf.open()
    page = document.new_page(width=200, height=200)
    page.draw_rect(pymupdf.Rect(0, 0, 200, 200), color=None, fill=(1, 1, 1))
    page.draw_rect(pymupdf.Rect(20, 10, 120, 26), color=None, fill=(0.85, 0.85, 0.85))   # 顶部水印
    page.draw_rect(pymupdf.Rect(20, 90, 180, 110), color=None, fill=(0.2, 0.2, 0.2))    # 正文
    page.draw_rect(pymupdf.Rect(20, 180, 120, 194), color=None, fill=(0.85, 0.85, 0.85))  # 底部水印
    return page.get_pixmap(dpi=72).tobytes("png")


def pixel_at(png: bytes, x: int, y: int) -> tuple[int, int, int]:
    return tuple(pymupdf.Pixmap(png).pixel(x, y)[:3])


WHITE = (255, 255, 255)


def check_region() -> None:
    print("\n[1] 去水印位置模式")
    payload = corner_watermark_png()

    full = watermark_remove.remove_watermark(payload, "a.png")
    check("自动模式成功", full.ok, full.message)
    check("自动：顶部水印被擦", pixel_at(full.output, 60, 18) == WHITE, str(pixel_at(full.output, 60, 18)))
    check("自动：底部水印被擦", pixel_at(full.output, 60, 187) == WHITE)
    check("自动：正文保留", pixel_at(full.output, 100, 100) != WHITE)

    top = watermark_remove.remove_watermark(payload, "a.png", "strip")
    check("顶部模式成功", top.ok, top.message)
    check("顶部：顶部水印被涂白", pixel_at(top.output, 60, 18) == WHITE)
    check("顶部：底部水印保留", pixel_at(top.output, 60, 187) != WHITE, str(pixel_at(top.output, 60, 187)))
    check("顶部：正文保留", pixel_at(top.output, 100, 100) != WHITE)
    check("日志注明处理位置", any("页面顶部" in line for line in top.log), str(top.log))

    bottom = watermark_remove.remove_watermark(payload, "a.png", "footer")
    check("底部模式：底部水印被涂白", pixel_at(bottom.output, 60, 187) == WHITE)
    check("底部模式：顶部水印保留", pixel_at(bottom.output, 60, 18) != WHITE)
    check("底部模式：正文保留", pixel_at(bottom.output, 100, 100) != WHITE)

    # 裁切模式：尺寸必须真的变小，且水印所在那一边消失
    crop = watermark_remove.remove_watermark(payload, "a.png", "crop_bottom")
    check("裁切模式成功", crop.ok, crop.message)
    before, after = pymupdf.Pixmap(payload).height, pymupdf.Pixmap(crop.output).height
    check("裁切后高度变小", after < before, f"{before} → {after}")

    app = AppTest.from_file("app.py", default_timeout=180).run()
    check("页面有位置下拉", has(app, "selectbox", "watermark_mode"))
    check("不再有“水印上的字”输入框", not any("水印上的字" in (t.label or "") for t in app.text_input))
    texts = [m.value or "" for m in app.markdown]
    check("宣传语不再宣称万能去水印", not any("抹除不留痕迹" in text for text in texts))
    check("旧品牌已清除", not any("职接办" in text for text in texts))


def check_no_disk_trace() -> None:
    """「文件不留存」守卫：PDF 转 Word 不得把用户文档的内嵌字体写进本机缓存。"""
    print("\n[2] 文件不留存（不把用户文档内容落盘）")
    import glob
    import shutil
    import tempfile

    import fonts

    cache = fonts.CACHE_DIR
    shutil.rmtree(cache, ignore_errors=True)
    check("起点：字体缓存目录已清空", not cache.exists())

    with open("示例-中文报告.pdf", "rb") as handle:
        payload = handle.read()
    outcome = app_module.to_word(payload, "示例-中文报告.pdf")
    check("带内嵌字体的 PDF 转换成功", outcome.ok, outcome.message)

    left = sorted(path.name for path in cache.glob("*")) if cache.exists() else []
    check("转换后字体缓存目录仍为空（无用户文档字体落盘）", not left, str(left))

    prefix = ("zjb_compress_", "zjb_word_", "zjb_replace_", "pdf2word-", "pdf_replace_")
    stale = [
        path for pattern in prefix
        for path in glob.glob(f"{tempfile.gettempdir()}/{pattern}*")
    ]
    check("无遗留的项目临时目录", not stale, str(stale))


def main() -> None:
    check_region()
    check_no_disk_trace()
    print()
    if FAILED:
        print("❌ 未通过：" + "；".join(FAILED))
        raise SystemExit(1)
    print("ALL WATERMARK CHECKS PASSED")


if __name__ == "__main__":
    main()
