from streamlit.testing.v1 import AppTest

app = AppTest.from_file("app.py", default_timeout=180).run()
assert not app.exception, app.exception

# 先过邀请码门禁
app.session_state["verified"] = True
app.run()
assert not app.exception, app.exception


def find(kind, key):
    for item in app.get(kind):
        if item.key == key:
            return item
    raise AssertionError(f"NOT FOUND {kind}[{key}]")


find("file_uploader", "batch_excel")
find("file_uploader", "batch_template")
find("button", "batch_preview_btn")
btn = find("button", "batch_btn")
find("download_button", "batch_sample_excel")
find("download_button", "batch_sample_pdf")
print("PASS 1: all batch widgets exist")

assert btn.disabled is True, f"expected disabled, got {btn.disabled}"
print("PASS 2: generate button disabled when nothing uploaded")

caps = [c.value for c in app.caption]
assert any("第一行请填写表头" in c for c in caps), caps
assert any("不知道怎么弄" in c for c in caps), caps
assert any("请先上传名单和模板" in c for c in caps), caps
print("PASS 3: all three hints rendered")

mds = [m.value for m in app.markdown]
assert any("录取通知书" in m and "告别手动复制粘贴" in m for m in mds)
print("PASS 4: new subtitle rendered")

import batch_fill

find("file_uploader", "batch_excel").set_value(
    [("示例数据.csv", batch_fill.sample_table_csv(), "text/csv")]
)
find("file_uploader", "batch_template").set_value(
    [("模板.pdf", batch_fill.sample_template_pdf(), "application/pdf")]
)
app.run()
assert not app.exception, app.exception
btn = find("button", "batch_btn")
assert btn.disabled is False, f"expected enabled, got {btn.disabled}"
print("PASS 5: button enabled after both uploads")
print("ALL PASS")
