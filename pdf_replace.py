#!/usr/bin/env python3
"""PDF 文字替换的“一站式”接口：读取 PDF → 应用替换规则 → 写出新 PDF。

和 pdf_edit.replace_text 的区别：
    pdf_edit.replace_text 一次只接受“一批规则 + 一个页码范围”，是底层实现；
    本模块把“一条规则可以带自己的页码”这种业务逻辑也收进来，任何入口
    （图形界面 / 命令行 / 网页）都只调用同一个函数，不再各自复制一遍。

规则格式（ReplaceRule）：
    ReplaceRule("钟智宇", "钟诗雅", "1")     # 第 1 页
    ReplaceRule("239", "237", "1,3")        # 第 1、3 页
    ReplaceRule("A", "B", "2-5")            # 第 2~5 页
    ReplaceRule("A", "B", None)             # 所有页（页码留空）

“所有页”的规则会合并成一批处理（更快）；限定页码的规则各自一批，
按顺序把上一批的结果作为下一批的输入，最终只写出一个文件。

用法：
    import pdf_replace

    result = pdf_replace.replace_pdf_text(
        "报告.pdf",
        [pdf_replace.ReplaceRule("钟智宇", "钟诗雅", "1"), pdf_replace.ReplaceRule("239", "237", "1")],
        output_pdf_path="~/Desktop/新报告.pdf",
    )
    print(result.ok, result.output_pdf_path, result.message)
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence, Union

import pdf_edit

ProgressCallback = pdf_edit.ProgressCallback

# 页码写法：空串 = 全部页；"1"、"1,3"、"2-5"
PagesText = str | None


class RuleError(ValueError):
    """替换规则本身不合法（例如页码格式写错）。"""


@dataclass
class ReplaceRule:
    """一条替换规则：把第 pages 页里的 old 换成 new。

    pages 用字符串描述页范围（与界面输入框一致）：None / "" 表示全部页，
    也可以写 "1"、"1,3"、"2-5"。转换失败会抛 RuleError。
    """

    old: str
    new: str
    pages: PagesText = None

    @property
    def same_length(self) -> bool:
        return len(self.old) == len(self.new)

    def page_list(self) -> list[int] | None:
        """把页码描述解析成 [1, 3, ...]；None 表示全部页。"""
        return parse_pages(self.pages)


def parse_pages(value: PagesText) -> list[int] | None:
    """'1' -> [1]；'1,3' -> [1, 3]；'2-5' -> [2, 3, 4, 5]；空 -> None（全部页）。"""
    text = (value or "").strip()
    if not text:
        return None
    pages: list[int] = []
    for chunk in text.split(","):
        part = chunk.strip()
        start_text, _, end_text = part.partition("-")
        start_text, end_text = start_text.strip(), end_text.strip()
        if not start_text.isdigit() or (end_text and not end_text.isdigit()):
            raise RuleError(f"页码格式不对：{part}（可写 1、1,3 或 2-5）")
        start = int(start_text)
        end = int(end_text) if end_text else start
        if start < 1 or end < start:
            raise RuleError(f"页码不对：{part}（从 1 开始，结束页不小于开始页）")
        pages.extend(range(start, end + 1))
    return sorted(set(pages))


@dataclass
class ReplaceResult:
    """替换结果。ok / output_pdf_path / message 是给各入口直接用的字段。"""

    source: str
    output_pdf_path: str | None = None
    ok: bool = False
    message: str = ""
    # 底层最后一次替换的明细（次数、文件大小等）
    detail: pdf_edit.EditResult | None = None
    log: list[str] = field(default_factory=list)

    @property
    def output(self) -> Path | None:
        """兼容 pdf_edit.EditResult 的 output 写法。"""
        return Path(self.output_pdf_path) if self.output_pdf_path else None

    @property
    def source_size(self) -> int:
        return self.detail.source_size if self.detail is not None else 0

    @property
    def output_size(self) -> int:
        return self.detail.output_size if self.detail is not None else 0

    def describe(self) -> str:
        if not self.ok:
            return self.message
        if self.detail is not None:
            return self.detail.describe()
        return self.message


def _to_rule(rule: Union[ReplaceRule, Sequence[str]]) -> ReplaceRule:
    """允许传 ReplaceRule，也允许传 (旧, 新) 或 (旧, 新, 页码)。"""
    if isinstance(rule, ReplaceRule):
        return rule
    items = list(rule)
    if len(items) not in (2, 3):
        raise RuleError(f"替换规则应为 (旧, 新) 或 (旧, 新, 页码)，收到：{rule!r}")
    old, new = str(items[0]), str(items[1])
    pages = None if len(items) == 2 else str(items[2] or "")
    return ReplaceRule(old, new, pages)


def _plan_batches(
    rules: Sequence[ReplaceRule],
) -> list[tuple[list[tuple[str, str]], list[int] | None]]:
    """把所有规则按页码分批：全部页合并为一批，限定页码的各自一批。"""
    parsed = [(rule.old, rule.new, rule.page_list()) for rule in rules]

    batches: list[tuple[list[tuple[str, str]], list[int] | None]] = []
    everything = [(old, new) for old, new, pages in parsed if pages is None]
    if everything:
        batches.append((everything, None))
    batches += [([(old, new)], pages) for old, new, pages in parsed if pages is not None]
    return batches


def replace_pdf_text(
    input_pdf_path: str | Path,
    rules: Sequence[Union[ReplaceRule, Sequence[str]]],
    output_pdf_path: str | Path | None = None,
    *,
    match_size: bool = True,
    target_size: int | None = None,
    overwrite: bool = True,
    log: ProgressCallback | None = None,
    keep_last: int = 0,
) -> ReplaceResult:
    """读取 PDF，按规则替换文字，写到 output_pdf_path。

    input_pdf_path:  源 PDF。
    rules:           ReplaceRule 列表，或 (旧, 新[, 页码]) 元组列表。
    output_pdf_path: 输出文件；省略时写到源文件旁边，名为“原名_已修改.pdf”。
                     默认允许覆盖（网页/批量场景每次输出到独立路径，不会冲突）。
    match_size:      True 时让输出文件大小与源文件一致（默认开启）。
    target_size:     想对齐的具体字节数；None 表示用源文件大小。
    log:             进度回调，接收字符串。
    keep_last:       每个内容流保留最后 keep_last 处不替换（0 = 全部替换），
                     供“只改一处”的逐页推进使用。

    返回 ReplaceResult（ok / output_pdf_path / message）。
    找不到文字、字数不等、编码不支持等业务失败都通过 ok=False 返回，不抛异常。
    """
    messages: list[str] = []

    def _log(message: str) -> None:
        messages.append(message)
        if log is not None:
            log(message)

    source = Path(input_pdf_path).expanduser()
    result = ReplaceResult(source=str(source))

    if not source.is_file():
        result.message = f"文件不存在：{source}"
        return result

    try:
        parsed = [_to_rule(rule) for rule in rules]
        if not parsed:
            raise RuleError("没有提供任何替换规则")
        batches = _plan_batches(parsed)
    except RuleError as exc:
        result.message = str(exc)
        return result

    destination = (
        Path(output_pdf_path).expanduser()
        if output_pdf_path
        else source.with_name(f"{source.stem}_已修改.pdf")
    )
    if destination.resolve() == source.resolve():
        result.message = "输出文件不能与源文件相同"
        return result
    if destination.exists() and not overwrite:
        result.message = f"输出已存在：{destination}（需要覆盖请传 overwrite=True）"
        return result

    _log(f"源文件：{source}")
    for rule in parsed:
        pages_text = (rule.pages or "").strip() or "全部页"
        _log(f"规则：{rule.old} → {rule.new}（{pages_text}）")

    # 多批处理时，中间结果放在临时目录；最后一批直接写到目标路径。
    workspace: tempfile.TemporaryDirectory[str] | None = None
    try:
        if len(batches) == 1:
            step_paths = [destination]
        else:
            workspace = tempfile.TemporaryDirectory(prefix="pdf_replace_")
            root = Path(workspace.name)
            step_paths = [root / f"step{index}.pdf" for index in range(len(batches) - 1)]
            step_paths.append(destination)

        pending = source
        last: pdf_edit.EditResult | None = None
        for index, ((pairs, pages), step_out) in enumerate(zip(batches, step_paths), start=1):
            if len(batches) > 1:
                scope = f"第 {'、'.join(map(str, pages))} 页" if pages else "全部页"
                summary = "，".join(f"{old}→{new}" for old, new in pairs)
                _log(f"[{index}/{len(batches)}] {summary}（{scope}）")
            last = pdf_edit.replace_text(
                pending,
                pairs,
                out=step_out,
                pages=pages,
                match_size=match_size,
                target_size=target_size,
                overwrite=True,
                log=_log,
                keep_last=keep_last,
            )
            if not last.ok:
                # 失败时别把半成品留在目标位置。
                if step_out == destination:
                    destination.unlink(missing_ok=True)
                result.message = last.message
                result.detail = last
                result.log = messages
                return result
            pending = last.output  # type: ignore[assignment]
    finally:
        if workspace is not None:
            workspace.cleanup()

    result.ok = True
    result.output_pdf_path = str(destination)
    result.detail = last
    result.message = last.describe() if last is not None else "完成"
    result.log = messages
    return result


if __name__ == "__main__":  # 简单自测：python pdf_replace.py 源.pdf "旧=新" "旧2=新2:页码"
    import sys

    if len(sys.argv) < 3:
        print(__doc__)
        raise SystemExit(2)
    checks = []
    for argument in sys.argv[2:]:
        old, _, rest = argument.partition("=")
        new, _, pages = rest.partition(":")
        checks.append(ReplaceRule(old, new, pages))
    outcome = replace_pdf_text(Path(sys.argv[1]), checks, log=print)
    print(outcome.ok, outcome.output_pdf_path, outcome.message)
    raise SystemExit(0 if outcome.ok else 1)
