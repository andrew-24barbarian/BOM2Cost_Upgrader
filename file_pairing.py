"""文件配对模块：按图号匹配 xlsx/xlsm 文件。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class FilePair:
    drawing_no: str
    xlsx_path: Path
    xlsm_path: Path


def _extract_drawing_no(filename: str) -> str:
    """从文件名中提取图号（第一个 '_' 前的字符串）。"""
    return filename.split("_")[0]


def pair_files(input_dir: str | Path) -> list[FilePair]:
    """扫描 input_dir，按图号配对 xlsx 和 xlsm 文件。

    忽略 ~$ 开头的 Excel 临时文件。

    Returns:
        按图号排序的 FilePair 列表。

    Raises:
        ValueError: 如果某个图号没有恰好 1 个 xlsx 和 1 个 xlsm。
    """
    input_dir = Path(input_dir)
    xlsx_map: dict[str, Path] = {}
    xlsm_map: dict[str, Path] = {}

    for entry in sorted(input_dir.iterdir()):
        if not entry.is_file():
            continue
        name = entry.name
        if name.startswith("~$"):
            continue
        drawing = _extract_drawing_no(name)
        if name.endswith(".xlsx"):
            xlsx_map[drawing] = entry
        elif name.endswith(".xlsm"):
            xlsm_map[drawing] = entry

    all_keys = sorted(set(xlsx_map) | set(xlsm_map))
    errors: list[str] = []
    pairs: list[FilePair] = []

    for key in all_keys:
        has_xlsx = key in xlsx_map
        has_xlsm = key in xlsm_map
        if has_xlsx and has_xlsm:
            pairs.append(FilePair(key, xlsx_map[key], xlsm_map[key]))
        elif has_xlsx and not has_xlsm:
            errors.append(f"图号 {key}: 有 xlsx 但缺少 xlsm ({xlsx_map[key].name})")
        else:
            errors.append(f"图号 {key}: 有 xlsm 但缺少 xlsx ({xlsm_map[key].name})")

    if errors:
        raise ValueError("文件配对失败:\n" + "\n".join(errors))

    return pairs
