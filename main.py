"""BOM2Cost_Upgrader 入口脚本。"""

from __future__ import annotations

import sys
from pathlib import Path

from file_pairing import pair_files
from upgrader import upgrade


def main(input_dir: str = "input", output_dir: str = "output") -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists():
        print(f"错误: 输入目录不存在: {input_path}")
        sys.exit(1)

    output_path.mkdir(exist_ok=True)

    try:
        pairs = pair_files(input_path)
    except ValueError as e:
        print(f"错误: {e}")
        sys.exit(1)

    print(f"找到 {len(pairs)} 对文件:")
    success = 0
    for pair in pairs:
        print(f"  [{pair.drawing_no}] {pair.xlsx_path.name} <-> {pair.xlsm_path.name}")
        try:
            dest = output_path / pair.xlsm_path.name
            upgrade(pair.xlsx_path, pair.xlsm_path, dest)
            print(f"    -> 已保存到 {dest}")
            success += 1
        except Exception as e:
            print(f"    -> 升级失败: {e}")

    print(f"\n完成: {success}/{len(pairs)} 对文件处理成功")


if __name__ == "__main__":
    main()
