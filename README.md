# BOM2Cost_Upgrader

BOM (Bill of Materials) 自动分解计价文件升级工具。

将 `.xlsx` 公式文件中的计算数据，按 5 条匹配规则批量复制到对应的 `.xlsm` 升级文件中，保留 xlsm 内的公式、VBA 宏和 ActiveX 控件。

## 工作原理

1. 扫描 `input/` 目录，按**图号**（文件名第一个 `_` 前的字符串）自动配对 `.xlsx` 和 `.xlsm`
2. 读取 xlsx 中的计算值（`data_only=True`）
3. 通过字符串级 XML 操作将数据写入 xlsm（不使用 openpyxl 保存，避免破坏宏和控件）
4. 输出到 `output/` 目录

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 将配对的 Excel 文件放入 input/ 目录
# 格式：<图号>_*.xlsx 和 <图号>_*.xlsm

# 运行
python main.py
```

## 5 个数据匹配操作

| # | 工作表 | 操作 | 匹配方式 |
|---|--------|------|----------|
| 1 | 总表 | A4:B21 | 按行直接复制 |
| 2 | BOM分解 | Z4:Z6, S9 | 按单元格直接复制 |
| 3 | 材料费 | G,I,AB,AF,AU,AX,BA,BD,BE 列 | 按 B 列（索引）匹配行 |
| 4 | 加工费 | J:AU 列 | 按 B 列（索引）匹配行 |
| 5 | X部品及外购件 | 第 2 行起全部 | 按行直接复制 |

## 目录结构

```
BOM2Cost_Upgrader/
├── main.py              # 入口脚本
├── file_pairing.py      # 文件配对逻辑
├── upgrader.py          # 核心升级逻辑（5 个操作）
├── requirements.txt     # Python 依赖
├── input/               # 放入配对的 xlsx + xlsm 文件
└── output/              # 升级后的 xlsm 输出
```

## 注意事项

- `~$` 开头的 Excel 临时文件会自动忽略
- 每个图号必须恰好有 1 个 xlsx 和 1 个 xlsm，否则报错
- 输出文件会移除 `calcChain.xml`，Excel 打开时会自动重新计算
