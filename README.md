# Hengda Excel Processing

一个面向企业财务/开票场景的 Excel 数据处理项目，当前主要用于：

- 从开票导出表中筛选包含“含”的软件相关记录
- 结合销售发票序时簿补全生产编号
- 生成适合后续成本分析的目标模板表
- 按发票号码汇总并比对自动生成表与人工表的金额差异

这个仓库更接近“内部工具原型 + 验证脚本”的状态，而不是已经完整封装好的产品。README 下面会明确区分“已实现能力”和“尚未完成/存在硬编码的部分”。

## 项目背景

仓库来自“企业智能化、自动化更新升级”相关科研/业务实践，业务目标大致包括：

- 处理企业开票与成本数据
- 自动形成嵌入式软件相关的成本计算表
- 保留并分析特殊开票数据
- 为月度财务分析和报告整理基础数据

从代码实际情况看，目前已经落地的重点是第一步：Excel 数据清洗、模板重组、人工结果比对。

## 当前已实现的能力

### 1. 模板转换

核心文件：[src/modify_template.py](/Users/yujunchen/Grad/Spring2026/hengda/src/modify_template.py)

已实现逻辑：

- 读取源 Excel 的第一个工作表
- 跳过表头后逐行处理数据
- 使用第 17 列和第 19 列计算“合计”
- 筛选备注列中包含“含”的记录
- 从货物名称中截取第二个 `*` 之后的内容
- 根据发票号码，到辅助表中查找生产编号
- 生成新的目标工作表 `ModifiedSheet`

输出列为：

- 序号
- 生产编号
- 开票日期
- 发票号码
- 客户名称
- 货物名称
- 备注软件名称
- 数量
- 不含税金额
- 税额
- 合计

### 2. 金额差异比对

核心文件：[src/compare_files.py](/Users/yujunchen/Grad/Spring2026/hengda/src/compare_files.py)

已实现逻辑：

- 分别读取人工表和目标表
- 以“发票号码”为 key 汇总不含税金额
- 计算 `Target - Hand` 的差值
- 将结果分成正差、负差、零差三类写入输出文件
- 用颜色标记差异方向

输出列为：

- 发票号码
- Target 不含税金额总和
- Hand 不含税金额总和
- 差值 (Target - Hand)

### 3. PyQt 图形界面原型

核心文件：[src/excel_merger.py](/Users/yujunchen/Grad/Spring2026/hengda/src/excel_merger.py)

界面中已经有：

- 选择源文件
- 选择目标文件
- 触发模板修改
- 触发文件比对

但这个 GUI 目前还是原型状态，存在未对齐的函数签名和未实现调用，不能视为稳定可用的正式入口。

## 仓库结构

```text
hengda/
├── src/
│   ├── data_processing.py     # Excel 读取/保存等基础函数
│   ├── modify_template.py     # 源表 -> 目标模板
│   ├── compare_files.py       # 按发票号码做金额比对
│   └── excel_merger.py        # PyQt 图形界面原型
├── tests/                     # 脚本式验证代码，不是标准测试套件
├── eg/                        # 示例 Excel 文件
├── include%/                  # 早期实验代码/特殊分支逻辑
├── pandas_test1.py            # 基于 pandas 的另一版模板转换实验
└── requirements.txt
```

## 示例文件

`eg/` 目录提供了几个样例工作簿：

- `input_source.xlsx`：源数据样例
- `help.xlsx`：辅助映射表，含生产编号等信息
- `target.xlsx`：目标模板样例
- `hand.xlsx`：人工整理后的对照样例

从样例文件可以看出，项目依赖较强的业务字段约定，而不是通用 Excel 处理框架。

## 依赖环境

`requirements.txt` 当前只有两个依赖：

```txt
PyQt5==5.15.7
openpyxl==3.0.10
```

建议使用 Python 3.10+。

安装方式：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

如果你只需要处理 Excel，不运行 GUI，实际核心依赖是 `openpyxl`。

## 使用方式

### 方式 1：直接调用模板转换函数

可以在 Python 中直接调用：

```python
from src.modify_template import modify_excel_data

modify_excel_data(
    source_file="eg/input_source.xlsx",
    shengchan_file="eg/help.xlsx",
    target_file="output.xlsx",
)
```

执行后会在目标文件中创建或更新 `ModifiedSheet`。

### 方式 2：直接调用差异比对函数

```python
from src.compare_files import compare_excel_files

compare_excel_files(
    target_file="target.xlsx",
    hand_file="hand.xlsx",
    output_file="output_comparison.xlsx",
)
```

### 方式 3：尝试运行 GUI 原型

```bash
python3 -m src.excel_merger
```

注意：当前 GUI 代码和底层函数并没有完全对齐，运行前通常还需要先修正部分调用关系。

## 数据假设

当前实现依赖大量固定列位和固定工作表位置，主要包括：

- 源数据默认取第一个工作表
- 备注字段默认位于第 27 列
- 发票号码默认位于第 4 列
- 客户名称默认位于第 8 列
- 开票日期默认位于第 9 列
- 数量默认位于第 15 列
- 不含税金额默认位于第 17 列
- 税额默认位于第 19 列
- 辅助表中发票号码默认位于第 5 列，生产编号位于第 2 列

这意味着：

- 只要上游 Excel 模板有列顺序变化，脚本就可能失效
- 当前代码没有字段名映射层，也没有严格的输入校验
- 更适合处理固定格式的企业内部报表，而不是任意 Excel 文件

## 已知限制

这是当前仓库最需要提前说明的部分。

### 1. GUI 入口未完成

[src/excel_merger.py](/Users/yujunchen/Grad/Spring2026/hengda/src/excel_merger.py) 中存在以下问题：

- `merge_files(...)` 被调用，但仓库中没有对应实现
- `modify_template(...)` 传参和 `modify_excel_data(...)` 的函数签名不一致
- 比对输出路径硬编码为 `output/output_comparison.xlsx`

### 2. 比对逻辑依赖工作表索引

[src/compare_files.py](/Users/yujunchen/Grad/Spring2026/hengda/src/compare_files.py) 默认读取：

- 人工表第 1 个工作表
- 目标表第 2 个工作表

但示例 `eg/target.xlsx` 当前只有一个工作表 `TopRows`。如果直接拿样例文件运行现有比对函数，会因为工作表索引假设不匹配而失败。

### 3. 输出列和业务目标还没完全对齐

项目背景里提到：

- 保留 `%` 开票数据并单独输出
- 自动检测成本偏高/偏低
- 支持完整成本计算表

但这些能力在当前主代码里还没有完整实现，部分只出现在实验脚本或背景说明里。

### 4. 缺少标准测试

`tests/` 目录中的内容目前主要是一次性脚本，例如：

- 截取前 250 行
- 删除 sheet
- 早期版本的比对逻辑

它们不是基于 `pytest` 的自动化测试，也没有形成稳定的回归验证体系。

### 5. 存在多份重复/演化中的实现

同类逻辑分散在以下位置：

- [src/modify_template.py](/Users/yujunchen/Grad/Spring2026/hengda/src/modify_template.py)
- [pandas_test1.py](/Users/yujunchen/Grad/Spring2026/hengda/pandas_test1.py)
- [tests/main.py](/Users/yujunchen/Grad/Spring2026/hengda/tests/main.py)
- [include%/src.py](/Users/yujunchen/Grad/Spring2026/hengda/include%/src.py)

说明项目还处于“方案迭代”阶段，主入口尚未完全收敛。

## 适合的使用方式

如果你现在要继续使用这个仓库，比较稳妥的方式是：

1. 以 [src/modify_template.py](/Users/yujunchen/Grad/Spring2026/hengda/src/modify_template.py) 为主流程
2. 用 `eg/` 里的样例文件确认输入列是否仍然一致
3. 根据实际报表格式，先修正 `compare_files.py` 的工作表选择逻辑
4. 暂时不要把 GUI 当成正式交付入口

## 后续建议

如果这个项目要变成可维护的内部工具，下一步最值得做的是：

1. 把所有列索引改成“按表头名匹配”
2. 提供一个明确的 CLI 入口，而不是只靠脚本或 GUI 原型
3. 统一 `src/` 和 `tests/`、`include%/` 中重复的实现
4. 补齐 `%` 开票单独输出、成本异常检测等业务需求
5. 增加最小可运行的自动化测试

## 目前状态总结

这不是一个“完整产品”，而是一个已经能处理部分真实业务数据的 Excel 自动化原型。它最有价值的部分是：

- 已经固化了一套企业报表清洗逻辑
- 能把源表转换成目标模板
- 能对人工表和自动表做基础金额核对

它目前的主要问题也很明确：

- 入口不统一
- 硬编码较多
- 文档与代码状态此前不一致

如果后续要继续开发，建议先围绕 `src/modify_template.py` 和 `src/compare_files.py` 做收敛。
