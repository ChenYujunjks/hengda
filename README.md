# 📊 Hengda Excel Processing（企业 Excel 自动化处理工具）

## 🧩 项目简介

这是一个面向企业财务 / 开票场景的 Python 自动化工具，主要用于：

* 从原始开票 Excel 中筛选符合业务规则的数据
* 结合辅助映射表补全关键字段（如生产编号）
* 将原始数据重组为标准化目标模板
* 对自动生成结果与人工整理结果进行金额差异比对

👉 本项目更偏向 **“企业内部自动化工具原型”**，而不是通用 Excel 处理库。

---

## 🎯 解决的问题（业务视角）

在实际业务中，企业经常遇到：

* 多个 Excel 表之间数据无法直接关联
* 原始数据字段不规范，需要清洗
* 人工整理报表耗时且容易出错
* 自动生成结果缺乏验证机制

本项目的目标是：

👉 **用 Python 将这一整套流程自动化**

---

## ⚙️ 整体处理流程

```text
输入：
    input_source.xlsx（原始开票数据）
    help.xlsx（辅助映射表：发票号 -> 生产编号）

处理流程：
    1. 筛选符合条件的数据（备注中包含“含”）
    2. 提取并清洗关键字段（如货物名称、备注信息）
    3. 基于发票号，从辅助表补全生产编号
    4. 重组为标准目标模板（ModifiedSheet）
    5. 与人工表（hand.xlsx）做金额差异比对

输出：
    target.xlsx（自动生成模板）
    output_comparison.xlsx（差异分析结果）
```

---

## 🏗️ 项目结构（架构设计）

```text
hengda/
├── src/
│   ├── modify_template.py   # 核心：数据清洗 + 模板重组
│   ├── compare_files.py    # 差异比对逻辑
│   ├── data_processing.py  # Excel 读写封装
│
├── eg/                     # 示例数据
│   ├── input_source.xlsx
│   ├── help.xlsx
│   ├── target.xlsx
│   └── hand.xlsx
│
├── tests/                  # 实验/验证脚本
├── README.md
└── requirements.txt
```

---

## 🧠 核心模块设计（重点）

### 1️⃣ 数据清洗与模板重组（modify_template.py）

核心函数：

```python
modify_excel_data(source_file, shengchan_file, target_file)
```

主要职责：

* 遍历源 Excel
* 按业务规则筛选数据
* 清洗字符串字段
* 通过发票号关联辅助表
* 输出结构化模板

👉 核心能力：

* Excel 数据清洗
* 多表数据关联
* 业务规则驱动处理

---

### 2️⃣ 数据关联（关键逻辑）

```python
scbh = get_scbh_from_shengchan(shengchan_file, fphm)
```

说明：

* 使用“发票号码”作为 key
* 在辅助表中查找“生产编号”
* 实现多 Excel 数据联动

👉 对应真实业务中的：
**跨表数据补全 / join 操作**

---

### 3️⃣ 字段清洗（非结构化处理）

```python
if "含" in row[26]:
```

```python
hwmc = row[11][row6index+1:]
```

说明：

* 从备注字段中筛选业务数据
* 从字符串中提取有效信息

👉 对应：
**非结构化字段清洗能力**

---

### 4️⃣ 差异比对（compare_files.py）

功能：

* 按发票号聚合金额
* 对比自动结果 vs 人工结果
* 输出差异（正 / 负 / 零）

👉 核心价值：

**验证自动化结果的正确性**

---

## 🖥️ CLI 使用方式

### 模板处理

```bash
python -c "
from src.modify_template import modify_excel_data
modify_excel_data(
    source_file='eg/input_source.xlsx',
    shengchan_file='eg/help.xlsx',
    target_file='output.xlsx'
)
"
```

---

### 差异比对

```bash
python -c "
from src.compare_files import compare_excel_files
compare_excel_files(
    target_file='target.xlsx',
    hand_file='hand.xlsx',
    output_file='output_comparison.xlsx'
)
"
```

---

## 📦 技术栈

* Python 3
* openpyxl（Excel 处理）
* 基础数据处理逻辑（类似 pandas 但更底层控制）

---

## ⚠️ 当前限制（很重要）

本项目目前是“可运行的业务原型”，存在：

* 列索引依赖（如 row[26]）
* 输入格式要求严格
* 缺少字段名映射层
* 异常处理不完善
* CLI 入口较简单

---

## 🚀 后续优化方向

* 改为基于“列名”而非“列索引”
* 增加输入校验
* 封装统一 CLI 工具（argparse / click）
* 增加日志系统
* 提升性能（避免逐行查找）

---

## 📌 项目总结

本项目实现了：

* 多 Excel 数据处理自动化
* 结构化清洗 + 模板重组
* 跨表数据补全
* 自动结果验证

👉 本质是一个：

**“面向企业业务流程的 Excel 自动化工具原型”**

