# Hengda Excel Processing (CLI Version)

## 📌 项目简介

这是一个面向企业财务/开票场景的 Excel 自动化处理工具，核心目标是：

> **从原始开票数据中筛选特定业务记录，结合辅助映射表补全信息，并重组为可用于后续分析的标准化模板。**

项目基于 Python 实现，主要用于替代人工 Excel 处理流程，提高数据处理效率与准确性。

---

## 🧠 核心业务逻辑

该项目围绕以下数据流展开：

### 输入数据

* `input_source.xlsx`：原始开票数据（主数据源）
* `help.xlsx`：辅助映射表（用于查找生产编号）

---

### 核心处理流程

1. **筛选业务记录**

   * 从原始 Excel 中筛选出备注列（第 27 列）中包含“含”的记录

2. **字段清洗与提取**

   * 从备注字段中提取软件相关信息
   * 从货物名称中解析有效字段（基于固定格式规则）

3. **跨表数据补全**

   * 以“发票号码”为 key
   * 在 `help.xlsx` 中查找对应的生产编号
   * 补全到目标数据中

4. **模板重组输出**

   * 将清洗后的数据重组为新的结构化表
   * 输出为标准 Excel 文件，供后续财务分析使用

---

### 输出结果

* `output.xlsx`

  * 包含筛选 + 清洗 + 补全后的结构化数据

---

## 🏗️ 项目架构

项目采用**简单分层结构**，将业务逻辑与基础操作分离：

```text
src/
├── modify_template.py     # 核心业务逻辑（筛选 + 清洗 + 补全 + 输出）
├── compare_files.py       # 数据核对（自动结果 vs 人工结果）
├── data_processing.py     # Excel 基础操作（加载 / 保存）
```

### 模块职责说明

#### 1️⃣ modify_template.py（核心模块）

负责完整数据处理流程：

* 读取源 Excel
* 按业务规则筛选数据
* 字段清洗与解析
* 调用辅助表进行数据补全
* 重组并输出目标表

👉 **这是项目最核心的部分**

---

#### 2️⃣ compare_files.py

用于结果校验：

* 按发票号码聚合金额
* 对比自动生成数据与人工数据
* 输出差异结果

👉 用于验证自动化处理的正确性

---

#### 3️⃣ data_processing.py

基础工具模块：

* Excel 文件加载
* Excel 文件保存

👉 避免业务代码与 IO 操作耦合

---

## ⚙️ CLI 使用方式

项目当前以 **CLI（命令行）方式运行**，适合自动化执行和脚本集成。

---

### 1️⃣ 安装依赖

```bash
pip install -r requirements.txt
```

---

### 2️⃣ 执行模板处理

```bash
python -m src.modify_template
```

或在 Python 中调用：

```python
from src.modify_template import modify_excel_data

modify_excel_data(
    source_file="eg/input_source.xlsx",
    shengchan_file="eg/help.xlsx",
    target_file="output.xlsx"
)
```

---

### 3️⃣ 执行结果对比

```python
from src.compare_files import compare_excel_files

compare_excel_files(
    target_file="target.xlsx",
    hand_file="hand.xlsx",
    output_file="output_comparison.xlsx"
)
```

---

## 📁 项目结构

```text
hengda/
├── src/                    # 核心逻辑模块
├── eg/                     # 示例数据
├── tests/                  # 测试脚本（非标准测试）
├── include%/               # 早期实验代码
├── pandas_test1.py         # pandas 实验版本
├── README.md
└── requirements.txt
```

---

## ⚠️ 当前限制

* 依赖固定列索引（未使用字段名映射）
* 输入 Excel 格式需稳定一致
* 缺少统一 CLI 入口（当前为函数调用方式）
* 缺少标准化自动测试（pytest）

---

## 🚀 后续优化方向

* 使用字段名替代列索引（提升鲁棒性）
* 增加输入数据校验
* 提供统一 CLI 命令入口（argparse）
* 引入日志系统（logging）
* 重构为可复用内部工具

---

## 📊 项目总结

该项目的核心价值在于：

* 将**人工 Excel 处理流程自动化**
* 实现**多工作簿数据联动处理**
* 构建**从数据清洗到结果验证的完整链路**

👉 本质上是一个 **面向企业内部使用的数据自动化工具原型**
