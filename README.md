# 合同批量生成工具

根据 Excel 中的合同信息，批量生成 Word 合同文件。

## 功能特性

- 支持 `{{字段名}}` 占位符格式
- 自动处理段落、表格、页眉页脚中的占位符
- 文件命名规则: `合同号-客户名称-BU名称.docx`
- 自动检查和安装依赖
- 支持国内 PyPI 镜像源（清华、阿里云等）

## 安装

```bash
# 克隆仓库
git clone https://github.com/你的用户名/contract_generator.git
cd contract_generator

# 安装依赖
pip install -r requirements.txt
```

## 使用方法

### 检查环境
```bash
python contract_generator.py --check
```

### 生成合同
```bash
# 使用默认配置
python contract_generator.py

# 指定文件路径
python contract_generator.py -e data/contracts.xlsx -t data/template.docx -o output/
```

### 命令行参数

| 参数 | 简写 | 说明 | 默认值 |
|------|------|------|--------|
| --excel | -e | Excel文件路径 | data/contracts.xlsx |
| --template | -t | Word模板路径 | data/template.docx |
| --output | -o | 输出目录 | output/ |
| --check | -c | 仅检查环境 | - |

## 文件结构

```
contract_generator/
├── contract_generator.py    # 主程序
├── create_samples.py        # 创建示例文件
├── requirements.txt         # 依赖列表
├── data/                    # 数据目录
│   ├── contracts.xlsx       # 合同信息Excel
│   └── template.docx        # Word模板
└── output/                  # 输出目录
```

## 模板格式

在 Word 模板中使用 `{{字段名}}` 作为占位符，例如：
- `{{合同号}}`
- `{{客户名称}}`
- `{{BU名称}}`
- `{{签约日期}}`

Excel 文件第一行为表头（字段名），后续行为数据。

## 依赖

- Python 3.9+
- openpyxl
- python-docx
