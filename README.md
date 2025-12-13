# 文档自动提取关键信息系统

## 功能说明

本系统可以自动处理本地文档（Word、PDF），识别文档类型，提取关键信息，并将结果存储到对应的Excel表格中。

## 快速开始

### 1. 环境准备

```bash
# 激活 conda 环境
conda activate big-data-doc-auto-shell

# 确保已安装依赖
pip install -e .
```

### 2. 配置检查

确保 `.env` 文件已正确配置：
- `OPENAI_API_KEY`: API密钥
- `OPENAI_BASE_URL`: API地址
- `MODEL_NAME`: 模型名称

### 3. 运行程序

```bash
python main.py
```

## 使用说明

1. 将待处理的文档放入 `data/` 目录
2. 运行 `python main.py`
3. 程序会自动：
   - 扫描 `data/` 目录下的所有 Word 和 PDF 文件
   - 使用 LLM 对文档进行分类
   - 提取关键信息
   - 将结果保存到 `output/` 目录对应的 Excel 文件中

## 输出文件

处理后的数据会保存在 `output/` 目录：
- `2办会材料信息.xlsx` - 办会材料
- `3办文材料信息.xlsx` - 办文材料
- `4政策文件信息.xlsx` - 政策文件

## 日志

处理日志保存在 `processing.log` 文件中。

