"""配置文件模块 - 加载环境变量和项目配置"""
import os
from pathlib import Path
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent

# 路径配置
DATA_DIR = Path(os.getenv("DATA_DIR", "data"))
TEMPLATE_DIR = Path(os.getenv("TEMPLATE_DIR", "template"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output"))

# 确保路径是相对于项目根目录的
if not DATA_DIR.is_absolute():
    DATA_DIR = PROJECT_ROOT / DATA_DIR
if not TEMPLATE_DIR.is_absolute():
    TEMPLATE_DIR = PROJECT_ROOT / TEMPLATE_DIR
if not OUTPUT_DIR.is_absolute():
    OUTPUT_DIR = PROJECT_ROOT / OUTPUT_DIR

# OpenAI API 配置（兼容 Qwen 大模型）
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "gpt-3.5-turbo")

# API 请求配置
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "60"))

# 并发配置
MAX_WORKERS = int(os.getenv("MAX_WORKERS", "5"))  # 默认5个并发线程

# Excel模板文件映射
TEMPLATE_MAPPING = {
    "办会材料信息": "2办会材料信息.xlsx",
    "办文材料信息": "3办文材料信息.xlsx",
    "政策文件信息": "4政策文件信息.xlsx",
    "政策问答对": "5政策问答对.xlsx",
}

# 文档类型映射（LLM返回的数字到类型名称）
DOC_TYPE_MAPPING = {
    "1": "办会材料信息",
    "2": "办文材料信息",
    "3": "政策文件信息",
}

# 验证配置
def validate_config():
    """验证配置是否正确"""
    errors = []
    
    if not OPENAI_API_KEY:
        errors.append("OPENAI_API_KEY 未设置")
    
    if not DATA_DIR.exists():
        errors.append(f"数据目录不存在: {DATA_DIR}")
    
    if not TEMPLATE_DIR.exists():
        errors.append(f"模板目录不存在: {TEMPLATE_DIR}")
    
    # 创建输出目录（如果不存在）
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    if errors:
        raise ValueError("配置错误:\n" + "\n".join(f"  - {e}" for e in errors))
    
    return True

