"""文档分类模块 - 使用LLM进行文档分类"""
import logging
import json
from typing import Optional
from openai import OpenAI
from src.config import OPENAI_API_KEY, OPENAI_BASE_URL, MODEL_NAME, MAX_RETRIES, REQUEST_TIMEOUT, DOC_TYPE_MAPPING

logger = logging.getLogger(__name__)


class DocumentClassifier:
    """文档分类器"""
    
    # 分类Prompt模板
    CLASSIFICATION_PROMPT = """你是一个专业的文档分类专家。请仔细分析以下文档内容，判断它属于以下哪种类型：

类型说明：
1. 办会材料信息：专门与会议相关的文档。包括：
   - 会议通知（明确说明要召开会议，包含会议时间、地点、参会人员）
   - 会议方案（会议筹备方案、会议议程安排）
   - 会议纪要（记录会议讨论内容和决议）
   - 会议安排（会议日程、会议流程）
   关键特征：文档的核心目的是组织、通知或记录会议活动，必须明确涉及"会议"相关内容。

2. 办文材料信息：日常行政办公文档和公文。包括：
   - 公文类：函、通知、报告、请示、批复、意见、通报、决定等各类公文
   - 工作文档：工作总结、工作报告、工作要点、工作调研、工作要求
   - 交流文档：交流谈话、工作交流
   - 其他：工作安排、工作部署、工作协调等
   关键特征：用于日常行政办公、工作协调、信息传达、工作汇报等，不涉及会议组织。特别注意：如果文档标题或内容明确是"函"、"通知"（非会议通知）、"报告"等公文类型，应归类为办文材料信息。

3. 政策文件信息：政策性和法规性文档。包括：
   - 政策文件（政策规定、政策意见）
   - 法规文件（法律法规、规章）
   - 规范性文件（行政规范性文件）
   - 政策解读（政策说明、政策解释）
   关键特征：涉及政策制定、法规条文、政策解读等政策性内容。

重要区分原则：
- 如果文档是"函"（如"关于XX的函"），应归类为"办文材料信息"（类型2）
- 如果文档是"会议通知"、"会议纪要"等明确与会议相关，应归类为"办会材料信息"（类型1）
- 如果文档是普通"通知"但内容不涉及会议，应归类为"办文材料信息"（类型2）

文档内容：
{content}

请仔细分析文档的内容特征、格式特征、文档类型（如函、通知、报告等）和用途，准确判断文档类型。
请只返回类型编号（1、2或3），不要返回其他任何内容。"""
    
    def __init__(self):
        """初始化分类器"""
        if not OPENAI_API_KEY:
            raise ValueError("OPENAI_API_KEY 未设置")
        
        self.client = OpenAI(
            api_key=OPENAI_API_KEY,
            base_url=OPENAI_BASE_URL,
            timeout=REQUEST_TIMEOUT,
        )
        self.model = MODEL_NAME
    
    def classify(self, content: str, file_name: str = "") -> Optional[str]:
        """
        对文档进行分类
        
        Args:
            content: 文档内容
            file_name: 文件名（用于日志）
            
        Returns:
            文档类型名称（办会材料信息、办文材料信息、政策文件信息），失败返回None
        """
        if not content:
            logger.warning(f"文档内容为空: {file_name}")
            return None
        
        prompt = self.CLASSIFICATION_PROMPT.format(content=content)
        
        for attempt in range(MAX_RETRIES):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,  # 降低温度以提高分类准确性
                )
                
                result = response.choices[0].message.content.strip()
                
                # 提取数字（可能包含其他字符）
                doc_type_num = None
                for char in result:
                    if char in "123":
                        doc_type_num = char
                        break
                
                if doc_type_num and doc_type_num in DOC_TYPE_MAPPING:
                    doc_type = DOC_TYPE_MAPPING[doc_type_num]
                    logger.info(f"文档分类成功: {file_name} -> {doc_type}")
                    return doc_type
                else:
                    logger.warning(f"分类结果无效: {file_name}, 返回: {result}")
                    if attempt < MAX_RETRIES - 1:
                        continue
                    return None
                    
            except Exception as e:
                logger.error(f"分类失败 (尝试 {attempt + 1}/{MAX_RETRIES}): {file_name}, 错误: {e}")
                if attempt < MAX_RETRIES - 1:
                    continue
                return None
        
        return None

