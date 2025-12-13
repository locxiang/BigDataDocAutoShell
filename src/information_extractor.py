"""信息提取模块 - 根据文档类型提取关键信息"""
import logging
import json
import re
from pathlib import Path
from typing import Optional, Dict, Any
from openai import OpenAI
from src.config import OPENAI_API_KEY, OPENAI_BASE_URL, MODEL_NAME, MAX_RETRIES, REQUEST_TIMEOUT

logger = logging.getLogger(__name__)


class InformationExtractor:
    """信息提取器"""
    
    # 各类型文档的提取Prompt模板
    EXTRACTION_PROMPTS = {
        "办会材料信息": """你是一个信息提取专家。请从以下文档中提取关键信息，并以JSON格式返回。

文档内容（可能来自Word或PDF文件）：
{content}

请提取以下字段：
- PolicyCategory（文件类别）：会议通知、会议方案、会议纪要、其他
- IssuingAuthority（所属单位）：发文单位
- EffectiveDate（成文日期）：格式为YYYY-MM-DD
- Position（会议层级）：科级、处级、局级（地厅级）、省部级、国家级
  注意：重庆是直辖市，行政级别与其他市不同。重庆市级会议相当于省部级，重庆区级会议相当于局级（地厅级），重庆区级部门会议相当于处级，重庆街道/镇级会议相当于科级。
- Topic（主题分类）：必须严格匹配以下类别之一，请仔细分析文档内容，选择最准确的类别：
  社会服务与治理、食品药品安全、教育、卫生健康与医疗、商务、经济与金融、环境保护与生态文明、创新创业发展、农业农村发展、工业和信息化、文化与旅游、城市建设与规划、国家能源发展与规划、网络与信息安全、国际交流与合作、交通管理与规划、其他
  注意：必须从上述列表中选择，不能自行创建新的类别名称。如果文档内容无法明确归类，则选择"其他"。
- Refrence（参考资料）：如有参考资料请提取，否则为空
- Remarks（备注）：提取文档的完整标题（文档正文中的标题）。标题通常位于文档开头，可能跨越多行。对于PDF文件，标题可能因为格式问题被分割成多行（中间可能有空格或换行符），需要将所有行合并成完整的标题，去除多余的空格和换行。对于Word文件，如果标题因为换行被分割，也需要合并。必须提取完整的标题，不能简写、省略或截断。政府文件标题通常较长且完整，请仔细检查文档开头部分（前5-10行），确保提取全部内容。如果没有明确的文档标题，则此字段为空。

请严格按照以下JSON格式返回，不要添加任何其他文字：
{{
  "PolicyCategory": "",
  "IssuingAuthority": "",
  "EffectiveDate": "",
  "Position": "",
  "Topic": "",
  "Refrence": "",
  "Remarks": ""
}}""",
        
        "办文材料信息": """你是一个信息提取专家。请从以下文档中提取关键信息，并以JSON格式返回。

文档内容（可能来自Word或PDF文件）：
{content}

请提取以下字段：
- PolicyCategory（文件类别）：工作总结、工作报告、工作要点、交流谈话、工作调研、工作要求、其他
- IssuingAuthority（所属单位）：发文单位
- EffectiveDate（成文日期）：格式为YYYY-MM-DD
- Position（材料层级）：科级、处级、局级（地厅级）、省部级、国家级
  注意：重庆是直辖市，行政级别与其他市不同。重庆市级会议相当于省部级，重庆区级会议相当于局级（地厅级），重庆区级部门会议相当于处级，重庆街道/镇级会议相当于科级。
- ObjectOriented（面向对象）：涉及到各有关单位等相关模糊名称的，使用各区级单位、镇街
- Topic（主题分类）：必须严格匹配以下类别之一，请仔细分析文档内容，选择最准确的类别：
  社会服务与治理、食品药品安全、教育、卫生健康与医疗、商务、经济与金融、环境保护与生态文明、创新创业发展、农业农村发展、工业和信息化、文化与旅游、城市建设与规划、国家能源发展与规划、网络与信息安全、国际交流与合作、交通管理与规划、其他
  注意：必须从上述列表中选择，不能自行创建新的类别名称。如果文档内容无法明确归类，则选择"其他"。
- Refrence（参考资料）：如有参考资料请提取，否则为空
- Language（公文文种）：决议、决定、命令、公报、公告、通告、意见、通知、通报、报告、请示、批复、议案、函、纪要、其他
- Remarks（备注）：提取文档的完整标题（文档正文中的标题）。标题通常位于文档开头，可能跨越多行。对于PDF文件，标题可能因为格式问题被分割成多行（中间可能有空格或换行符），需要将所有行合并成完整的标题，去除多余的空格和换行。对于Word文件，如果标题因为换行被分割，也需要合并。必须提取完整的标题，不能简写、省略或截断。政府文件标题通常较长且完整，请仔细检查文档开头部分（前5-10行），确保提取全部内容。如果没有明确的文档标题，则此字段为空。

请严格按照以下JSON格式返回，不要添加任何其他文字：
{{
  "PolicyCategory": "",
  "IssuingAuthority": "",
  "EffectiveDate": "",
  "Position": "",
  "ObjectOriented": "",
  "Topic": "",
  "Refrence": "",
  "Language": "",
  "Remarks": ""
}}""",
        
        "政策文件信息": """你是一个信息提取专家。请从以下文档中提取关键信息，并以JSON格式返回。

文档内容（可能来自Word或PDF文件）：
{content}

请提取以下字段：
- PolicyCategory（政策类别）：党内法规与党建制度、国务院文件、国家各部委规章、地方法规、政府规章、行政规范性文件、政策解读、其他文件
- DocumentNumber（发文文号）：文件的文号
- IssuingAuthority（发布单位）：发文单位
- EffectiveDate（成文日期）：格式为YYYY-MM-DD
- ImplementationDate（施行日期）：格式为YYYY-MM-DD，如无则为空
- ValidUntil（有效期至）：格式为YYYY-MM-DD，如无则为空
- ResponsibleDepartment（责任部门）：使用发文部门
- CollaborativeDepartment（协同部门）：如有协同部门请提取，否则为空
- ApplicableObject（适用对象）：适用对象描述
- Fields（涉及领域）：经济发展与财政、民生与社会服务、农业农村、资源环境与城乡建设、公共安全与监督、文化科教与旅游、重大项目建设
- Remarks（备注）：提取文档的完整标题（文档正文中的标题）。标题通常位于文档开头，可能跨越多行。对于PDF文件，标题可能因为格式问题被分割成多行（中间可能有空格或换行符），需要将所有行合并成完整的标题，去除多余的空格和换行。对于Word文件，如果标题因为换行被分割，也需要合并。必须提取完整的标题，不能简写、省略或截断。政府文件标题通常较长且完整，请仔细检查文档开头部分（前5-10行），确保提取全部内容。如果没有明确的文档标题，则此字段为空。

请严格按照以下JSON格式返回，不要添加任何其他文字：
{{
  "PolicyCategory": "",
  "DocumentNumber": "",
  "IssuingAuthority": "",
  "EffectiveDate": "",
  "ImplementationDate": "",
  "ValidUntil": "",
  "ResponsibleDepartment": "",
  "CollaborativeDepartment": "",
  "ApplicableObject": "",
  "Fields": "",
  "Remarks": ""
}}""",
    }
    
    def __init__(self):
        """初始化提取器"""
        if not OPENAI_API_KEY:
            raise ValueError("OPENAI_API_KEY 未设置")
        
        self.client = OpenAI(
            api_key=OPENAI_API_KEY,
            base_url=OPENAI_BASE_URL,
            timeout=REQUEST_TIMEOUT,
        )
        self.model = MODEL_NAME
    
    def extract(self, content: str, doc_type: str, file_name: str = "") -> Optional[Dict[str, Any]]:
        """
        提取文档关键信息
        
        Args:
            content: 文档内容
            doc_type: 文档类型（办会材料信息、办文材料信息、政策文件信息）
            file_name: 文件名（用于日志和提取文件名称）
            
        Returns:
            提取的信息字典，失败返回None
        """
        if not content:
            logger.warning(f"文档内容为空: {file_name}")
            return None
        
        if doc_type not in self.EXTRACTION_PROMPTS:
            logger.error(f"不支持的文档类型: {doc_type}")
            return None
        
        # 在Prompt中包含文件名信息
        prompt = self.EXTRACTION_PROMPTS[doc_type].format(content=content)
        
        for attempt in range(MAX_RETRIES):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                )
                
                result_text = response.choices[0].message.content.strip()
                
                # 尝试提取JSON
                json_data = self._parse_json(result_text)
                
                if json_data:
                    # 处理文件名：去掉扩展名
                    file_name_without_ext = Path(file_name).stem if file_name else ""
                    
                    # 强制设置PolicyFileName为文件名（不含扩展名）
                    json_data["PolicyFileName"] = file_name_without_ext
                    
                    # 处理Remarks字段：如果文档标题与文件名一致，则清空Remarks
                    if "Remarks" in json_data and json_data["Remarks"]:
                        doc_title = json_data["Remarks"].strip()
                        # 判断文档标题是否与文件名一致（去除空格、标点符号后比较）
                        title_normalized = doc_title.replace(" ", "").replace("　", "").replace("、", "").replace("（", "(").replace("）", ")")
                        file_name_normalized = file_name_without_ext.replace(" ", "").replace("　", "").replace("、", "").replace("（", "(").replace("）", ")")
                        
                        # 如果标题与文件名一致，清空Remarks
                        if title_normalized == file_name_normalized or doc_title == file_name_without_ext:
                            json_data["Remarks"] = ""
                        else:
                            # 保留完整的文档标题
                            json_data["Remarks"] = doc_title
                    
                    logger.info(f"信息提取成功: {file_name}")
                    return json_data
                else:
                    logger.warning(f"JSON解析失败: {file_name}, 返回: {result_text[:200]}")
                    if attempt < MAX_RETRIES - 1:
                        continue
                    return None
                    
            except Exception as e:
                logger.error(f"提取失败 (尝试 {attempt + 1}/{MAX_RETRIES}): {file_name}, 错误: {e}")
                if attempt < MAX_RETRIES - 1:
                    continue
                return None
        
        return None
    
    @staticmethod
    def _parse_json(text: str) -> Optional[Dict[str, Any]]:
        """
        从文本中解析JSON
        
        Args:
            text: 包含JSON的文本
            
        Returns:
            解析后的字典，失败返回None
        """
        # 尝试直接解析
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
        
        # 尝试提取JSON代码块
        json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', text, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError:
                pass
        
        # 尝试提取markdown代码块中的JSON
        code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
        if code_block_match:
            try:
                return json.loads(code_block_match.group(1))
            except json.JSONDecodeError:
                pass
        
        return None
    
    @staticmethod
    def clean_data(data: Dict[str, Any]) -> Dict[str, Any]:
        """
        数据清洗和验证
        
        Args:
            data: 原始数据字典
            
        Returns:
            清洗后的数据字典
        """
        # Topic字段的合法值列表
        VALID_TOPICS = [
            "社会服务与治理", "食品药品安全", "教育", "卫生健康与医疗",
            "商务", "经济与金融", "环境保护与生态文明", "创新创业发展",
            "农业农村发展", "工业和信息化", "文化与旅游", "城市建设与规划",
            "国家能源发展与规划", "网络与信息安全", "国际交流与合作",
            "交通管理与规划", "其他"
        ]
        
        cleaned = {}
        
        for key, value in data.items():
            if isinstance(value, str):
                # 去除前后空格
                value = value.strip()
                # 处理特殊字符（保留换行符）
                value = value.replace('\r\n', '\n').replace('\r', '\n')
                
                # 验证Topic字段
                if key == "Topic" and value:
                    # 检查是否严格匹配合法值
                    if value not in VALID_TOPICS:
                        # 尝试模糊匹配（去除空格、标点等）
                        normalized_value = value.replace(" ", "").replace("、", "")
                        matched = False
                        for valid_topic in VALID_TOPICS:
                            normalized_topic = valid_topic.replace(" ", "").replace("、", "")
                            if normalized_value == normalized_topic or normalized_value in normalized_topic or normalized_topic in normalized_value:
                                value = valid_topic
                                matched = True
                                logger.warning(f"Topic字段值 '{data.get(key)}' 已修正为 '{value}'")
                                break
                        
                        # 如果无法匹配，设置为"其他"
                        if not matched:
                            logger.warning(f"Topic字段值 '{value}' 不在合法列表中，已设置为'其他'")
                            value = "其他"
            
            cleaned[key] = value
        
        return cleaned

