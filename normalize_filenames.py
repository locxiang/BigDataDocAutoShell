"""文件名规范化脚本 - 处理文件名中的空格和括号问题，并同步更新Excel"""
import sys
import logging
import re
from pathlib import Path
from typing import List, Tuple
from openpyxl import load_workbook
from src.config import OUTPUT_DIR, TEMPLATE_MAPPING

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('normalize_filenames.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class FilenameNormalizer:
    """文件名规范化器"""
    
    # 支持的文档文件扩展名
    DOC_EXTENSIONS = {'.docx', '.doc', '.pdf', '.xlsx', '.xls'}
    
    # 分类文件夹到Excel文件的映射
    CATEGORY_TO_EXCEL = {
        "办会材料信息": "2办会材料信息.xlsx",
        "办文材料信息": "3办文材料信息.xlsx",
        "政策文件信息": "4政策文件信息.xlsx",
    }
    
    def __init__(self):
        """初始化规范化器"""
        self.output_dir = OUTPUT_DIR
        self.stats = {
            'total_files': 0,
            'files_renamed': 0,
            'excel_rows_updated': 0,
            'errors': 0,
        }
    
    def normalize_filename(self, filename: str) -> str:
        """
        规范化文件名
        
        规则：
        1. 空格、Tab等空白字符 -> 下划线
        2. 英文圆括号 () -> 中文圆括号 （）
        
        Args:
            filename: 原始文件名（不含路径，但包含扩展名）
            
        Returns:
            规范化后的文件名
        """
        # 分离文件名和扩展名
        path_obj = Path(filename)
        stem = path_obj.stem  # 不含扩展名的文件名
        suffix = path_obj.suffix  # 扩展名
        
        # 1. 替换空白字符（空格、Tab、换行等）为下划线
        # 使用正则表达式匹配所有空白字符
        normalized_stem = re.sub(r'\s+', '_', stem)
        
        # 2. 替换英文圆括号为中文圆括号
        normalized_stem = normalized_stem.replace('(', '（').replace(')', '）')
        
        # 组合文件名和扩展名
        normalized_filename = normalized_stem + suffix
        
        return normalized_filename
    
    def needs_normalization(self, filename: str) -> bool:
        """
        检查文件名是否需要规范化
        
        Args:
            filename: 文件名
            
        Returns:
            是否需要规范化
        """
        normalized = self.normalize_filename(filename)
        return normalized != filename
    
    def scan_files(self) -> List[Tuple[Path, str]]:
        """
        扫描output目录下的所有文档文件
        
        Returns:
            列表，每个元素为(文件路径, 分类名称)
        """
        files_to_process = []
        
        # 扫描三个分类文件夹
        for category in self.CATEGORY_TO_EXCEL.keys():
            category_dir = self.output_dir / category
            if not category_dir.exists():
                logger.info(f"分类文件夹不存在，跳过: {category_dir}")
                continue
            
            logger.info(f"扫描分类文件夹: {category_dir}")
            
            # 扫描该文件夹下的所有文档文件
            for file_path in category_dir.iterdir():
                if file_path.is_file() and file_path.suffix.lower() in self.DOC_EXTENSIONS:
                    self.stats['total_files'] += 1
                    
                    # 检查是否需要规范化
                    if self.needs_normalization(file_path.name):
                        files_to_process.append((file_path, category))
        
        return files_to_process
    
    def rename_file(self, file_path: Path, new_filename: str) -> bool:
        """
        重命名文件
        
        Args:
            file_path: 原文件路径
            new_filename: 新文件名
            
        Returns:
            是否重命名成功
        """
        try:
            old_filename = file_path.name
            new_path = file_path.parent / new_filename
            
            # 检查新文件名是否已存在
            if new_path.exists() and new_path != file_path:
                logger.warning(f"目标文件已存在，跳过重命名: {new_path}")
                return False
            
            file_path.rename(new_path)
            logger.info(f"文件重命名成功: {old_filename} -> {new_filename}")
            return True
        except Exception as e:
            logger.error(f"文件重命名失败: {file_path}, 错误: {e}")
            return False
    
    @staticmethod
    def _extract_key_from_header(header: str) -> str:
        """
        从表头中提取英文键名
        
        Args:
            header: 表头字符串，例如 "ID(序号)" 或 "PolicyCategory(文件类别)"
            
        Returns:
            提取的键名，例如 "ID" 或 "PolicyCategory"
        """
        if not header:
            return ""
        
        # 如果包含括号，提取括号前的部分
        if "(" in header:
            key = header.split("(")[0].strip()
            return key
        
        # 如果没有括号，直接返回
        return header.strip()
    
    def update_excel_filename(self, excel_file: Path, old_filename_without_ext: str, new_filename_without_ext: str) -> int:
        """
        更新Excel中PolicyFileName字段的值
        
        Args:
            excel_file: Excel文件路径
            old_filename_without_ext: 旧文件名（不含扩展名）
            new_filename_without_ext: 新文件名（不含扩展名）
            
        Returns:
            更新的行数
        """
        if not excel_file.exists():
            logger.warning(f"Excel文件不存在: {excel_file}")
            return 0
        
        try:
            wb = load_workbook(excel_file)
            if "YS" not in wb.sheetnames:
                logger.warning(f"Excel文件中没有YS Sheet: {excel_file}")
                wb.close()
                return 0
            
            ws = wb["YS"]
            
            # 查找PolicyFileName列
            header_row = 1
            policy_filename_col = None
            
            # 查找表头中的PolicyFileName列
            for col_idx, cell in enumerate(ws[header_row], start=1):
                header_value = cell.value if cell.value else ""
                key = self._extract_key_from_header(str(header_value))
                if key == "PolicyFileName":
                    policy_filename_col = col_idx
                    break
            
            if policy_filename_col is None:
                logger.warning(f"Excel文件中未找到PolicyFileName列: {excel_file}")
                wb.close()
                return 0
            
            # 查找匹配的行并更新
            updated_count = 0
            for row_idx in range(2, ws.max_row + 1):
                cell_value = ws.cell(row_idx, policy_filename_col).value
                if cell_value:
                    # 比较文件名（不含扩展名）
                    cell_filename = str(cell_value).strip()
                    if cell_filename == old_filename_without_ext:
                        # 更新为新文件名
                        ws.cell(row_idx, policy_filename_col).value = new_filename_without_ext
                        updated_count += 1
                        logger.info(f"更新Excel行 {row_idx}: {old_filename_without_ext} -> {new_filename_without_ext}")
            
            if updated_count > 0:
                wb.save(excel_file)
                logger.info(f"从 {excel_file.name} 中更新了 {updated_count} 行数据")
            
            wb.close()
            return updated_count
            
        except Exception as e:
            logger.error(f"更新Excel失败: {excel_file}, 错误: {e}")
            return 0
    
    def process_files(self, files_to_process: List[Tuple[Path, str]]):
        """
        处理文件：重命名并更新Excel
        
        Args:
            files_to_process: 需要处理的文件列表
        """
        for file_path, category in files_to_process:
            try:
                # 1. 计算新文件名
                old_filename = file_path.name
                new_filename = self.normalize_filename(old_filename)
                
                # 2. 提取文件名（不含扩展名）
                old_filename_without_ext = Path(old_filename).stem
                new_filename_without_ext = Path(new_filename).stem
                
                logger.info(f"处理文件: {old_filename} -> {new_filename}")
                
                # 3. 重命名文件
                if self.rename_file(file_path, new_filename):
                    self.stats['files_renamed'] += 1
                    
                    # 4. 更新Excel中的PolicyFileName字段
                    excel_file = self.output_dir / self.CATEGORY_TO_EXCEL[category]
                    updated_rows = self.update_excel_filename(
                        excel_file,
                        old_filename_without_ext,
                        new_filename_without_ext
                    )
                    self.stats['excel_rows_updated'] += updated_rows
                else:
                    self.stats['errors'] += 1
                    
            except Exception as e:
                logger.error(f"处理文件失败: {file_path}, 错误: {e}")
                self.stats['errors'] += 1
    
    def print_summary(self):
        """打印统计信息"""
        print("\n" + "=" * 60)
        print("文件名规范化完成统计")
        print("=" * 60)
        print(f"扫描文件总数: {self.stats['total_files']}")
        print(f"重命名文件数量: {self.stats['files_renamed']}")
        print(f"更新Excel行数: {self.stats['excel_rows_updated']}")
        print(f"错误数量: {self.stats['errors']}")
        print("=" * 60)
    
    def run(self):
        """运行规范化流程"""
        try:
            print("=" * 60)
            print("文件名规范化脚本")
            print("=" * 60)
            print(f"输出目录: {self.output_dir}")
            print()
            print("规范化规则:")
            print("  1. 空格、Tab等空白字符 -> 下划线 (_)")
            print("  2. 英文圆括号 () -> 中文圆括号 （）")
            print()
            
            # 1. 扫描文件
            print("步骤1: 扫描文件...")
            files_to_process = self.scan_files()
            print(f"扫描完成，共找到 {self.stats['total_files']} 个文件")
            print(f"需要规范化的文件: {len(files_to_process)} 个")
            print()
            
            if len(files_to_process) == 0:
                print("没有需要规范化的文件，退出")
                return
            
            # 2. 显示将要处理的文件列表
            print("将要处理的文件列表:")
            for file_path, category in files_to_process:
                old_filename = file_path.name
                new_filename = self.normalize_filename(old_filename)
                print(f"  [{category}] {old_filename} -> {new_filename}")
            print()
            
            # 3. 确认操作
            response = input(f"是否继续处理 {len(files_to_process)} 个文件？(y/n): ").strip().lower()
            
            if response != 'y':
                print("操作已取消")
                return
            
            print()
            
            # 4. 处理文件
            print("步骤2: 重命名文件并更新Excel...")
            self.process_files(files_to_process)
            print()
            
            # 5. 打印统计信息
            self.print_summary()
            
        except KeyboardInterrupt:
            print("\n\n程序被用户中断")
            sys.exit(0)
        except Exception as e:
            logger.error(f"程序运行失败: {e}", exc_info=True)
            print(f"\n错误: {e}")
            sys.exit(1)


def main():
    """主函数"""
    normalizer = FilenameNormalizer()
    normalizer.run()


if __name__ == "__main__":
    main()

