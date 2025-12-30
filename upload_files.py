"""批量并行上传文件脚本"""
import os
import sys
import logging
import time
import threading
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple

import requests
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('upload.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class FileUploader:
    """文件上传器"""
    
    def __init__(self):
        """初始化上传器"""
        # 从环境变量读取配置
        self.upload_dir = Path(os.getenv("UPLOAD_DIR", "data"))
        self.upload_url = os.getenv("UPLOAD_URL", "")
        self.token = os.getenv("UPLOAD_TOKEN", "")
        self.max_workers = int(os.getenv("UPLOAD_MAX_WORKERS", "5"))
        self.max_retries = int(os.getenv("UPLOAD_MAX_RETRIES", "3"))
        
        # 确保路径是绝对路径
        if not self.upload_dir.is_absolute():
            self.upload_dir = Path(__file__).parent / self.upload_dir
        
        # 统计信息
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'failed_files': [],
            'start_time': None,
            'end_time': None,
        }
        # 线程安全的统计锁
        self.stats_lock = threading.Lock()
        
        # 验证配置
        self._validate_config()
    
    def _validate_config(self):
        """验证配置是否正确"""
        errors = []
        
        if not self.upload_url:
            errors.append("UPLOAD_URL 未设置")
        
        if not self.token:
            errors.append("UPLOAD_TOKEN 未设置")
        
        if not self.upload_dir.exists():
            errors.append(f"上传目录不存在: {self.upload_dir}")
        
        if self.max_workers < 1:
            errors.append("UPLOAD_MAX_WORKERS 必须大于0")
        
        if self.max_retries < 0:
            errors.append("UPLOAD_MAX_RETRIES 必须大于等于0")
        
        if errors:
            raise ValueError("配置错误:\n" + "\n".join(f"  - {e}" for e in errors))
    
    def _get_content_type(self, file_path: Path) -> str:
        """
        根据文件扩展名获取Content-Type
        
        Args:
            file_path: 文件路径
            
        Returns:
            Content-Type字符串
        """
        ext = file_path.suffix.lower()
        content_types = {
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xls': 'application/vnd.ms-excel',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc': 'application/msword',
            '.pdf': 'application/pdf',
        }
        return content_types.get(ext, 'application/octet-stream')
    
    def upload_file(self, file_path: Path, retry_count: int = 0) -> Tuple[bool, str]:
        """
        上传单个文件
        
        Args:
            file_path: 文件路径
            retry_count: 当前重试次数
            
        Returns:
            (是否成功, 错误信息)
        """
        file_name = file_path.name
        
        try:
            # 准备请求头
            headers = {
                'token': self.token
            }
            
            # 获取Content-Type
            content_type = self._get_content_type(file_path)
            
            # 准备文件数据
            with open(file_path, 'rb') as f:
                files = {
                    'file': (file_name, f, content_type)
                }
                
                # 发送POST请求
                response = requests.post(
                    self.upload_url,
                    headers=headers,
                    files=files,
                    timeout=60
                )
            
            # 检查响应
            if response.status_code == 200:
                try:
                    result = response.json()
                    if result.get('code') == 200:
                        logger.info(f"✓ 上传成功: {file_name}")
                        return True, ""
                    else:
                        error_msg = result.get('msg', '未知错误')
                        logger.error(f"✗ 上传失败: {file_name} - {error_msg}")
                        return False, error_msg
                except ValueError:
                    # 响应不是JSON格式，但状态码是200
                    logger.warning(f"上传响应不是JSON格式: {file_name}")
                    return True, ""
            else:
                error_msg = f"HTTP {response.status_code}: {response.text[:100]}"
                logger.error(f"✗ 上传失败: {file_name} - {error_msg}")
                return False, error_msg
                
        except requests.exceptions.Timeout:
            error_msg = "请求超时"
            logger.error(f"✗ 上传失败: {file_name} - {error_msg}")
            return False, error_msg
        except requests.exceptions.RequestException as e:
            error_msg = f"请求异常: {str(e)}"
            logger.error(f"✗ 上传失败: {file_name} - {error_msg}")
            return False, error_msg
        except Exception as e:
            error_msg = f"未知错误: {str(e)}"
            logger.error(f"✗ 上传失败: {file_name} - {error_msg}", exc_info=True)
            return False, error_msg
    
    def upload_file_with_retry(self, file_path: Path) -> bool:
        """
        上传文件（带重试机制）
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否上传成功
        """
        file_name = file_path.name
        
        for attempt in range(self.max_retries + 1):
            success, error_msg = self.upload_file(file_path, retry_count=attempt)
            
            if success:
                # 上传成功，删除文件
                try:
                    file_path.unlink()
                    logger.info(f"✓ 已删除文件: {file_name}")
                except Exception as e:
                    logger.warning(f"删除文件失败: {file_name} - {e}")
                return True
            
            # 如果还有重试机会，等待后重试
            if attempt < self.max_retries:
                wait_time = (attempt + 1) * 2  # 递增等待时间：2秒、4秒、6秒...
                logger.info(f"⏳ {file_name} 将在 {wait_time} 秒后重试 (第 {attempt + 1}/{self.max_retries} 次)")
                time.sleep(wait_time)
        
        # 所有重试都失败
        with self.stats_lock:
            self.stats['failed_files'].append({
                'file': file_name,
                'error': error_msg
            })
        return False
    
    def scan_files(self) -> List[Path]:
        """
        递归扫描待上传的文件（包括所有子目录）
        
        Returns:
            文件路径列表
        """
        # 支持常见的Excel和文档格式
        extensions = ['.xlsx', '.xls', '.docx', '.doc', '.pdf']
        files = []
        
        # 使用 rglob 递归扫描所有子目录
        for ext in extensions:
            files.extend(self.upload_dir.rglob(f"*{ext}"))
        
        # 只保留文件（排除目录）
        files = [f for f in files if f.is_file()]
        
        # 按文件路径排序
        files.sort(key=lambda p: str(p))
        
        return files
    
    def print_header(self):
        """打印启动信息"""
        print("=" * 60)
        print("批量并行上传文件工具")
        print("=" * 60)
        print(f"上传目录: {self.upload_dir}")
        print(f"上传URL: {self.upload_url}")
        print(f"并发数: {self.max_workers}")
        print(f"最大重试次数: {self.max_retries}")
        print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        print()
    
    def print_summary(self):
        """打印统计信息"""
        duration = self.stats['end_time'] - self.stats['start_time']
        minutes = int(duration // 60)
        seconds = int(duration % 60)
        duration_str = f"{minutes}分{seconds}秒" if minutes > 0 else f"{seconds}秒"
        
        print()
        print("=" * 60)
        print("上传完成统计")
        print("=" * 60)
        print(f"总文件数: {self.stats['total']}")
        print(f"成功: {self.stats['success']}")
        print(f"失败: {self.stats['failed']}")
        print(f"耗时: {duration_str}")
        
        if self.stats['failed_files']:
            print()
            print("失败文件列表:")
            for item in self.stats['failed_files']:
                print(f"  - {item['file']}: {item['error']}")
        
        print("=" * 60)
    
    def run(self):
        """运行上传程序"""
        try:
            # 打印启动信息
            self.print_header()
            
            # 扫描文件
            print("正在扫描待上传文件...")
            files = self.scan_files()
            
            if not files:
                print("未找到任何待上传文件！")
                return
            
            print(f"找到 {len(files)} 个文件")
            print()
            
            # 初始化统计信息
            self.stats['total'] = len(files)
            self.stats['start_time'] = time.time()
            self.stats['success'] = 0
            self.stats['failed'] = 0
            self.stats['failed_files'] = []
            
            # 使用线程池并发上传文件
            print(f"开始并发上传（并发数: {self.max_workers}）...")
            print()
            
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # 提交所有任务
                future_to_file = {
                    executor.submit(self.upload_file_with_retry, file_path): file_path
                    for file_path in files
                }
                
                # 处理完成的任务
                completed = 0
                for future in as_completed(future_to_file):
                    file_path = future_to_file[future]
                    completed += 1
                    
                    try:
                        success = future.result()
                        
                        # 线程安全地更新统计
                        with self.stats_lock:
                            if success:
                                self.stats['success'] += 1
                            else:
                                self.stats['failed'] += 1
                        
                        # 打印进度
                        progress = f"[{completed}/{self.stats['total']}]"
                        status = "✓" if success else "✗"
                        print(f"{progress} {status} {file_path.name}")
                        
                    except Exception as e:
                        logger.error(f"处理文件时出错: {file_path.name}, 错误: {e}", exc_info=True)
                        with self.stats_lock:
                            self.stats['failed'] += 1
                            self.stats['failed_files'].append({
                                'file': file_path.name,
                                'error': str(e)
                            })
                        print(f"[{completed}/{self.stats['total']}] ✗ {file_path.name} - 处理异常")
            
            # 记录结束时间
            self.stats['end_time'] = time.time()
            
            # 打印统计信息
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
    uploader = FileUploader()
    uploader.run()


if __name__ == "__main__":
    main()

