"""政策问答对生成入口程序"""
import sys
import logging
import threading
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from src.config import validate_config, OUTPUT_DIR, MAX_WORKERS
from src.qa_generator import QAGenerator

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('qa_generation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def main():
    """主函数"""
    try:
        # 验证配置
        validate_config()
        
        # 政策文件Excel路径
        policy_excel = OUTPUT_DIR / "4政策文件信息.xlsx"
        
        if not policy_excel.exists():
            logger.error(f"政策文件Excel不存在: {policy_excel}")
            print(f"\n错误: 找不到政策文件Excel: {policy_excel}")
            print("请先运行 main.py 处理文档，生成政策文件信息。")
            sys.exit(1)
        
        print("=" * 60)
        print("政策问答对生成系统")
        print("=" * 60)
        print(f"政策文件: {policy_excel.name}")
        print(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        print()
        
        # 初始化生成器
        generator = QAGenerator()
        
        # 读取政策文件
        print("正在读取政策文件...")
        policies = generator.read_policy_file(policy_excel)
        
        if not policies:
            print("未找到任何政策记录！")
            return
        
        print(f"共找到 {len(policies)} 条政策记录")
        print(f"并发数: {MAX_WORKERS}")
        print()
        
        # 统计信息（主进程计数器，线程安全）
        total_policies = len(policies)
        stats = {
            'success': 0,
            'failed': 0,
            'total_qa_pairs': 0,
        }
        stats_lock = threading.Lock()
        
        def process_policy(policy: dict, index: int) -> tuple[bool, int]:
            """
            处理单个政策（用于并发执行）
            
            Args:
                policy: 政策信息字典
                index: 政策索引（用于显示）
                
            Returns:
                (是否成功, 生成的问答对数量)
            """
            policy_id = policy.get("ID", "unknown")
            policy_name = policy.get("Remarks", policy.get("PolicyFileName", "未知政策"))
            
            try:
                # 生成问答对
                qa_pairs = generator.generate_qa_pairs(policy)
                
                if qa_pairs:
                    # 保存问答对（线程安全）
                    success = generator.save_qa_pairs(qa_pairs, policy)
                    
                    if success:
                        logger.info(f"[{index}/{total_policies}] 成功处理政策: {policy_name} (ID: {policy_id}), 生成 {len(qa_pairs)} 个问答对")
                        return True, len(qa_pairs)
                    else:
                        logger.error(f"[{index}/{total_policies}] 保存失败: {policy_name} (ID: {policy_id})")
                        return False, 0
                else:
                    logger.error(f"[{index}/{total_policies}] 生成问答对失败: {policy_name} (ID: {policy_id})")
                    return False, 0
                
            except Exception as e:
                logger.error(f"[{index}/{total_policies}] 处理政策失败: {policy_id}, 错误: {e}", exc_info=True)
                return False, 0
        
        # 使用线程池并发处理政策
        print("开始并发处理政策文件...")
        print()
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            # 提交所有任务
            future_to_policy = {
                executor.submit(process_policy, policy, idx): (policy, idx)
                for idx, policy in enumerate(policies, 1)
            }
            
            # 处理完成的任务
            for future in as_completed(future_to_policy):
                policy, index = future_to_policy[future]
                policy_id = policy.get("ID", "unknown")
                policy_name = policy.get("Remarks", policy.get("PolicyFileName", "未知政策"))
                
                try:
                    success, qa_count = future.result()
                    
                    # 线程安全地更新主进程的计数器
                    with stats_lock:
                        if success:
                            stats['success'] += 1
                            stats['total_qa_pairs'] += qa_count
                            print(f"[{stats['success'] + stats['failed']}/{total_policies}] ✓ {policy_name} - 生成 {qa_count} 个问答对")
                        else:
                            stats['failed'] += 1
                            print(f"[{stats['success'] + stats['failed']}/{total_policies}] ✗ {policy_name} - 处理失败")
                    
                except Exception as e:
                    logger.error(f"处理政策时出错: {policy_id}, 错误: {e}", exc_info=True)
                    with stats_lock:
                        stats['failed'] += 1
                    print(f"[{stats['success'] + stats['failed']}/{total_policies}] ✗ {policy_name} - 处理异常")
        
        # 获取最终统计信息
        success_count = stats['success']
        failed_count = stats['failed']
        total_qa_pairs = stats['total_qa_pairs']
        
        # 输出统计信息
        print("=" * 60)
        print("处理完成！")
        print("=" * 60)
        print(f"总政策数: {total_policies}")
        print(f"成功处理: {success_count}")
        print(f"失败: {failed_count}")
        print(f"生成问答对总数: {total_qa_pairs}")
        print(f"结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 60)
        
        # 输出文件位置
        output_file = OUTPUT_DIR / "5政策问答对.xlsx"
        if output_file.exists():
            print(f"\n问答对已保存到: {output_file}")
        
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        sys.exit(0)
    except Exception as e:
        logger.error(f"程序运行失败: {e}", exc_info=True)
        print(f"\n错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()

