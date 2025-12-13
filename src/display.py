"""显示模块 - 使用rich库提供实时显示界面"""
import sys
from typing import List, Optional
from datetime import datetime
from pathlib import Path
from rich.console import Console
from rich.live import Live
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn
from rich.table import Table
from rich.layout import Layout
from rich.text import Text


class Display:
    """实时显示管理器 - 使用rich库实现类似top命令的显示界面"""
    
    def __init__(self, log_lines: int = 10):
        """
        初始化显示管理器
        
        Args:
            log_lines: 日志显示行数（默认10行）
        """
        self.log_lines = log_lines
        self.log_buffer: List[str] = []
        self.console = Console()
        self.live: Optional[Live] = None
        self.current_file = ""
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'index': 0,
            'start_time': None,
        }
    
    def add_log(self, message: str, is_error: bool = False):
        """
        添加日志消息
        
        Args:
            message: 日志消息
            is_error: 是否为错误/失败消息（将显示为黄色）
        """
        timestamp = datetime.now().strftime('%H:%M:%S')
        
        # 限制单条日志的最大长度
        max_message_length = 70
        if len(message) > max_message_length:
            message = message[:max_message_length - 3] + "..."
        
        # 构建日志条目
        log_entry = f"[{timestamp}] {message}"
        if is_error:
            log_entry = f"[{timestamp}] [yellow]{message}[/yellow]"
        
        self.log_buffer.append(log_entry)
        
        # 保持日志缓冲区大小
        if len(self.log_buffer) > self.log_lines:
            self.log_buffer = self.log_buffer[-self.log_lines:]
    
    def _create_layout(self) -> Layout:
        """创建布局"""
        layout = Layout()
        
        # 创建系统名称区域（最顶部）
        title_text = Text()
        title_text.append("文档自动提取关键信息系统", style="bold blue")
        title_text.append(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", style="dim")
        title_panel = Panel(title_text, border_style="blue", padding=(0, 1))
        
        # 创建顶部状态区域
        header_text = self._create_header()
        header_panel = Panel(header_text, border_style="blue")
        
        # 创建当前处理文件区域
        file_text = self._create_file_info()
        file_panel = Panel(file_text, title="当前处理", border_style="green")
        
        # 创建日志区域
        log_text = self._create_log_display()
        log_panel = Panel(log_text, title="处理日志", border_style="cyan")
        
        # 垂直布局
        # 系统名称区域（最顶部）+ 状态区域 + 文件区域 + 日志区域
        layout.split_column(
            Layout(title_panel, size=3, minimum_size=3),
            Layout(header_panel, size=6, minimum_size=6),
            Layout(file_panel, size=6, minimum_size=6),
            Layout(log_panel)
        )
        
        return layout
    
    def _create_header(self) -> Text:
        """创建顶部状态信息"""
        text = Text()
        
        if self.stats['start_time']:
            elapsed = (datetime.now() - self.stats['start_time']).total_seconds()
            elapsed_str = f"{int(elapsed // 60)}分{int(elapsed % 60)}秒" if elapsed >= 60 else f"{int(elapsed)}秒"
            
            if elapsed > 0 and self.stats['index'] > 0:
                speed = self.stats['index'] / elapsed * 60
                remaining = (self.stats['total'] - self.stats['index']) / (self.stats['index'] / elapsed) if self.stats['index'] > 0 else 0
                remaining_str = f"{int(remaining // 60)}分{int(remaining % 60)}秒" if remaining >= 60 else f"{int(remaining)}秒"
                speed_str = f"{speed:.1f} 文件/分钟 | 剩余: {remaining_str}"
            else:
                speed_str = "计算中..."
        else:
            elapsed_str = "0秒"
            speed_str = "计算中..."
        
        # 进度条
        if self.stats['total'] > 0:
            progress_pct = int((self.stats['index'] / self.stats['total']) * 100)
            progress_bar_length = 50
            filled = int(progress_bar_length * self.stats['index'] / self.stats['total'])
            bar = '█' * filled + '░' * (progress_bar_length - filled)
            progress_info = f"[{bar}] {progress_pct}% ({self.stats['index']}/{self.stats['total']})"
        else:
            progress_info = "等待开始..."
        
        text.append(f"进度: {progress_info}\n", style="bold")
        text.append(f"统计: 成功 [green]{self.stats['success']}[/green] | 失败 [red]{self.stats['failed']}[/red] | 耗时 {elapsed_str} | {speed_str}\n")
        text.append(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", style="dim")
        
        return text
    
    def _create_file_info(self) -> Text:
        """创建当前处理文件信息"""
        text = Text()
        if self.current_file:
            # 换行显示长文件名
            import textwrap
            wrapped_lines = textwrap.wrap(self.current_file, width=70, break_long_words=False, break_on_hyphens=False)
            for line in wrapped_lines[:3]:  # 最多显示3行
                text.append(f"  {line}\n")
            if len(wrapped_lines) > 3:
                text.append(f"  {wrapped_lines[2][:67]}...\n", style="dim")
        else:
            text.append("  等待处理...", style="dim")
        return text
    
    def _create_log_display(self) -> Text:
        """创建日志显示"""
        text = Text()
        recent_logs = self.log_buffer[-self.log_lines:]
        for log in recent_logs:
            # rich支持markup语法，直接添加
            text.append(log + "\n")
        
        # 填充剩余行
        remaining_lines = self.log_lines - len(recent_logs)
        for _ in range(remaining_lines):
            text.append("\n")
        
        return text
    
    def render(self, header_lines: List[str], current_file: str = "", progress: str = ""):
        """
        渲染显示界面
        
        Args:
            header_lines: 顶部状态行列表（保留兼容性，实际不使用）
            current_file: 当前处理的文件
            progress: 进度信息（保留兼容性，实际不使用）
        """
        self.current_file = current_file
        
        if self.live:
            layout = self._create_layout()
            self.live.update(layout)
    
    def init_display(self):
        """初始化显示"""
        layout = self._create_layout()
        self.live = Live(layout, console=self.console, refresh_per_second=4)
        self.live.start()
    
    def show_summary(self, summary_data: dict):
        """
        显示最终统计信息
        
        Args:
            summary_data: 包含统计信息的字典
                - total: 总文件数
                - success: 成功数
                - failed: 失败数
                - duration_str: 耗时字符串
                - avg_speed: 平均速度
                - failed_files: 失败文件列表
                - output_files: 输出文件信息列表
        """
        if self.live:
            self.live.stop()
        
        # 创建总结布局
        layout = Layout()
        
        # 创建统计信息面板
        stats_text = Text()
        stats_text.append("处理完成！\n\n", style="bold green")
        stats_text.append(f"总文件数: {summary_data.get('total', 0)}\n")
        stats_text.append(f"成功处理: [green]{summary_data.get('success', 0)}[/green]\n")
        stats_text.append(f"失败: [red]{summary_data.get('failed', 0)}[/red]\n")
        stats_text.append(f"总耗时: {summary_data.get('duration_str', '0秒')}\n")
        stats_text.append(f"平均速度: {summary_data.get('avg_speed', 0):.1f} 文件/分钟\n")
        
        stats_panel = Panel(stats_text, title="统计信息", border_style="green")
        
        # 创建失败文件表格
        failed_files = summary_data.get('failed_files', [])
        if failed_files:
            failed_table = Table(show_header=True, header_style="bold red", border_style="red", expand=True)
            failed_table.add_column("后缀", style="cyan", width=8, no_wrap=True, min_width=8, max_width=8)
            failed_table.add_column("文件名", style="yellow", min_width=30, overflow="fold")
            failed_table.add_column("错误原因", style="red", min_width=20, overflow="fold")
            
            for item in failed_files:
                file_name = item['file']
                file_ext = Path(file_name).suffix if file_name else ""
                error_reason = item.get('simplified_error', item.get('error', ''))
                
                failed_table.add_row(file_ext, file_name, error_reason)
            
            failed_panel = Panel(failed_table, title="失败文件", border_style="red", expand=True)
        else:
            failed_panel = None
        
        # 创建输出文件信息
        output_files = summary_data.get('output_files', [])
        if output_files:
            output_text = Text()
            total_records = summary_data.get('total_output_records', 0)
            success_count = summary_data.get('success', 0)
            
            for file_info in output_files:
                output_text.append(f"  - {file_info['name']}: {file_info['records']}条记录\n")
            
            output_text.append(f"\n输出文件总记录数: {total_records}条\n")
            output_text.append(f"处理成功文件数: {success_count}个\n")
            
            if total_records != success_count:
                output_text.append(f"⚠ 注意: 记录数({total_records})与成功文件数({success_count})不一致", style="yellow")
            
            output_panel = Panel(output_text, title="输出文件", border_style="blue")
        else:
            output_panel = None
        
        # 组合布局
        panels = [Layout(stats_panel, size=8)]
        if failed_panel:
            panels.append(Layout(failed_panel))
        if output_panel:
            panels.append(Layout(output_panel))
        
        if len(panels) > 1:
            layout.split_column(*panels)
        else:
            layout = panels[0]
        
        self.console.print(layout)
        self.console.print()  # 换行
    
    def cleanup_display(self):
        """清理显示"""
        if self.live:
            self.live.stop()
    
    def update_stats(self, **kwargs):
        """更新统计信息"""
        for key, value in kwargs.items():
            if key == 'start_time':
                if isinstance(value, datetime):
                    self.stats[key] = value
                elif value is not None:
                    # 如果是时间戳，转换为datetime
                    self.stats[key] = datetime.fromtimestamp(value)
                else:
                    self.stats[key] = None
            else:
                self.stats[key] = value
