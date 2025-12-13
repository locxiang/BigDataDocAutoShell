"""显示模块 - 使用Textual库提供实时显示界面"""
import sys
import threading
from typing import List, Optional
from datetime import datetime
from pathlib import Path
from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, ScrollableContainer, Grid
from textual.widgets import Static, DataTable, Header, Footer, Label
from textual import events
from textual.reactive import reactive
from textual.screen import Screen


class ProcessingApp(App):
    """处理界面应用"""
    
    TITLE = "文档自动提取关键信息系统"
    
    CSS = """
    Screen {
        background: $surface;
    }
    
    #main-container {
        layout: vertical;
    }
    
    #title-bar {
        height: 3;
        border: solid $primary;
        background: $primary 20%;
        text-align: center;
        padding: 1;
        width: 100%;
    }
    
    #title {
        text-align: center;
        width: 100%;
        text-style: bold;
    }
    
    #stats-container {
        height: auto;
        min-height: 6;
        border: solid $primary;
        padding: 1;
        layout: vertical;
    }
    
    #stats-container > Static.stat-label {
        height: 1;
        text-style: bold;
        color: $primary;
    }
    
    #progress-bar {
        width: 100%;
        height: auto;
        min-height: 1;
        margin: 1 0;
        text-align: left;
    }
    
    #stats-text {
        width: 100%;
        height: auto;
        min-height: 1;
        margin-top: 1;
    }
    
    #current-file-container {
        height: 7;
        border: solid $success;
        padding: 1;
        layout: vertical;
    }
    
    #current-file-container > Static.stat-label {
        height: 1;
        text-style: bold;
        color: $success;
    }
    
    #current-file-name {
        width: 100%;
        height: 2;
        margin-top: 1;
        text-style: bold;
    }
    
    #current-file-status {
        width: 100%;
        height: 1;
        margin-top: 1;
        color: $accent;
    }
    
    #log-container {
        border: solid $accent;
        padding: 1;
        layout: vertical;
    }
    
    #log-container > Static.stat-label {
        height: 1;
        text-style: bold;
        color: $accent;
    }
    
    #log-content {
        width: 100%;
        margin-top: 1;
    }
    
    .stat-label {
        text-style: bold;
        color: $text;
        width: 100%;
    }
    
    .stat-value {
        color: $success;
    }
    
    .stat-error {
        color: $error;
    }
    
    .current-file-name {
        text-style: bold;
        color: $text;
    }
    
    .current-file-status {
        color: $accent;
    }
    """
    
    # 响应式数据
    progress_percent = reactive(0.0)
    success_count = reactive(0)
    failed_count = reactive(0)
    speed = reactive(0.0)
    elapsed_time = reactive("0秒")
    current_file = reactive("")
    current_status = reactive("")
    
    def __init__(self, log_lines: int = 20):
        super().__init__()
        self.log_lines = log_lines
        self.log_buffer: List[str] = []
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'index': 0,
            'start_time': None,
        }
        self._lock = threading.Lock()
    
    def compose(self) -> ComposeResult:
        """组合界面"""
        yield Header(show_clock=True)
        yield Footer()
        
        with Vertical(id="main-container"):
            # 标题栏
            with Container(id="title-bar"):
                yield Static("文档自动提取关键信息系统", id="title")
            
            # 统计面板（单独一行）
            with Container(id="stats-container"):
                yield Static("📊 统计信息", classes="stat-label")
                yield Static("[░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░] 0%", id="progress-bar")  # 文本进度条
                yield Static("等待开始...", id="stats-text")
            
            # 当前处理文件
            with Container(id="current-file-container"):
                yield Static("📄 当前处理文件", classes="stat-label")
                yield Static("等待处理...", id="current-file-name", classes="current-file-name")
                yield Static("", id="current-file-status", classes="current-file-status")
            
            # 日志区域
            with ScrollableContainer(id="log-container"):
                yield Static("📋 处理日志", classes="stat-label")
                yield Static("暂无日志", id="log-content")
    
    def on_mount(self) -> None:
        """挂载时初始化"""
        # 初始化进度条显示（确保进度条有初始值）
        try:
            progress_bar = self.query_one("#progress-bar", Static)
            if progress_bar:
                progress_bar.update("[░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░] 0%")
            else:
                # 如果找不到进度条，等待一下再试
                self.set_timer(0.1, self._init_progress_bar)
        except Exception:
            pass
        
        # 初始化显示
        self.update_display()
        
        # 设置定时刷新，确保 UI 实时更新（每0.5秒更新一次统计信息和进度）
        self.set_interval(0.5, self._refresh_display)
    
    def _init_progress_bar(self) -> None:
        """延迟初始化进度条"""
        try:
            progress_bar = self.query_one("#progress-bar", Static)
            if progress_bar:
                progress_bar.update("[░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░] 0%")
        except Exception:
            pass
    
    def action_quit(self) -> None:
        """处理退出操作（Ctrl+C 等）"""
        self.exit()
    
    def on_key(self, event: events.Key) -> None:
        """处理键盘事件"""
        # 允许 Ctrl+C 退出
        if event.key == "ctrl+c":
            self.exit()
        # 其他按键不做处理，让 Textual 正常处理
    
    def _refresh_display(self) -> None:
        """定时刷新显示"""
        try:
            # 更新统计信息（包括重新计算耗时和速度）
            total = self.stats.get('total', 0)
            index = self.stats.get('index', 0)
            start_time = self.stats.get('start_time')
            
            if start_time and total > 0:
                # 重新计算耗时和速度
                elapsed = (datetime.now() - start_time).total_seconds()
                if elapsed >= 60:
                    self.elapsed_time = f"{int(elapsed // 60)}分{int(elapsed % 60)}秒"
                else:
                    self.elapsed_time = f"{int(elapsed)}秒"
                
                if elapsed > 0 and index > 0:
                    self.speed = index / elapsed * 60
                else:
                    self.speed = 0.0
                
                # 更新进度（这会触发 watch_progress_percent）
                self.progress_percent = (index / total) * 100
            
            # 更新统计信息显示和进度条
            self.update_stats_display()
            self._update_progress_bar()
        except Exception:
            pass
    
    def watch_progress_percent(self, progress: float) -> None:
        """监听进度变化"""
        self._update_progress_bar()
    
    def watch_success_count(self, count: int) -> None:
        """监听成功数变化"""
        self.update_stats_display()
    
    def watch_failed_count(self, count: int) -> None:
        """监听失败数变化"""
        self.update_stats_display()
    
    def watch_speed(self, speed: float) -> None:
        """监听速度变化"""
        self.update_stats_display()
    
    def watch_elapsed_time(self, time_str: str) -> None:
        """监听耗时变化"""
        self.update_stats_display()
    
    def watch_current_file(self, file_name: str) -> None:
        """监听当前文件变化"""
        try:
            file_widget = self.query_one("#current-file-name", Static)
            if file_widget:
                display_name = file_name if file_name else "等待处理..."
                file_widget.update(display_name)
        except Exception:
            pass
    
    def watch_current_status(self, status: str) -> None:
        """监听状态变化"""
        try:
            status_widget = self.query_one("#current-file-status", Static)
            if status_widget:
                status_widget.update(status if status else "")
        except Exception:
            pass
    
    def update_stats_display(self) -> None:
        """更新统计信息显示"""
        try:
            stats_text = self.query_one("#stats-text", Static)
            if not stats_text:
                # 调试：记录找不到元素的情况
                return
            
            total = self.stats.get('total', 0)
            index = self.stats.get('index', 0)
            
            if total > 0:
                progress_pct = int((index / total) * 100)
                # Textual 不支持 rich markup，使用纯文本
                stats_display = (
                    f"进度: {progress_pct}% ({index}/{total})  |  "
                    f"成功: {self.success_count}  |  "
                    f"失败: {self.failed_count}  |  "
                    f"速度: {self.speed:.1f} 文件/分钟  |  "
                    f"耗时: {self.elapsed_time}"
                )
            else:
                stats_display = "等待开始..."
            
            stats_text.update(stats_display)
        except Exception as e:
            # 调试：记录异常以便排查问题
            import traceback
            # 只在开发时输出，避免影响用户体验
            pass
    
    def update_display(self) -> None:
        """更新整个显示"""
        self.update_stats_display()
        self.update_log_display()
    
    def update_log_display(self) -> None:
        """更新日志显示"""
        try:
            log_widget = self.query_one("#log-content", Static)
            with self._lock:
                recent_logs = self.log_buffer[-self.log_lines:]
            
            log_text = "\n".join(recent_logs) if recent_logs else "暂无日志"
            log_widget.update(log_text)
        except Exception:
            pass
    
    def add_log(self, message: str, is_error: bool = False) -> None:
        """添加日志消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        
        # 限制单条日志的最大长度
        max_message_length = 80
        if len(message) > max_message_length:
            message = message[:max_message_length - 3] + "..."
        
        # 构建日志条目（Textual 不支持 rich markup，使用纯文本）
        log_entry = f"[{timestamp}] {message}"
        
        with self._lock:
            self.log_buffer.append(log_entry)
            # 保持日志缓冲区大小
            if len(self.log_buffer) > self.log_lines * 2:  # 保留更多历史
                self.log_buffer = self.log_buffer[-self.log_lines * 2:]
        
        # 异步更新显示
        try:
            self.call_from_thread(self.update_log_display)
        except Exception:
            # 如果在同一线程，直接调用
            self.update_log_display()
    
    def update_stats(self, **kwargs) -> None:
        """更新统计信息"""
        # 更新 stats 字典
        for key, value in kwargs.items():
            if key == 'start_time':
                if isinstance(value, datetime):
                    self.stats[key] = value
                elif value is not None:
                    self.stats[key] = datetime.fromtimestamp(value)
                else:
                    self.stats[key] = None
            else:
                self.stats[key] = value
        
        # 计算并更新响应式属性（这会触发 watch 方法）
        total = self.stats.get('total', 0)
        index = self.stats.get('index', 0)
        success = self.stats.get('success', 0)
        failed = self.stats.get('failed', 0)
        start_time = self.stats.get('start_time')
        
        # 更新响应式属性（这会自动触发 watch 方法）
        if total > 0:
            self.progress_percent = (index / total) * 100
        else:
            self.progress_percent = 0.0
        
        self.success_count = success
        self.failed_count = failed
        
        # 计算耗时和速度
        if start_time:
            elapsed = (datetime.now() - start_time).total_seconds()
            if elapsed >= 60:
                self.elapsed_time = f"{int(elapsed // 60)}分{int(elapsed % 60)}秒"
            else:
                self.elapsed_time = f"{int(elapsed)}秒"
            
            if elapsed > 0 and index > 0:
                self.speed = index / elapsed * 60
            else:
                self.speed = 0.0
        else:
            self.elapsed_time = "0秒"
            self.speed = 0.0
        
        # 手动触发 UI 更新（确保立即显示）
        # 注意：这个方法应该在主线程中调用，如果从后台线程调用，应该使用 call_from_thread
        try:
            self.update_stats_display()
            self._update_progress_bar()
        except Exception:
            # 如果更新失败，静默处理（可能是元素还未挂载）
            pass
    
    def _update_progress_bar(self) -> None:
        """更新文本进度条"""
        try:
            progress_bar = self.query_one("#progress-bar", Static)
            if not progress_bar:
                # 调试：如果找不到进度条元素，返回
                return
            
            progress_value = max(0.0, min(100.0, self.progress_percent))
            progress_pct = int(progress_value)
            bar_length = 50
            filled = int(bar_length * progress_value / 100)
            bar = '█' * filled + '░' * (bar_length - filled)
            progress_text = f"[{bar}] {progress_pct}%"
            progress_bar.update(progress_text)
        except Exception as e:
            # 调试：记录异常以便排查问题
            pass


class SummaryScreen(Screen):
    """总结信息界面"""
    
    CSS = """
    Screen {
        background: $surface;
    }
    
    #main-container {
        layout: vertical;
        padding: 1;
    }
    
    #stats-container {
        height: auto;
        border: solid $success;
        padding: 1;
        margin-bottom: 1;
    }
    
    #stats-container > Static.stat-title {
        text-style: bold;
        color: $success;
        margin-bottom: 1;
    }
    
    #failed-container {
        border: solid $error;
        padding: 1;
        margin-bottom: 1;
    }
    
    #failed-container > Static.stat-title {
        text-style: bold;
        color: $error;
        margin-bottom: 1;
    }
    
    #output-container {
        border: solid $primary;
        padding: 1;
        margin-bottom: 1;
    }
    
    #output-container > Static.stat-title {
        text-style: bold;
        color: $primary;
        margin-bottom: 1;
    }
    
    #footer-hint {
        text-align: center;
        padding: 1;
        text-style: bold;
        color: $warning;
    }
    
    DataTable {
        height: auto;
        max-height: 20;
    }
    """
    
    def __init__(self, summary_data: dict):
        super().__init__()
        self.summary_data = summary_data
    
    def compose(self) -> ComposeResult:
        """组合总结界面"""
        yield Header(show_clock=False)
        
        with Vertical(id="main-container"):
            # 统计信息
            with Container(id="stats-container"):
                yield Static("📊 统计信息", classes="stat-title")
                yield Static(self._format_stats(), id="stats-content")
            
            # 失败文件
            failed_files = self.summary_data.get('failed_files', [])
            if failed_files:
                with Container(id="failed-container"):
                    yield Static("❌ 失败文件", classes="stat-title")
                    failed_table = DataTable(id="failed-table")
                    failed_table.add_columns("后缀", "文件名", "错误原因")
                    for item in failed_files:
                        file_name = item['file']
                        file_ext = Path(file_name).suffix if file_name else ""
                        error_reason = item.get('simplified_error', item.get('error', ''))
                        failed_table.add_row(file_ext, file_name, error_reason)
                    yield failed_table
            
            # 输出文件
            output_files = self.summary_data.get('output_files', [])
            if output_files:
                with Container(id="output-container"):
                    yield Static("📁 输出文件", classes="stat-title")
                    yield Static(self._format_output_files(), id="output-content")
            
            # 提示信息
            yield Static("按 Q 键退出程序", id="footer-hint")
        
        yield Footer()
    
    def _format_stats(self) -> str:
        """格式化统计信息"""
        total = self.summary_data.get('total', 0)
        success = self.summary_data.get('success', 0)
        failed = self.summary_data.get('failed', 0)
        duration = self.summary_data.get('duration_str', '0秒')
        avg_speed = self.summary_data.get('avg_speed', 0)
        
        return (
            f"处理完成！\n\n"
            f"总文件数: {total}\n"
            f"成功处理: {success}\n"
            f"失败: {failed}\n"
            f"总耗时: {duration}\n"
            f"平均速度: {avg_speed:.1f} 文件/分钟"
        )
    
    def _format_output_files(self) -> str:
        """格式化输出文件信息"""
        output_files = self.summary_data.get('output_files', [])
        total_records = self.summary_data.get('total_output_records', 0)
        success_count = self.summary_data.get('success', 0)
        
        lines = []
        for file_info in output_files:
            lines.append(f"  - {file_info['name']}: {file_info['records']}条记录")
        
        lines.append(f"\n输出文件总记录数: {total_records}条")
        lines.append(f"处理成功文件数: {success_count}个")
        
        if total_records != success_count:
            lines.append(f"⚠ 注意: 记录数({total_records})与成功文件数({success_count})不一致")
        
        return "\n".join(lines)
    
    def on_key(self, event: events.Key) -> None:
        """处理键盘事件"""
        if event.key == "q" or event.key == "Q":
            self.app.exit()
        elif event.key == "ctrl+c":
            self.app.exit()


class Display:
    """实时显示管理器 - 使用Textual库实现TUI界面"""
    
    def __init__(self, log_lines: int = 20):
        """
        初始化显示管理器
        
        Args:
            log_lines: 日志显示行数（默认20行）
        """
        self.log_lines = log_lines
        self.app: Optional[ProcessingApp] = None
        self.current_file = ""
        self.stats = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'index': 0,
            'start_time': None,
        }
    
    def init_display(self):
        """初始化显示 - 返回 App 实例，需要在主线程调用 run()"""
        self.app = ProcessingApp(log_lines=self.log_lines)
        return self.app
    
    def add_log(self, message: str, is_error: bool = False):
        """
        添加日志消息
        
        Args:
            message: 日志消息
            is_error: 是否为错误/失败消息（将显示为黄色）
        """
        if self.app:
            try:
                # 尝试从其他线程调用
                if hasattr(self.app, 'call_from_thread'):
                    self.app.call_from_thread(self.app.add_log, message, is_error)
                else:
                    # 如果在同一线程，直接调用
                    self.app.add_log(message, is_error)
            except Exception:
                # 如果失败，直接调用（可能在主线程）
                try:
                    self.app.add_log(message, is_error)
                except Exception:
                    pass
    
    def render(self, header_lines: List[str], current_file: str = "", progress: str = ""):
        """
        渲染显示界面
        
        Args:
            header_lines: 顶部状态行列表（保留兼容性，实际不使用）
            current_file: 当前处理的文件
            progress: 进度信息（状态信息）
        """
        self.current_file = current_file
        if self.app:
            try:
                # 更新当前文件显示
                if hasattr(self.app, 'call_from_thread'):
                    self.app.call_from_thread(setattr, self.app, "current_file", current_file)
                    if progress:
                        self.app.call_from_thread(setattr, self.app, "current_status", progress)
                else:
                    # 如果在主线程，直接设置
                    self.app.current_file = current_file
                    if progress:
                        self.app.current_status = progress
            except Exception:
                pass
    
    def update_stats(self, **kwargs):
        """更新统计信息"""
        # 更新 Display 的 stats（用于兼容性）
        for key, value in kwargs.items():
            if key == 'start_time':
                if isinstance(value, datetime):
                    self.stats[key] = value
                elif value is not None:
                    self.stats[key] = datetime.fromtimestamp(value)
                else:
                    self.stats[key] = None
            else:
                self.stats[key] = value
        
        # 更新 ProcessingApp 的 stats（实际显示）
        if self.app:
            try:
                # 尝试从其他线程调用
                if hasattr(self.app, 'call_from_thread'):
                    # 使用 call_from_thread 确保线程安全
                    self.app.call_from_thread(self.app.update_stats, **kwargs)
                else:
                    # 如果在主线程，直接调用
                    self.app.update_stats(**kwargs)
            except Exception as e:
                # 如果失败，尝试直接调用（可能在主线程）
                try:
                    self.app.update_stats(**kwargs)
                except Exception:
                    # 如果还是失败，记录但不中断程序
                    pass
    
    def cleanup_display(self):
        """清理显示"""
        if self.app:
            try:
                if hasattr(self.app, 'exit'):
                    self.app.exit()
            except Exception:
                pass
    
    def show_summary(self, summary_data: dict):
        """
        显示最终统计信息（使用 Textual UI）
        
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
        if self.app:
            try:
                # 切换到总结界面
                summary_screen = SummaryScreen(summary_data)
                self.app.push_screen(summary_screen)
            except Exception as e:
                # 如果失败，尝试直接退出
                try:
                    self.app.exit()
                except Exception:
                    pass
