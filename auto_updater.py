import requests
import json
import os
import sys
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, Any

# 当前版本号
CURRENT_VERSION = "v1.1.0"


class AutoUpdater:
    """
    自动更新模块，使用Gitee API检查和下载更新
    """

    def __init__(self, owner: str = "xianyuwangyou", repo: str = "TianDanAssistant"):
        """
        初始化自动更新器

        Args:
            owner: Gitee仓库所有者
            repo: Gitee仓库名称
        """
        self.owner = owner
        self.repo = repo
        self.api_base_url = f"https://gitee.com/api/v5/repos/{owner}/{repo}"
        self.current_version = CURRENT_VERSION

    def get_latest_release(self) -> Optional[Dict[Any, Any]]:
        """
        获取最新的发布版本信息

        Returns:
            包含最新版本信息的字典，如果出错则返回None
        """
        try:
            url = f"{self.api_base_url}/releases/latest"
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            if "403" in str(e):
                print("API请求频率超限，请稍后再试")
            else:
                print(f"获取最新版本信息失败: {e}")
            return None

    def check_for_updates(self) -> tuple[bool, Optional[str], Optional[Dict[Any, Any]]]:
        """
        检查是否有可用更新

        Returns:
            一个元组 (has_update, latest_version, release_info)
            has_update: 是否有更新
            latest_version: 最新版本号
            release_info: 最新版本信息
        """
        release_info = self.get_latest_release()
        if not release_info:
            return False, None, None

        latest_version = release_info.get("tag_name", "")
        
        # 简单的版本比较（实际项目中可能需要更复杂的版本比较逻辑）
        has_update = latest_version > self.current_version
        
        return has_update, latest_version, release_info

    def download_latest_version(self, download_dir: str = "download") -> bool:
        """
        下载最新版本的可执行文件

        Args:
            download_dir: 下载目录

        Returns:
            是否下载成功
        """
        print("正在检查最新版本信息...")
        release_info = self.get_latest_release()
        if not release_info:
            print("获取版本信息失败")
            return False

        try:
            # 查找可执行文件的下载链接
            assets = release_info.get("assets", [])
            exe_asset = None
            for asset in assets:
                if asset.get("name", "").endswith(".exe"):
                    exe_asset = asset
                    break

            if not exe_asset:
                print("未找到可执行文件")
                return False

            download_url = exe_asset.get("browser_download_url", "")
            filename = exe_asset.get("name", "latest_version.exe")
            print(f"找到可执行文件: {filename}")

            # 确保下载目录存在
            if not os.path.exists(download_dir):
                os.makedirs(download_dir)

            # 下载文件
            print(f"开始下载: {download_url}")
            response = requests.get(download_url, timeout=30, stream=True)
            response.raise_for_status()

            file_path = os.path.join(download_dir, filename)
            
            # 获取文件总大小
            total_size = int(response.headers.get('content-length', 0))
            if total_size == 0:
                print("无法确定文件大小")
            
            downloaded_size = 0
            chunk_size = 8192
            with open(file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=chunk_size):
                    if chunk:
                        f.write(chunk)
                        downloaded_size += len(chunk)
                        # 显示下载进度
                        if total_size > 0:
                            percent = (downloaded_size / total_size) * 100
                            progress_msg = f"下载进度: {percent:.1f}% ({downloaded_size}/{total_size} bytes)"
                            print(f"\r{progress_msg}", end='', flush=True)
                        else:
                            progress_msg = f"已下载: {downloaded_size} bytes"
                            print(f"\r{progress_msg}", end='', flush=True)
            
            print(f"\n文件已下载到: {file_path}")
            return True
        except Exception as e:
            print(f"\n下载失败: {e}")
            return False

    def show_update_dialog(self, parent: tk.Tk, latest_version: str, release_info: Dict[Any, Any]) -> None:
        """
        显示更新对话框

        Args:
            parent: 父窗口
            latest_version: 最新版本号
            release_info: 最新版本信息
        """
        release_notes = release_info.get("body", "暂无更新说明")
        update_title = f"发现新版本 {latest_version}"

        # 创建更新对话框
        dialog = tk.Toplevel(parent)
        dialog.title("软件更新")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        # 居中显示对话框
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

         # 创建主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(main_frame, text=update_title, font=("微软雅黑", 12, "bold"))
        title_label.pack(pady=(0, 10))

        # 更新说明标签
        notes_label = ttk.Label(main_frame, text="更新说明:")
        notes_label.pack(anchor="w", padx=5)

        # 创建文本框显示更新说明
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True, pady=(5, 10))

        text_box = tk.Text(text_frame, wrap="word", height=8)
        text_box.insert("1.0", release_notes)
        text_box.config(state="disabled")

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_box.yview)
        scrollbar.pack(side="right", fill="y")
        text_box.pack(side="left", fill="both", expand=True)

        text_box.config(yscrollcommand=scrollbar.set)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        def download_update():
            dialog.destroy()
            # 触发实际的下载逻辑
            # 在新线程中执行下载，避免阻塞UI
            import threading
            download_thread = threading.Thread(target=self._download_update_in_background)
            download_thread.daemon = True
            download_thread.start()

        def close_dialog():
            dialog.destroy()

        # 下载按钮
        download_btn = ttk.Button(button_frame, text="前往下载", command=download_update, width=10)
        download_btn.pack(side="left", padx=10)

        # 关闭按钮
        close_btn = ttk.Button(button_frame, text="下次再说", command=close_dialog, width=10)
        close_btn.pack(side="left", padx=10)
    def _download_update_in_background(self):
        """
        在后台下载更新
        """
        try:
            # 创建下载进度对话框
            progress_dialog = tk.Toplevel()
            progress_dialog.title("下载更新")
            progress_dialog.geometry("350x100")
            progress_dialog.resizable(False, False)
            progress_dialog.transient()
            progress_dialog.grab_set()
            # 设置隐藏
            progress_dialog.withdraw()
            
            # 居中显示对话框
            progress_dialog.update_idletasks()
            x = (progress_dialog.winfo_screenwidth() // 2) - (progress_dialog.winfo_width() // 2)
            y = (progress_dialog.winfo_screenheight() // 2) - (progress_dialog.winfo_height() // 2)
            progress_dialog.geometry(f"+{x}+{y}")
            # 居中后显示窗口
            progress_dialog.deiconify()

            # 进度条框架
            progress_frame = ttk.Frame(progress_dialog, padding="10")
            progress_frame.pack(fill=tk.BOTH, expand=True)
            
            # 状态标签
            status_label = ttk.Label(progress_frame, text="正在准备下载...")
            status_label.pack(pady=(0, 5))
            
            # 进度条
            progress_bar = ttk.Progressbar(progress_frame, mode="indeterminate")
            progress_bar.pack(fill="x", pady=(0, 5))
            progress_bar.start()
            
            # 状态栏（单行）
            status_bar = ttk.Label(progress_dialog, text="准备下载...", relief=tk.SUNKEN, anchor=tk.W)
            status_bar.pack(side=tk.BOTTOM, fill=tk.X)
            
            # 定义更新状态栏的函数
            def update_status(message):
                # 检查窗口是否仍然存在
                if progress_dialog and progress_dialog.winfo_exists():
                    try:
                        status_bar.config(text=message)
                        progress_dialog.update_idletasks()
                    except tk.TclError:
                        # 窗口已被销毁，忽略更新
                        pass
            
            # 重写下载函数以显示状态
            def download_with_status(download_dir: str = "download"):
                """
                带状态显示的下载函数
                """
                update_status("正在检查最新版本信息...")
                release_info = self.get_latest_release()
                if not release_info:
                    update_status("获取版本信息失败")
                    return False

                try:
                    # 查找可执行文件的下载链接
                    assets = release_info.get("assets", [])
                    exe_asset = None
                    for asset in assets:
                        if asset.get("name", "").endswith(".exe"):
                            exe_asset = asset
                            break

                    if not exe_asset:
                        update_status("未找到可执行文件")
                        return False

                    download_url = exe_asset.get("browser_download_url", "")
                    filename = exe_asset.get("name", "latest_version.exe")
                    update_status(f"找到可执行文件: {filename}")

                    # 确保下载目录存在
                    if not os.path.exists(download_dir):
                        os.makedirs(download_dir)

                    # 下载文件
                    update_status(f"开始下载: {download_url}")
                    response = requests.get(download_url, timeout=30, stream=True)
                    response.raise_for_status()

                    file_path = os.path.join(download_dir, filename)
                    
                    # 获取文件总大小
                    total_size = int(response.headers.get('content-length', 0))
                    if total_size == 0:
                        update_status("无法确定文件大小")
                    
                    downloaded_size = 0
                    chunk_size = 8192
                    with open(file_path, "wb") as f:
                        for chunk in response.iter_content(chunk_size=chunk_size):
                            if chunk:
                                f.write(chunk)
                                downloaded_size += len(chunk)
                                # 显示下载进度
                                if total_size > 0:
                                    percent = (downloaded_size / total_size) * 100
                                    progress_msg = f"下载进度: {percent:.1f}% ({downloaded_size}/{total_size} bytes)"
                                    if downloaded_size % (chunk_size * 100) == 0:  # 每100个块更新一次状态
                                        # 检查窗口是否仍然存在
                                        if progress_dialog and progress_dialog.winfo_exists():
                                            update_status(progress_msg)
                                        else:
                                            # 窗口已关闭，退出下载
                                            return False
                                else:
                                    progress_msg = f"已下载: {downloaded_size} bytes"
                                    if downloaded_size % (chunk_size * 100) == 0:  # 每100个块更新一次状态
                                        # 检查窗口是否仍然存在
                                        if progress_dialog and progress_dialog.winfo_exists():
                                            update_status(progress_msg)
                                        else:
                                            # 窗口已关闭，退出下载
                                            return False
                    
                    update_status(f"文件已下载到: {file_path}")
                    return True
                except Exception as e:
                    update_status(f"下载失败: {e}")
                    return False
            
            # 下载文件
            success = download_with_status()
            
            # 关闭进度对话框
            if progress_dialog and progress_dialog.winfo_exists():
                progress_dialog.destroy()
            
            # 显示结果
            if success:
                messagebox.showinfo("下载完成", "新版本已下载完成，请检查 download 文件夹下的可执行文件。")
            else:
                if progress_dialog and progress_dialog.winfo_exists():
                    messagebox.showerror("下载失败", "下载新版本时出错，请稍后重试或手动下载。")
                
        except Exception as e:
            # 只有在窗口仍存在时才显示错误消息
            try:
                if 'progress_dialog' in locals() and progress_dialog and progress_dialog.winfo_exists():
                    messagebox.showerror("下载出错", f"下载过程中发生错误: {str(e)}")
            except:
                pass
        
    def _update_progress(self, dialog, label, progress_bar, text, value, max_value):
        """
        更新进度对话框
        """
        if dialog and dialog.winfo_exists():
            label.config(text=text)
            if max_value > 0:
                progress_bar.config(value=value, maximum=max_value)
            dialog.update_idletasks()

def check_updates_background(parent: tk.Tk) -> None:
    """
    在后台检查更新并在有更新时通知用户

    Args:
        parent: 父窗口
    """
    updater = AutoUpdater()
    has_update, latest_version, release_info = updater.check_for_updates()

    if has_update and latest_version and release_info:
        # 在主线程中显示更新对话框
        parent.after(0, lambda: updater.show_update_dialog(parent, latest_version, release_info))


def main():
    """
    主函数 - 用于测试自动更新功能
    """
    updater = AutoUpdater()
    has_update, latest_version, release_info = updater.check_for_updates()

    if has_update:
        print(f"发现新版本: {latest_version}")
        print(f"当前版本: {updater.current_version}")
        release_notes = release_info.get("body", "暂无更新说明") if release_info else "暂无更新说明"
        print(f"更新说明:\n{release_notes}")
    else:
        print("当前已是最新版本")


if __name__ == "__main__":
    main()