#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股票行情查看器最终版 - 在窗体内嵌显示网页
使用pywebview实现，Python 3.13兼容

功能要求：
1. 无最大化最小化按钮的窗体（通过Windows API移除）
2. 默认宽400px高600px，可调节大小并保存设置
3. 无系统托盘图标，窗口在任务栏显示
4. 网页顶部182px和底部88px不显示（通过CSS注入实现）

依赖安装：
pip install pywebview

注意：此版本使用Windows API修改窗口样式，仅支持Windows系统。
"""

import sys
import json
import time
import threading
from pathlib import Path
import webview
import ctypes
import ctypes.wintypes

# Windows API常量
GWL_STYLE = -16
WS_MAXIMIZEBOX = 0x00010000
WS_MINIMIZEBOX = 0x00020000
WS_THICKFRAME = 0x00040000  # 可调整大小的边框

class StockViewer:
    def __init__(self):
        self.window = None
        
        # 默认设置
        self.default_width = 330
        self.default_height = 600
        self.crop_top = 182
        self.crop_bottom = 88
        
        # 窗体收起相关变量
        self.is_collapsed = False
        self.collapsed_height = 8  # 收起后的高度（只显示标题栏）
        self.original_height = self.default_height  # 保存收起前的原始高度
        self.collapse_threshold = 10  # 距离顶部多少像素触发收起
        self.monitor_thread = None
        self.monitor_running = False
        self.window_hwnd = None  # 保存窗口句柄
        
        # 创建主窗口
        self.create_window()
        
            
    def remove_window_buttons(self):
        """移除窗口的最大化和最小化按钮（Windows API）"""
        if not self.window_hwnd:
            return
        try:
            # 使用保存的窗口句柄
            hwnd = self.window_hwnd
            # 获取当前窗口样式
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_STYLE)
            # 移除最大化最小化按钮，但保留可调整大小的边框
            style = style & ~WS_MAXIMIZEBOX & ~WS_MINIMIZEBOX
            # 确保有可调整大小的边框
            style = style | WS_THICKFRAME
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_STYLE, style)
            # 刷新窗口
            ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0002 | 0x0001)
            print("窗口按钮已移除")
        except Exception as e:
            print(f"修改窗口样式失败: {e}")
            
    def center_window(self):
        """将窗口居中显示（Windows API）"""
        if not self.window_hwnd:
            return
        try:
            # 使用保存的窗口句柄
            hwnd = self.window_hwnd
            # 获取屏幕尺寸
            screen_width = ctypes.windll.user32.GetSystemMetrics(0)
            screen_height = ctypes.windll.user32.GetSystemMetrics(1)
            
            # 获取窗口尺寸
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            window_width = rect.right - rect.left
            window_height = rect.bottom - rect.top
            
            # 计算居中位置
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            
            # 设置窗口位置
            SWP_NOSIZE = 0x0001
            SWP_NOZORDER = 0x0004
            ctypes.windll.user32.SetWindowPos(
                hwnd, 0, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER
            )
            print("窗口已居中")
        except Exception as e:
            print(f"窗口居中失败: {e}")
            
    def find_my_window(self):
        """查找并返回窗口句柄"""
        try:
            # 通过标题枚举所有窗口
            def enum_windows_callback(hwnd, lparam):
                try:
                    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                    if length > 0:
                        buffer = ctypes.create_unicode_buffer(length + 1)
                        ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
                        title = buffer.value
                        
                        # 只检查标题是否匹配
                        if title == '股票行情查看器':
                            # 将hwnd值写入lparam指向的位置
                            ctypes.cast(lparam, ctypes.POINTER(ctypes.wintypes.HWND)).contents.value = hwnd
                            print(f"找到窗口: 标题={title}, hwnd={hwnd}")
                            return False  # 停止枚举
                except:
                    pass
                return True  # 继续枚举
            
            # 等待窗口创建
            for retry in range(10):
                time.sleep(0.3)
                
                hwnd_found = ctypes.wintypes.HWND(0)
                WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
                ctypes.windll.user32.EnumWindows(WNDENUMPROC(enum_windows_callback), ctypes.byref(hwnd_found))
                
                if hwnd_found.value != 0:
                    self.window_hwnd = hwnd_found.value
                    print(f"成功获取窗口句柄: {hwnd_found.value}")
                    return hwnd_found.value
                else:
                    print(f"重试 {retry + 1}/10: 未找到窗口")
            
            print("警告: 10次尝试后仍未找到窗口，监控功能将无法工作")
            print("提示：窗体收起功能可能无法正常使用")
                
        except Exception as e:
            print(f"查找窗口失败: {e}")
            import traceback
            traceback.print_exc()
        return None
        
    def on_loaded(self):
        """页面加载完成事件"""
        # 注入CSS
        self.inject_css()
        # 移除窗口按钮（必须在窗口创建后调用）
        time.sleep(0.5)  # 等待窗口完全创建
        # 获取并保存窗口句柄
        self.find_my_window()
        if self.window_hwnd:
            self.remove_window_buttons()
            self.center_window()
        # 启动窗口监控线程
        self.start_monitor()
        
    def create_window(self):
        """创建主窗口"""
        # 创建有边框窗口，以便可以通过边缘调整大小
        # 注意：不设置min_size，或者设置一个很小的值，以允许窗口收起
        self.window = webview.create_window(
            title='股票行情查看器',
            url='https://i.jzj9999.com/quoteh5/',
            width=self.default_width,
            height=self.default_height,
            resizable=True,
            fullscreen=False,
            min_size=(300, 50),  # 允许最小高度50px，以便收起
            background_color='#FFFFFF',
            frameless=False,  # 有边框，以便调整大小
            easy_drag=False,  # 不使用自定义拖动
            on_top=False,
            confirm_close=False  # 关闭时不显示确认对话框
        )
        
        # 设置窗口事件
        self.window.events.closed += self.on_closed
        self.window.events.resized += self.on_resized
        self.window.events.loaded += self.on_loaded
        
    def on_closed(self):
        """窗口关闭事件"""
        self.monitor_running = False
        sys.exit(0)
        
    def on_resized(self, width, height):
        """窗口调整大小事件"""
        self.default_width = width
        self.default_height = height
        # 只在未收起且高度大于100px时才更新原始高度
        # 避免收起后的小高度覆盖原始高度
        if not self.is_collapsed and height > 100:
            self.original_height = height
        
    def inject_css(self):
        """注入CSS来裁剪网页"""
        css = f"""
        <style>
            /* 裁剪网页 */
            body {{
                margin-top: -{self.crop_top}px !important;
                margin-bottom: -{self.crop_bottom}px !important;
                overflow: hidden !important;
            }}
            html {{
                overflow: hidden !important;
            }}
            
            /* 隐藏可能出现的固定元素 */
            .fixed-top, .header, .top-bar {{
                display: none !important;
            }}
            .fixed-bottom, .footer, .bottom-bar {{
                display: none !important;
            }}
            
            /* 确保内容区域可滚动（如果需要） */
            .main-content {{
                height: 100vh;
                overflow-y: auto;
            }}
        </style>
        """
        
        js = f"""
        (function() {{
            // 注入CSS
            let style = document.createElement('style');
            style.textContent = `{css}`;
            document.head.appendChild(style);
            
            // 给所有非body元素添加main-content类
            let bodyChildren = document.body.children;
            for (let i = 0; i < bodyChildren.length; i++) {{
                bodyChildren[i].classList.add('main-content');
            }}
            
            console.log('CSS注入完成');
        }})();
        """
        
        try:
            self.window.evaluate_js(js)
            print("CSS注入成功")
        except Exception as e:
            print(f"CSS注入失败: {e}")
            # 重试一次
            time.sleep(1)
            try:
                self.window.evaluate_js(js)
            except:
                pass
                
    def get_window_position(self):
        """获取窗口位置（Windows API）"""
        if not self.window_hwnd:
            return None
        try:
            # 使用保存的窗口句柄
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(self.window_hwnd, ctypes.byref(rect))
            return rect.left, rect.top, rect.right, rect.bottom
        except:
            pass
        return None
        
    def set_window_height(self, height):
        """设置窗口高度（Windows API）"""
        if not self.window_hwnd:
            return False
        try:
            # 使用保存的窗口句柄
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(self.window_hwnd, ctypes.byref(rect))
            window_width = rect.right - rect.left
            old_height = rect.bottom - rect.top
            
            print(f"设置窗口高度: 当前={old_height}, 目标={height}")
            
            # 保持窗口位置不变，只改变高度
            SWP_NOMOVE = 0x0002
            SWP_NOZORDER = 0x0004
            SWP_FRAMECHANGED = 0x0020
            
            result = ctypes.windll.user32.SetWindowPos(
                self.window_hwnd, 0, rect.left, rect.top, window_width, height,
                SWP_NOMOVE | SWP_NOZORDER | SWP_FRAMECHANGED
            )
            
            # 等待并验证高度是否真的改变了
            time.sleep(0.1)
            new_rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(self.window_hwnd, ctypes.byref(new_rect))
            new_height = new_rect.bottom - new_rect.top
            print(f"设置后窗口高度: {new_height}, 结果={result}")
            
            return True
        except Exception as e:
            print(f"设置窗口高度失败: {e}")
            import traceback
            traceback.print_exc()
        return False
        
    def collapse_window(self):
        """收起窗口"""
        if not self.is_collapsed:
            pos = self.get_window_position()
            if pos:
                # 保存当前高度（确保不是收起后的高度）
                current_height = pos[3] - pos[1]
                if current_height > 100:  # 只有大于100px才认为是真正的原始高度
                    self.original_height = current_height
                    print(f"保存原始高度: {self.original_height}")
                else:
                    print(f"当前高度{current_height}太小，使用原始高度: {self.original_height}")
                    
                # 设置收起高度（尝试不同的高度值）
                success = False
                for h in [50]:  # 由于最小高度限制，只能设置到50px
                    if self.set_window_height(h):
                        success = True
                        break
                
                if success:
                    self.is_collapsed = True
                    print(f"窗口已收起到{h}px（受最小高度限制）")
                    
    def expand_window(self):
        """展开窗口"""
        if self.is_collapsed:
            if self.set_window_height(self.original_height):
                self.is_collapsed = False
                print("窗口已展开")
                
    def is_mouse_over_window(self):
        """检测鼠标是否在窗口上（整个窗口区域）"""
        if not self.window_hwnd:
            return False
        try:
            # 使用保存的窗口句柄
            # 获取鼠标位置
            point = ctypes.wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
            
            # 获取窗口矩形
            rect = ctypes.wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(self.window_hwnd, ctypes.byref(rect))
            
            # 检查鼠标是否在窗口的任何位置（整个窗口）
            mouse_in_window = (rect.left <= point.x <= rect.right and
                             rect.top <= point.y <= rect.bottom + 10)  # 整个窗口+底部10px缓冲
            return mouse_in_window
        except:
            pass
        return False
        
    def start_monitor(self):
        """启动窗口监控线程"""
        self.monitor_running = True
        self.monitor_thread = threading.Thread(target=self.monitor_window, daemon=True)
        self.monitor_thread.start()
        
    def monitor_window(self):
        """监控窗口位置状态"""
        print("监控线程已启动")
        last_y = None
        is_dragging = False
        collapse_cooldown = 0
        expand_cooldown = 0  # 添加展开冷却，避免闪烁
        
        while self.monitor_running:
            try:
                time.sleep(0.1)  # 每100ms检查一次
                
                pos = self.get_window_position()
                if pos is None:
                    print("无法获取窗口位置")
                    continue
                    
                current_y = pos[1]
                current_height = pos[3] - pos[1]
                
                # 每隔一段时间打印当前位置（调试用）
                if int(time.time() * 10) % 50 == 0:  # 每5秒打印一次
                    print(f"窗口位置: Y={current_y}, 高度={current_height}, 收起状态={self.is_collapsed}")
                
                # 检测拖拽状态
                if last_y is not None:
                    if current_y != last_y:
                        is_dragging = True
                        collapse_cooldown = 5  # 拖拽时等待5个周期后再检测收起
                    else:
                        is_dragging = False
                
                last_y = current_y
                
                # 减少冷却计数
                if collapse_cooldown > 0:
                    collapse_cooldown -= 1
                if expand_cooldown > 0:
                    expand_cooldown -= 1
                
                # 检测是否应该收起（只在非收起状态且不在拖拽时）
                if (not self.is_collapsed and 
                    not is_dragging and 
                    collapse_cooldown == 0 and
                    current_y <= self.collapse_threshold and
                    current_height > self.collapsed_height):
                    self.collapse_window()
                    collapse_cooldown = 10  # 收起后冷却
                    
                # 检测是否应该展开
                # 情况1：鼠标在收起的窗口上
                # 情况2：窗口离开顶部区域（拖动下来）
                if (self.is_collapsed and expand_cooldown == 0):
                    should_expand = False
                    
                    # 检查鼠标是否在窗口上
                    if self.is_mouse_over_window():
                        should_expand = True
                        print("鼠标在窗口上，准备展开")
                    
                    # 检查窗口是否离开顶部区域（拖动下来）
                    elif current_y > self.collapse_threshold + 20:
                        should_expand = True
                        print(f"窗口离开顶部(Y={current_y})，准备展开")
                    
                    if should_expand:
                        self.expand_window()
                        expand_cooldown = 30  # 展开后等待3秒，避免立即收起
                        collapse_cooldown = 20
                    
            except Exception as e:
                print(f"监控线程错误: {e}")
                time.sleep(0.5)
                
    def run(self):
        """运行应用程序"""
        webview.start(debug=False)

def main():
    """主函数"""
    # 检查必要的模块
    try:
        import webview
    except ImportError as e:
        print("缺少必要的模块，请安装：")
        print("pip install pywebview")
        print(f"错误详情: {e}")
        sys.exit(1)
        
    # 检查操作系统
    import platform
    if platform.system() != 'Windows':
        print("警告：此版本专为Windows系统设计")
        print("在其他系统上可能无法移除窗口按钮")
        
    # 运行应用程序
    app = StockViewer()
    app.run()
    
if __name__ == "__main__":
    main()
