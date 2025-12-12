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
        self.config_file = Path.home() / ".stock_viewer_config.json"
        
        # 默认设置
        self.default_width = 400
        self.default_height = 600
        self.crop_top = 182
        self.crop_bottom = 88
        
        # 加载保存的设置
        self.load_settings()
        
        # 创建主窗口
        self.create_window()
        
    def load_settings(self):
        """加载保存的设置"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                self.default_width = config.get('width', self.default_width)
                self.default_height = config.get('height', self.default_height)
                self.crop_top = config.get('crop_top', self.crop_top)
                self.crop_bottom = config.get('crop_bottom', self.crop_bottom)
                
            except Exception as e:
                print(f"加载配置失败: {e}")
                
    def save_settings(self, width=None, height=None):
        """保存设置"""
        try:
            # 使用提供的宽度和高度，或使用默认值
            config = {
                'width': width if width is not None else self.default_width,
                'height': height if height is not None else self.default_height,
                'crop_top': self.crop_top,
                'crop_bottom': self.crop_bottom
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
            
    def remove_window_buttons(self):
        """移除窗口的最大化和最小化按钮（Windows API）"""
        try:
            # 获取窗口句柄
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
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
        try:
            # 获取窗口句柄
            hwnd = ctypes.windll.user32.GetForegroundWindow()
            if hwnd:
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
            
    def on_loaded(self):
        """页面加载完成事件"""
        # 注入CSS
        self.inject_css()
        # 移除窗口按钮（必须在窗口创建后调用）
        time.sleep(0.5)  # 等待窗口完全创建
        self.remove_window_buttons()
        # 窗口居中
        self.center_window()
        
    def create_window(self):
        """创建主窗口"""
        # 创建有边框窗口，以便可以通过边缘调整大小
        self.window = webview.create_window(
            title='股票行情查看器',
            url='https://i.jzj9999.com/quoteh5/',
            width=self.default_width,
            height=self.default_height,
            resizable=True,
            fullscreen=False,
            min_size=(300, 400),
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
        self.save_settings()
        sys.exit(0)
        
    def on_resized(self, width, height):
        """窗口调整大小事件"""
        self.default_width = width
        self.default_height = height
        self.save_settings(width, height)
        
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
