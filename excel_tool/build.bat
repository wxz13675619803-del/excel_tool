@echo off
chcp 65001
echo ========================================
echo   Excel智能处理工具 - 打包脚本
echo ========================================
echo.

REM 检查Python环境
python --version
if errorlevel 1 (
    echo [错误] 未找到Python环境，请先安装Python
    pause
    exit /b
)

REM 安装依赖
echo [1/3] 安装依赖...
pip install -r requirements.txt -q

REM 创建启动脚本
echo [2/3] 创建启动器...
(
echo import subprocess
echo import sys
echo import os
echo.
echo def main():
echo     # 获取当前目录
echo     if getattr(sys, 'frozen', False):
echo         app_dir = os.path.dirname(sys.executable)
echo     else:
echo         app_dir = os.path.dirname(os.path.abspath(__file__))
echo.
echo     app_path = os.path.join(app_dir, 'app.py')
echo.
echo     # 启动streamlit
echo     subprocess.run([
echo         sys.executable, '-m', 'streamlit', 'run', app_path,
echo         '--server.headless', 'true',
echo         '--browser.gatherUsageStats', 'false',
echo         '--theme.primaryColor', '#667eea',
echo         '--theme.backgroundColor', '#ffffff',
echo         '--theme.secondaryBackgroundColor', '#f0f2f6',
echo         '--server.port', '8501',
echo     ])
echo.
echo if __name__ == '__main__':
echo     main()
) > launcher.py

REM 方案A：简单打包（推荐先试这个）
echo [3/3] 开始打包...
echo.
echo 推荐方案：直接发送文件夹（最稳定）
echo.

REM 创建一键启动脚本
(
echo @echo off
echo chcp 65001
echo echo 正在启动 Excel智能处理工具...
echo echo 请稍候，浏览器将自动打开...
echo echo.
echo echo 提示：关闭此窗口即可退出程序
echo echo ========================================
echo cd /d "%%~dp0"
echo python -m streamlit run app.py --server.headless true --browser.gatherUsageStats false --theme.primaryColor "#667eea" --server.port 8501
echo pause
) > 启动工具.bat

echo.
echo ========================================
echo   打包完成！
echo ========================================
echo.
echo 发送给别人的方式（推荐方案）：
echo.
echo 1. 将整个文件夹打成zip发送
echo 2. 对方电脑需要安装Python
echo 3. 对方先运行: pip install -r requirements.txt
echo 4. 然后双击"启动工具.bat"即可
echo.
echo 如果对方没有Python环境，请使用以下命令打包exe：
echo   pyinstaller --onefile --noconsole launcher.py
echo.
pause