"""
硬件项目管理系统 - 主入口
"""
if __name__ == '__main__':
    import sys
    import os

    # 设置环境变量
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    os.environ['APP_BASE_DIR'] = BASE_DIR

    # 创建 Flask 应用并启动
    from server import app

    print('[INFO] 硬件项目管理系统已启动')
    print('[INFO] 访问地址: http://127.0.0.1:5000')
    print('[INFO] 按 Ctrl+C 停止服务器')
    print()

    # 直接用 Flask 的 run 方法启动
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)
