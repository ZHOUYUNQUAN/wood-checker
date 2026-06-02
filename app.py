from flask import Flask, redirect, url_for

from config import Config
from models import init_db

app = Flask(__name__)
app.config.from_object(Config)


# 反向代理支持（与服务器一致）
class ReverseProxied:
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        script_name = environ.get('HTTP_X_SCRIPT_NAME', '')
        if script_name:
            environ['SCRIPT_NAME'] = script_name
            path_info = environ.get('PATH_INFO', '')
            if path_info.startswith(script_name):
                environ['PATH_INFO'] = path_info[len(script_name):]
        return self.wsgi_app(environ, start_response)


app.wsgi_app = ReverseProxied(app.wsgi_app)

# 注册检尺蓝图
from check import check_bp
app.register_blueprint(check_bp, url_prefix='/check')


def ensure_runtime():
    """确保运行所需目录和数据库表存在。"""
    import os
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    init_db()


ensure_runtime()


@app.route('/')
def index():
    return redirect(url_for('check.index'))


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5050)
