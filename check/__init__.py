from flask import Blueprint

check_bp = Blueprint('check', __name__, template_folder='../templates/check', static_folder='../static')

from . import views