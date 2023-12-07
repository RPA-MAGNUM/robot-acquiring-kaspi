import ctypes
import sys
from pathlib import Path

from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning

from logs import init_logger
from rpamini import json_read, net_use, get_hostname

disable_warnings(InsecureRequestWarning)

root_path = Path(__file__).parent

local_path = Path.home().joinpath(f'AppData\\Local\\.rpa')
local_env_path = local_path.joinpath('env.json')
local_env_data = json_read(local_env_path)

global_path = Path(local_env_data['global_path'])
global_username = local_env_data['global_username']
global_password = local_env_data['global_password']
net_use(global_path, global_username, global_password)
global_env_path = global_path.joinpath('env.json')
global_env_data = json_read(global_env_path)

orc_host = global_env_data['orc_host']
tg_token = global_env_data['tg_token']
smtp_host = global_env_data['smtp_host']
smtp_author = global_env_data['smtp_author']
sprut_username = global_env_data['sprut_username']
sprut_password = global_env_data['sprut_password']
sprut_username_personal = global_env_data['sprut_username_personal']
sprut_password_personal = global_env_data['sprut_password_personal']
odines_username = global_env_data['odines_username']
odines_password = global_env_data['odines_password']
odines_username_rpa = global_env_data['odines_username_rpa']
odines_password_rpa = global_env_data['odines_password_rpa']
owa_username = global_env_data['owa_username']
owa_password = global_env_data['owa_password']

sed_username = global_env_data['sed_username']
sed_password = global_env_data['sed_password']
cups_host = global_env_data['cups_host']
cups_username = global_env_data['cups_username']
cups_password = global_env_data['cups_password']
cas_username = global_env_data['cas_username']
cas_password = global_env_data['cas_password']

db_host = global_env_data['postgre_ip']

db_port = global_env_data['postgre_port']
db_name = global_env_data['postgre_db_name']
db_schema = 'robot'
db_user = global_env_data['postgre_db_username']
db_pass = global_env_data['postgre_db_password']

# * Edit from here
robot_name = "robot-acquiring-kaspi"
robot_name_russian = "Робот Эквайринг Каспи"

temp_folder = local_path.joinpath(f".agent\\{robot_name}\\temp")
temp_folder.mkdir(exist_ok=True, parents=True)

config_path = global_path.joinpath(f'.agent\\{robot_name}\\{get_hostname()}\\config.json')

sprut_base = json_read(config_path)["sprut_base"]

common_network_folder = Path(json_read(config_path)["common_network_folder"])
net_use(common_network_folder, owa_username, owa_password)


main_directory_folder = json_read(config_path)["main_directory_folder"]
str_path_mapping_excel_file = json_read(config_path)["str_path_mapping_excel_file"]

str_date_working_file = json_read(config_path)['str_date_working_file']


process_list_path = local_path.joinpath('process_list.json')
months = ['', 'январь', 'февраль', 'март', 'апрель', 'май',
          'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

months_for_folders = ['', '01. Январь', '02. Февраль', '03. Март', '04. Апрель', '05. Май',
                      '06. Июнь', '07. Июль', '08. Август', '09. Сентябрь', '10. Октябрь', '11. Ноябрь', '12. Декабрь']
upload_timeout_minutes = 120
ip_address = get_hostname()
upload_folder = local_path.joinpath(f".agent\\{robot_name}\\upload")
upload_folder.mkdir(exist_ok=True, parents=True)

screenshots_folder = global_path.joinpath(f".agent\\{robot_name}\\screenshots")
screenshots_folder.mkdir(exist_ok=True, parents=True)

log_path = global_path.joinpath(f".agent/{robot_name}/{ip_address}")
log_path.mkdir(exist_ok=True, parents=True)
log_path = log_path.joinpath(f'{sys.argv[1]}.log' if len(sys.argv) > 1 else "log.log")
logger = init_logger(tg_token=tg_token, chat_id='-1001905447645', log_path=log_path)

str_parking_folder = json_read(config_path)['str_parking_folder']
str_sales_folder = json_read(config_path)['str_sales_folder']
to_whom = json_read(config_path)['to_whom']
cc_whom = json_read(config_path)['cc_whom']
bool_use_prod_1c = True

transaction_retry_count = 2
if ctypes.windll.user32.GetKeyboardLayout(0) != 67699721:
    __err__ = 'Смените раскладку на ENG'
    raise Exception(__err__)
