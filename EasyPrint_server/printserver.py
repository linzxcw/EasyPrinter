import subprocess
import sys
import time
import os
from win32 import win32print
import yaml


def print_action(filename):
    if sys.platform == 'win32':
        args = [f"{os.path.dirname(__file__)}\plug\PDFPrinter.exe",
                    f"{filename}",
                    f"{win32print.GetDefaultPrinter ()}",
                    ]
        subprocess.run(args, encoding="utf-8", shell=True)

print("轻松打印客户端正在启动，请稍后......")
time.sleep(2)

with open('./config.yaml', 'r', encoding='utf8') as file:  # utf8可识别中文
    config = yaml.safe_load(file)
folder_path = config['path0']
print("服务端共享文件夹路径是：" + folder_path)
delete_pending = {}

print("运行中......")
print("*请把需要打印的pdf文件发送到服务端的共享文件夹中")
while True:
    if len(os.listdir(folder_path)) > 0:
        for file in os.listdir(folder_path):
            if file not in delete_pending.keys():
                delete_pending[file] = time.time()
                print_action(os.path.join(folder_path, file))
            else:
                if time.time() - delete_pending[file] > 20:
                    os.remove(os.path.join(folder_path, file))
                    del delete_pending[file]
    time.sleep(1)

