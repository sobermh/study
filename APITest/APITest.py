"""
@autor : maohui
@time  : 2022/2/17 23:04
"""

import requests

# 定义一个全局变量；类变量。通过类名调用
token = ""


class TestRequests():
    def login_get_token(self):
        url = "https://clinms.top/api/login"
        data = {
            "appname": "LungScr",
            "password": "LungScr"
        }
        res = requests.post(url=url, data=data)
        print(res.json())
        # data:可以传纯键值对的dict（非嵌套的dict），也可以穿str格式
        # json：可以传任何形式的dict（包括嵌套的dict）
        global token
        token = res.json()['token']

    def add_sample(self):
        global token
        url = "https://clinms.top/api/sample"
        data = {
            "sys_id": 0,
            "type": 0,
            "cid": "",
            "collect_time": "2020-12-22 13:42:00",
             "channel": "",
             "risk":"",
             "result": "",
             "guidance": ""
        }
        headers={
            "Authorization": "Bearer<" + token + ">"
        }
        res = requests.post(url, json=data,headers=headers)
        print(res.status_code)


if __name__ == "__main__":
    TestRequests().login_get_token()
    print(token)
    TestRequests().add_sample()


#不管是get还是post还是put和delete，都是调用requests.request方法。
#requests.requests方法调用的是session.reques方法

def setup(self):
    print("在每个用例之前执行一次：初始化日志对象，初始化数据库连接")
def teardowm(self):
    print("在每个用例之后执行一次；关闭日志对象，初始化数据库连接")

def setup_class(self):
    print("在每个类之前执行")
def teardowm_class(self):
    print("在每个类之前执行")