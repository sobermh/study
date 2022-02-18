"""
@autor : maohui
@time  : 2022/2/17 23:04
"""

import requests


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

if __name__=="__main__":
    TestRequests().login_get_token()
