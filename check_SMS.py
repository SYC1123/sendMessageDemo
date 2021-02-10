# python3.6.5
# 需要引入requests包 ：运行终端->进入python/Scripts ->输入：pip install requests
from ShowapiRequest import ShowapiRequest

# "showapi_res_body": {"ret_code": "0", "remark": "提交成功,请等待管理员审核!联系电话:4009988033 qq:3007663665", "tNum": "T170317006554","showapi_fee_code": 0}

r = ShowapiRequest("http://route.showapi.com/28-3", "530851", "e75fe3242d6c440d8b64b3a6534da0ae")
r.addBodyPara("page", "1")
res = r.post()
print(res.text)  # 返回信息
