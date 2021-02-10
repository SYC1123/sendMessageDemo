# python3.6.5
# 需要引入requests包 ：运行终端->进入python/Scripts ->输入：pip install requests
from ShowapiRequest import ShowapiRequest

r = ShowapiRequest("http://route.showapi.com/28-2", "530851", "e75fe3242d6c440d8b64b3a6534da0ae")
r.addBodyPara("content", "您好同学,您的今日校园签到任务未完成，请尽快完成，谢谢！")
r.addBodyPara("title", "计软集团")
r.addBodyPara("notiPhone", "18545626763")
r.addBodyPara("userIp", "")
res = r.post()
print(res.text)  # 返回信息
