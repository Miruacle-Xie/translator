import json
import random
import hashlib
from urllib import parse
import http.client
import excelOperate


class BaiduTranslate:
    def __init__(self, fromLang, toLang, appid, secretKey):
        self.url = "/api/trans/vip/translate"
        self.appid = appid  # 申请的账号
        self.secretKey = secretKey  # 账号密码
        self.fromLang = fromLang
        self.toLang = toLang
        self.salt = random.randint(32768, 65536)

    def BdTrans(self, text):
        sign = self.appid + text + str(self.salt) + self.secretKey
        md = hashlib.md5()
        md.update(sign.encode(encoding='utf-8'))
        sign = md.hexdigest()
        myurl = self.url + \
                '?appid=' + self.appid + \
                '&q=' + parse.quote(text) + \
                '&from=' + self.fromLang + \
                '&to=' + self.toLang + \
                '&salt=' + str(self.salt) + \
                '&sign=' + sign
        try:
            httpClient = http.client.HTTPConnection('api.fanyi.baidu.com')
            httpClient.request('GET', myurl)
            response = httpClient.getresponse()
            html = response.read().decode('utf-8')
            html = json.loads(html)
            dst = html["trans_result"][0]["dst"]
            return True, dst
        except Exception as e:
            return False, e
