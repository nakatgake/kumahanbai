import urllib.request
import urllib.parse
from http.cookiejar import CookieJar

cookie_jar = CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))

print("Logging in...")
data = urllib.parse.urlencode({
    "username": "nakamura@connect-web.jp",
    "password": "N687nh4su4"
}).encode("utf-8")
login_req = urllib.request.Request("http://127.0.0.1:8000/login", data=data)
opener.open(login_req)

print("Fetching /invoices...")
res = opener.open("http://127.0.0.1:8000/invoices")
print("Invoices status:", res.getcode())

print("Testing bulk print POST...")
post_data = urllib.parse.urlencode([("invoice_ids[]", "2"), ("invoice_ids[]", "4")]).encode("utf-8")
bp_req = urllib.request.Request("http://127.0.0.1:8000/invoices/bulk_print", data=post_data)

try:
    bp_res = opener.open(bp_req)
    print("Bulk print status:", bp_res.getcode())
except urllib.error.HTTPError as e:
    print("Bulk print error:", e.code)
    print("ERROR BODY START")
    print(e.read().decode('utf-8'))
    print("ERROR BODY END")
except Exception as e:
    print("Other error:", str(e))
