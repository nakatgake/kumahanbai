import urllib.request
import urllib.parse
from http.cookiejar import CookieJar
import sqlite3

# Reset invoices to UNPAID so they trigger the rendering logic
conn = sqlite3.connect("kumanogo.db")
c = conn.cursor()
c.execute("UPDATE invoices SET status='UNPAID' WHERE id IN (SELECT id FROM invoices LIMIT 2)")
conn.commit()

cookie_jar = CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))

print("Logging in...")
data = urllib.parse.urlencode({
    "username": "nakamura@connect-web.jp",
    "password": "N687nh4su4"
}).encode("utf-8")
login_req = urllib.request.Request("http://127.0.0.1:8000/login", data=data)
opener.open(login_req)

c.execute("SELECT id FROM invoices WHERE status='UNPAID' LIMIT 2")
ids = [str(r[0]) for r in c.fetchall()]
conn.close()

print(f"Testing bulk print POST with valid ids: {ids}")
if ids:
    post_data = urllib.parse.urlencode([("invoice_ids[]", i) for i in ids]).encode("utf-8")
    bp_req = urllib.request.Request("http://127.0.0.1:8000/invoices/bulk_print", data=post_data)
    try:
        bp_res = opener.open(bp_req)
        print("Bulk print status:", bp_res.getcode())
    except urllib.error.HTTPError as e:
        print("Bulk print HTTP error:", e.code)
        print("ERROR BODY START")
        print(e.read().decode('utf-8', errors='replace'))
        print("ERROR BODY END")
    except Exception as e:
        print("Other error:", str(e))
else:
    print("No invoices to test.")
