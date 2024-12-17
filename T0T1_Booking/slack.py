import requests

# إعداد التوكن والرؤوس
token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJleHAiOjE3Mjc3ODc2MjY2NTgsImFkbWluX2lkIjoiMDI4ZjJkYmEtMzFiMy00ODNiLWE4YTAtZjEzZjczZDc5OTljIiwibmFtZSI6IkFobWVkIEhhZ2dhZyIsImVtYWlsIjoiYWhtZWQuaGFnZ2FnQHRobmRyLmFwcCJ9.La-zqUVq3wJ63OZpl-FOhMOF0ot6ojkwJfgqRDeHVcStffbVvmz_iqUUunS1UJpMjFR66Kj4nDGCs9t6XmcICq6fY-NheD4bRV3KlUAmKWQA1Hx_YHAyc_hpOuFh9WIY8GEuOm2XxL-Ll13Dl26ug7RUCfu3cigl7uAjWDxG_34"
cookies = {
    "security": "your_security_cookie_value"
}

# عنوان URL
url = "https://usertool.thndr-internal.app/users/0OmBrohbDJSnLmkHf1KjDM8yPL83"

# إعداد الرؤوس
headers = {
    "Authorization": f"Bearer {token}",
    "User-Agent": "Mozilla/5.0"
}

# إرسال الطلب
response = requests.get(url, headers=headers, cookies=cookies)

# التحقق من الاستجابة
if response.status_code == 200:
    data = response.json()
    print(data)
else:
    print(f"Failed to fetch data: {response.status_code} - {response.text}")