import requests
from requests_ntlm import HttpNtlmAuth

BASE = "http://faevmwcapp01.fae.dom/Windchill"
USER = r"FAE\\TU_USUARIO"
PASS = "TU_PASSWORD"

session = requests.Session()
session.auth = HttpNtlmAuth(USER, PASS)

# 1) Login real a Windchill (crea sesión)
login_url = f"{BASE}/j_security_check"
session.post(login_url, data={
    "j_username": USER,
    "j_password": PASS
})

# 2) Ahora sí, OData funciona porque ya hay cookie
doc_id = "OR:wt.doc.WTDocument:202687"
odata_url = f"{BASE}/servlet/odata/v2/DocMgmt/Documents('{doc_id}')/PrimaryContent/$value"

r = session.get(odata_url)

print("Status:", r.status_code)

with open("documento.xlsx", "wb") as f:
    f.write(r.content)

print("Descargado como documento.xlsx")
