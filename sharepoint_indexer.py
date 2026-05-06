import sqlite3
import requests
import re

DB_PATH = "sharepoint_index.db"

# 🔥 CACHE PARA NO PEDIR SITE_ID SIEMPRE
SITE_ID_CACHE = {}

# =========================================================
# INIT DB
# =========================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS files (
        id TEXT PRIMARY KEY,
        name TEXT,
        url TEXT
    )
    """)

    conn.commit()
    conn.close()


# =========================================================
# GET SITE ID
# =========================================================
def get_site_id(session, headers, site_url):

    if site_url in SITE_ID_CACHE:
        return SITE_ID_CACHE[site_url]

    url = f"https://graph.microsoft.com/v1.0/sites/{site_url}"

    try:
        r = session.get(url, headers=headers, timeout=30)

        if r.status_code != 200:
            print(f"❌ Error obteniendo site_id para {site_url}")
            return None

        site_id = r.json().get("id")
        SITE_ID_CACHE[site_url] = site_id

        print(f"✅ Site ID obtenido: {site_id}")

        return site_id

    except Exception as e:
        print(f"⚠️ Error en get_site_id: {e}")
        return None


# =========================================================
# SEARCH LOCAL
# =========================================================
def search_local(query: str):

    def normalize(text):
        text = text.lower()
        text = re.sub(r'[^a-z0-9]', '', text)
        return text

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    q_norm = normalize(query)

    cur.execute("SELECT name, url FROM files")

    results = []

    for name, url in cur.fetchall():

        name_norm = normalize(name)

        if q_norm in name_norm:
            results.append({
                "name": name,
                "url": url
            })

        if len(results) >= 5:
            break

    conn.close()
    return results


# =========================================================
# 🔥 NUEVO: SEARCH SHAREPOINT (SIN INDEXADO)
# =========================================================
def search_sharepoint(query, get_token, get_session):

    token = get_token()
    session = get_session()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = "https://graph.microsoft.com/v1.0/search/query"

    body = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {
                    "queryString": query
                },
                "from": 0,
                "size": 10
            }
        ]
    }

    try:
        r = session.post(url, headers=headers, json=body, timeout=30)

        if r.status_code != 200:
            print("❌ Error en búsqueda SharePoint")
            print(r.text)
            return []

        data = r.json()

        results = []

        hits = data.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

        for hit in hits:
            resource = hit.get("resource", {})

            results.append({
                "name": resource.get("name"),
                "url": resource.get("webUrl")
            })

        print(f"🔎 Resultados encontrados: {len(results)}")

        return results

    except Exception as e:
        print(f"⚠️ Error en search_sharepoint: {e}")
        return []


# =========================================================
# RECURSIVO (INDEXADOR)
# =========================================================
def _index_drive_recursive(session, headers, drive_id, folder_id, cur):

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"

    try:
        r = session.get(url, headers=headers, timeout=30)

        if r.status_code != 200:
            print(f"❌ Error en drive {drive_id} carpeta {folder_id}")
            return

        items = r.json().get("value", [])

        for it in items:

            # 📄 archivo
            if "file" in it:
                print("📄 INDEXANDO:", it.get("name"))

                cur.execute(
                    "INSERT OR REPLACE INTO files (id, name, url) VALUES (?, ?, ?)",
                    (it["id"], it["name"], it["webUrl"])
                )

            # 📁 carpeta → recursion
            if "folder" in it:
                _index_drive_recursive(session, headers, drive_id, it["id"], cur)

    except Exception as e:
        print(f"⚠️ Error en carpeta {folder_id}: {e}")


# =========================================================
# INDEXADOR PRINCIPAL
# =========================================================
def index_sharepoint(get_token, get_session):

    print("🚀 INICIANDO INDEXACIÓN SHAREPOINT")

    token = get_token()
    session = get_session()

    headers = {"Authorization": f"Bearer {token}"}

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("DELETE FROM files")

    SITE_URLS = [
        "acfae.sharepoint.com:/sites/PROYECTOSOEMCTELECTROMECNICA",
        "acfae.sharepoint.com:/sites/GestorDocumentalIT",
        "acfae.sharepoint.com:/sites/GestorDocumentalFinanzas",
        "acfae.sharepoint.com:/sites/GestorDocumentalIndustrial",
        "acfae.sharepoint.com:/sites/PROYECTOSOEMCTSENSORICA"
    ]

    for site_url in SITE_URLS:

        print(f"\n🔍 Site URL: {site_url}")

        site_id = get_site_id(session, headers, site_url)

        if not site_id:
            print("❌ Saltando site...")
            continue

        drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

        try:
            r = session.get(drives_url, headers=headers, timeout=30)

            if r.status_code != 200:
                print("❌ Error obteniendo drives")
                continue

            drives = r.json().get("value", [])

            for drive in drives:

                drive_id = drive["id"]
                print(f"   📁 Drive: {drive.get('name')}")

                _index_drive_recursive(session, headers, drive_id, "root", cur)

        except Exception as e:
            print(f"⚠️ Error en drives: {e}")

    conn.commit()
    conn.close()

    print("✅ Indexación completada")