import requests
import pandas as pd
from datetime import datetime, timedelta

yesterday = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')

# 1. Alamat API (Token langsung digabung di URL)
# Gunakan f-string supaya gampang ganti-ganti token atau tanggal
TOKEN_MAP = {
    'AA0101' : 'MmY4MTYxYWIxNzE4ODE3MzcwODI5NGQ0NjhmZWEzMGM2MDBiMDNmYw==',
    'AA0102' : 'ODA4MjVhNTdjNjJjNzc2MTk5NmJjYjQ2OWU5OTUyZGEzNzk2NzRiMA==',
    'AA0104' : 'NmRjNjgwNmRmZWE1Y2YyNGQ4M2JjNmExOGYzMzUxY2Q1OGQ1OGMxZQ==',
    'AA0104F' : 'ODRhMTJlZmUxOWQ5MWMwYTZjYWVkOWM2MTY3MjQ0ZjJkNGVmNzBlZQ==',
    'AA0105' : 'NTc4YzEzODNhNjZlYmVmNzgzMjVhMTVjM2QwYjNhOGZmYTcwMjM2OA==',
    'AA0106' : 'ZmQ2YWM4YzQ2NDMxZjcyYWMzZTU5MzgyMDY5YjdmMDBhN2Y3MTZiZA==',
    'AA0107' : 'YmQ1ZWQ2YTg0ZDM0OTE0MWU3NGY3ODY2NDY4NTc1ZWViZTRkNTUzMw==',
    'AA0108' : 'ZWM3NWE2YTgxYWE0Yjc0ZWY1NWYwZGQ3MjJjNjNkODEzYzBmN2Y5Mg==',
    'AA0109' : 'MGZmOGJmYzQ0NjliYjQ3ODk0NDAxODljOTBhNDhjOGM4ZDMwNWRmNg=='
}

URL_MAP = {
    'AA0101' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0101&accessToken={TOKEN_MAP['AA0101']}",
    'AA0102' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0102&accessToken={TOKEN_MAP['AA0102']}",
    'AA0104' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0104&accessToken={TOKEN_MAP['AA0104']}",
    'AA0104F' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0104F&accessToken={TOKEN_MAP['AA0104F']}",
    'AA0105' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0105&accessToken={TOKEN_MAP['AA0105']}",
    'AA0106' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0106&accessToken={TOKEN_MAP['AA0106']}",
    'AA0107' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0107&accessToken={TOKEN_MAP['AA0107']}",
    'AA0108' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0108&accessToken={TOKEN_MAP['AA0108']}",
    'AA0109' : f"https://yimmdpackwebapi.ymcapps.net/dpackweb/api/v1/fakturdata?dealerCd=AA0109&accessToken={TOKEN_MAP['AA0109']}"
}

# 2. Ambil Data (Headers dikosongkan karena token sudah ada di URL)
headers = {
    "Content-Type": "application/json",
    "Accept": "application/json"
}

payload = {
    "targetDate": yesterday
}

all_rows = []

for dealer, url in URL_MAP.items():
    try:
        print(f"Ambil data dealer: {dealer}")

        response = requests.post(url, headers=headers, json=payload)

        if response.status_code == 200:
            data_json = response.json()

            if data_json.get('code') == 200 and 'data' in data_json:
                rows = data_json['data']

                if len(rows) > 0:
                    for row in rows:
                        row['dealer'] = dealer  # tambahin kode dealer
                        all_rows.append(row)
                else:
                    print(f"{dealer} tidak ada data")
            else:
                print(f"{dealer} response aneh:", data_json)

        else:
            print(f"{dealer} gagal: {response.status_code}")

    except Exception as e:
        print(f"Error di {dealer}: {e}")

# ======================
# SIMPAN KE EXCEL
# ======================
if len(all_rows) > 0:
    df = pd.DataFrame(all_rows)

    # rapihin nama kolom
    df.columns = df.columns.str.replace('h.', '', regex=False)

    file_path = f"C:\\Automation\\STU\\Data_Dpack_AllDealer_{payload['targetDate']}.xlsx"
    df.to_excel(file_path, index=False)

    print(f"✅ Selesai! File: {file_path}")
else:
    print("Tidak ada data sama sekali")