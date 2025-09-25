import requests

# ⚠️ Ajusta estas variables con tu información real
BASE_URL = "https://cientificavirtual.cientifica.edu.pe/api/v1"
TOKEN = "21289~yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w"   # pega tu token aquí

# Endpoint de prueba: listar cursos
url = f"{BASE_URL}/courses"

headers = {
    "Authorization": f"Bearer {TOKEN}"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    cursos = response.json()
    print("✅ Conexión exitosa. Cursos disponibles:")
    for c in cursos:
        print(f"- {c.get('id')} : {c.get('name')}")
else:
    print("❌ Error:", response.status_code, response.text)
