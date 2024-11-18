import requests


vendedor_info = requests.get(f"https://api.mercadolibre.com/users/55304068").json()
print(vendedor_info.get("address").get("city"))