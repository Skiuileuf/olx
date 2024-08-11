import requests
import os

os.makedirs("api", exist_ok=True)

offset = 0
limit = 50
max_results = 1000
query = "pisica" # slightly encoded search query ???

for i in range(0, max_results, limit):
    offset = i
    response = requests.get(f'https://www.olx.ro/api/v1/offers?offset={offset}&limit={limit}&query={query}&')

    f = open(f"api/olx-{offset}-{limit}-{query}.json", "wb")
    f.write(response.content)
    f.close()
    print(f"Downloaded {offset}-{limit}-{query}.json")