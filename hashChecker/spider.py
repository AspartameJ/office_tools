import requests
from bs4 import BeautifulSoup
import json

base_url = "https://hf-mirror.com/Qwen/Qwen2-72B/blob/main/model-"
total_models = 37
hashes = {}

for i in range(1, total_models + 1):
    url = f"{base_url}{str(i).zfill(5)}-of-00037.safetensors"
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        hash_element = soup.find('strong', string='SHA256:')
        if hash_element:
            hash_value = hash_element.next_sibling.strip() if hash_element.next_sibling else "未找到哈希值"
            hashes[f"model-{str(i).zfill(5)}-of-00037.safetensors"] = hash_value
        else:
            hashes[f"model-{str(i).zfill(5)}-of-00037.safetensors"] = ""
    else:
        hashes[f"model-{str(i).zfill(5)}-of-00037.safetensors"] = ""

# 将哈希值写入hashes.json文件
with open('hashes.json', 'w') as f:
    json.dump(hashes, f, indent=4)