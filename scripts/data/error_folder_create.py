import os 
from datetime import datetime
import json

today = datetime.now().strftime("%Y-%m-%d")

folder = '/data/errors/'+today

if not os.path.isdir(folder):
    os.makedirs(folder)

result = {"ok": True, "folder": folder}
print(json.dumps(result))