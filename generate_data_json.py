import json
import os
import server

# Generate data.json from server's extraction logic
data = server.get_correspondence_data()
json_path = r"c:\Users\prasa\Desktop\TS-Correspondence Matrix\data.json"

with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("data.json successfully generated with", len(data.get('records', [])), "records.")
