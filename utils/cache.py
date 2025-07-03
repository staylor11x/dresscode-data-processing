import json
import os

CACHE_FILE = "processed_log.json"

def load_cache():
    if not os.path.exists(CACHE_FILE):
        return {}
    with open(CACHE_FILE, "r") as f:
        return json.load(f)

def save_cache(cache):
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)

def was_processed(cache, filename):
    return cache.get(filename, False)

def mark_processed(cache, filename):
    cache[filename] = True
