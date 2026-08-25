import json
try:
    from memory_snapshot import snapshot_memory
    print("MEMORY_SNAPSHOT_BEGIN")
    print(json.dumps(snapshot_memory(), ensure_ascii=False))
    print("MEMORY_SNAPSHOT_END")
except Exception as exc:
    print("MEMORY_SNAPSHOT_ERROR", repr(exc))
