import json
from pathlib import Path
from typing import Any, Dict, List


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default


def save_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_queries(path: Path) -> List[Dict[str, Any]]:
    return load_json(path, default=[])


def save_queries(path: Path, queries: List[Dict[str, Any]]) -> None:
    save_json(path, queries)


def load_seen_ids(path: Path) -> Dict[str, str]:
    return load_json(path, default={})


def save_seen_ids(path: Path, seen: Dict[str, str]) -> None:
    save_json(path, seen)
