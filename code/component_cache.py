"""Shared component cache — persists PCB/FPC names across tool runs.

Stores found component names in a JSON file next to the CST project.
When a tool finds a PCB or FPC, it saves the name. Next time another
tool needs the same component, it loads from cache and asks user to confirm.

Usage:
    cache = ComponentCache(project_path)
    
    # Save a found component
    cache.set("pcb", "COMP_PATH:SOLID_NAME")
    
    # Load a cached component (returns None if not cached)
    cached = cache.get("pcb")
    
    # Clear cache
    cache.clear("pcb")
"""

import os
import json


class ComponentCache:
    def __init__(self, project_path: str):
        """Initialize cache file path based on CST project location."""
        project_dir = os.path.dirname(os.path.abspath(project_path))
        model_name = os.path.splitext(os.path.basename(project_path))[0]
        self._path = os.path.join(project_dir, f"{model_name}_component_cache.json")
        self._data = self._load()

    def _load(self) -> dict:
        if os.path.isfile(self._path):
            try:
                with open(self._path, "r") as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def _save(self):
        try:
            with open(self._path, "w") as f:
                json.dump(self._data, f, indent=2)
        except OSError:
            pass

    def get(self, key: str) -> str | None:
        """Get cached component name. Returns None if not cached."""
        return self._data.get(key)

    def set(self, key: str, value: str):
        """Save component name to cache."""
        self._data[key] = value
        self._save()

    def clear(self, key: str = None):
        """Clear one key or entire cache."""
        if key:
            self._data.pop(key, None)
        else:
            self._data = {}
        self._save()
