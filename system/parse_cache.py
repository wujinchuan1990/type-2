"""Bounded immutable record cache keyed by content and parser implementation."""

from collections import OrderedDict
from copy import deepcopy
import hashlib
import threading


class ParseCache:
    def __init__(self, limit=24):
        self.limit = limit
        self.entries = OrderedDict()
        self.lock = threading.Lock()

    def read(self, path, parser_key, parse, *, refresh=False):
        key = (str(path.resolve()), hashlib.sha256(path.read_bytes()).hexdigest(), parser_key)
        with self.lock:
            if refresh:
                self.entries.pop(key, None)
            if key in self.entries:
                self.entries.move_to_end(key)
                return deepcopy(self.entries[key])
        result = parse()
        if hashlib.sha256(path.read_bytes()).hexdigest() != key[1]:
            raise ValueError("原始记录在解析期间发生变化，请重新上传并核对。")
        with self.lock:
            self.entries[key] = deepcopy(result)
            self.entries.move_to_end(key)
            while len(self.entries) > self.limit:
                self.entries.popitem(last=False)
        return result
