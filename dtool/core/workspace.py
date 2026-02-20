import zipfile
import io
from typing import Dict


class InMemoryDocx:
    def __init__(self, file_path: str):
        self.assets: Dict[str, bytes] = {}
        self.load(file_path)

    def load(self, file_path: str):
        with open(file_path, "rb") as f:
            file_bytes = f.read()

        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            for filename in z.namelist():
                self.assets[filename] = z.read(filename)

    def get_xml(self, filename: str) -> bytes:
        return self.assets.get(filename, b"")
