import os
import shutil
import time
from contextlib import contextmanager
from typing import Generator

from ..utils.constants import TEMP_COPY_DIR, TEMP_DIR


class Workspace:
    def __init__(self, base_path: str = None):
        self.base_path = base_path or os.path.dirname(os.path.realpath(__file__))
        self.container_path = ""
        self.temp_copy_path = ""
        self.temp_dir_path = ""

    def create(self) -> None:
        timestamp = time.time_ns()
        self.container_path = os.path.join(self.base_path, f"container_{timestamp}")
        self.temp_copy_path = os.path.join(
            self.container_path, f"{TEMP_COPY_DIR}_{timestamp}"
        )
        self.temp_dir_path = os.path.join(self.container_path, TEMP_DIR)

        os.makedirs(self.container_path, exist_ok=False)
        os.makedirs(self.temp_copy_path, exist_ok=True)
        os.makedirs(self.temp_dir_path, exist_ok=True)

    def cleanup(self) -> None:
        if os.path.exists(self.container_path):
            shutil.rmtree(self.container_path, ignore_errors=True)


@contextmanager
def temporary_workspace(
    base_path: str = None,
    cleanup: bool = True,
) -> Generator[Workspace, None, None]:
    workspace = Workspace(base_path)
    workspace.create()
    try:
        yield workspace
    finally:
        if cleanup:
            workspace.cleanup()
