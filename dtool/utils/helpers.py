import time
import shutil
from functools import wraps
import logging


def timer(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_time = time.perf_counter()

        elapsed_time_seconds = end_time - start_time
        minutes = int(elapsed_time_seconds // 60)
        remaining_seconds = elapsed_time_seconds % 60
        seconds = int(remaining_seconds)
        milliseconds = int((remaining_seconds - seconds) * 1000)

        logging.info(f"Function execution time: {minutes}m {seconds}s {milliseconds}ms")
        return result

    return wrapper


def generate_filename(basename: str, index: int) -> str:
    return f"{index+1:04d}_{basename}.docx"

