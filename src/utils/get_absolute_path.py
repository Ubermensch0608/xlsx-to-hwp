from pathlib import Path

def get_absolute_path(relative_path):
    relative = Path(relative_path)
    absolute = relative.resolve()
    return absolute
