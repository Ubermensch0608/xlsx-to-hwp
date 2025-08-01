import os
from utils.get_absolute_path import get_absolute_path

def delete_files_by_extension(dir, extension):
    files = [f for f in os.listdir(get_absolute_path(dir)) if f.endswith(extension)]
    for file in files:
        os.remove(get_absolute_path(f'{dir}/{file}'))
