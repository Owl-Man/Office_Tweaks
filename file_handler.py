import os


def delete_files_starts_with(key: str) -> None:
    for file in os.listdir(os.getcwd()):
        if file.startswith(key):
            os.remove(rf'{os.getcwd()}\{file}')


def delete_files_ends_with(key: str) -> None:
    for file in os.listdir(os.getcwd()):
        if file.split('.')[0].endswith(key):
            os.remove(rf'{os.getcwd()}\{file}')


def delete_files_with(key: str) -> None:
    for file in os.listdir(os.getcwd()):
        if file.count(key) > 0:
            os.remove(rf'{os.getcwd()}\{file}')


def delete_files_with_ext(key: str) -> None:
    for file in os.listdir(os.getcwd()):
        if file.count('.') == 1 and file.split('.')[1] == key:
            os.remove(rf'{os.getcwd()}\{file}')


def find_file_with_ext(ext: str) -> list:
    output = []

    for file in os.listdir(os.getcwd()):
        if file.count('.') == 1 and file.split('.')[1] == ext:
            output.append(file)

    return output


def find_file_with_exts(exts: list) -> list:
    output = []

    for ext in exts:
        output += find_file_with_ext(ext)

    return output
