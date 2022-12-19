import os
from PIL import Image
import file_handler


def resize_process(img_name: str, q: int) -> None:
    img_file = Image.open(rf'{os.getcwd()}\{img_name}')
    img_file.save(rf'{os.getcwd()}\{img_name}', quality=q)
    print(f'Файл "{img_name}" успешно преобразован!')


def image_resize() -> None:
    img_files = file_handler.find_file_with_exts(['jpeg', 'gif', 'png', 'jpg'])

    if len(img_files) == 0:
        print('В данном каталоге нет файлов с расширением (".jpeg", ".gif", ".png", ".jpg")\n')
        return

    print('Список файлов с раширением (".jpeg", ".gif", ".png", ".jpg") в данном каталоге: \n')

    for i in range(0, len(img_files)):
        print(f'{i + 1}. {img_files[i]}')

    chosen_file_id = int(
        input(
            'Введите номер файла для преобразования (чтобы преобразовать все файлы из данного каталога введите 0): '))
    quality = int(input('Введите параметры сжатия (от 0% до 100%): '))

    if chosen_file_id not in range(0, len(img_files) + 1) or quality not in range(0, 101):
        print('Неверный номер')
        return

    if chosen_file_id == 0:
        for j in range(0, len(img_files)):
            resize_process(img_files[j], quality)
    else:
        resize_process(img_files[chosen_file_id-1], quality)
