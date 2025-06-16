import os
import zipfile
from PIL import Image
import shutil

def extract_images_from_docx(docx_path, output_folder):
    """
    Извлекает все изображения из файла .docx и сохраняет их в указанную папку.
    
    :param docx_path: Путь к файлу .docx
    :param output_folder: Папка для сохранения изображений
    """
    # Нормализуем пути
    docx_path = os.path.abspath(docx_path)
    output_folder = os.path.abspath(output_folder)
    
    print(f"Пытаемся открыть файл: {docx_path}")
    
    # Создаем папку для изображений, если ее нет
    os.makedirs(output_folder, exist_ok=True)
    
    # Временная папка для распаковки .docx
    temp_folder = os.path.abspath("temp_docx_extract")
    os.makedirs(temp_folder, exist_ok=True)
    
    try:
        # Проверяем существование файла и права доступа
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"Файл не найден: {docx_path}")
            
        if not os.access(docx_path, os.R_OK):
            raise PermissionError(f"Нет прав на чтение файла: {docx_path}")
        
        # Распаковываем .docx как zip-архив
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_folder)
        
        # Путь к папке с медиафайлами в .docx
        media_path = os.path.join(temp_folder, 'word', 'media')
        
        if os.path.exists(media_path):
            # Копируем все файлы из папки media в целевую папку
            for filename in os.listdir(media_path):
                file_path = os.path.join(media_path, filename)
                
                # Проверяем, является ли файл изображением
                try:
                    with Image.open(file_path) as img:
                        # Формируем путь для сохранения
                        name, ext = os.path.splitext(filename)
                        output_path = os.path.join(output_folder, f"image_{name}{ext}")
                        
                        # Если файл с таким именем уже существует, добавляем суффикс
                        counter = 1
                        while os.path.exists(output_path):
                            output_path = os.path.join(output_folder, f"image_{name}_{counter}{ext}")
                            counter += 1
                        
                        # Сохраняем изображение
                        img.save(output_path)
                        print(f"Изображение сохранено: {output_path}")
                except:
                    # Пропускаем файлы, которые не являются изображениями
                    continue
            print(f"Успешно извлечено {len(os.listdir(media_path))} изображений.")
        else:
            print("В документе не найдены изображения.")
            
    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
    finally:
        # Удаляем временную папку
        if os.path.exists(temp_folder):
            try:
                shutil.rmtree(temp_folder)
            except:
                pass

if __name__ == "__main__":
    # Пример использования
    print("Извлечение изображений из .docx")
    print("--------------------------------")
    
    docx_file = input("Введите путь к файлу .docx: ").strip()
    output_dir = input("Введите папку для сохранения изображений (по умолчанию 'images'): ").strip() or "images"
    
    # Обработка пути, если он заключен в кавычки (как при drag&drop)
    if docx_file.startswith('"') and docx_file.endswith('"'):
        docx_file = docx_file[1:-1]
    
    extract_images_from_docx(docx_file, output_dir)
    input("Нажмите Enter для выхода...")
