#обязательно надо в системе поставить pandas - 
#apt get update 
#apt install python3-pip
#pip install pandas openpyxl
#pip install pandas PyPDF2 openpyxl
#pip install xlrd

import logging
import os
import pandas as pd
import PyPDF2
import warnings
warnings.filterwarnings('ignore')

# Настройка логирования
logging.basicConfig(
    filename='log.txt', 
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# Чтение файла с логинами
logins = []
with open('employees.txt', 'r', encoding='utf-8') as f:
    for line in f:
        parts = line.strip().split()
        if len(parts) >= 2:
            logins.append(parts[-1])  # Предполагаем, что логин последний в строке
        else:
            # Если в строке только логин без ФИО
            logins.append(line.strip())

# Ключевые слова для поиска
keywords = ["транспортные расходы", "кислород", "скидка", "скидки", "доставка"]
target_extensions = ['.pdf', '.xls', '.xlsx']
files_to_delete = []

# Список целевых папок для проверки внутри каждой директории пользователя
target_folders = ['Рабочий стол', 'Документы', 'Загрузки']

# Улучшенная проверка содержимого Excel файлов с явным указанием движка
def check_excel_content(file_path, keywords):
    try:
        # Явно определяем движок в зависимости от расширения файла
        if file_path.lower().endswith('.xlsx'):
            engine = 'openpyxl'
        elif file_path.lower().endswith('.xls'):
            engine = 'xlrd'
        else:
            # Для других расширений пытаемся использовать оба движка
            try:
                engine = 'openpyxl'
                xl = pd.ExcelFile(file_path, engine=engine)
            except:
                engine = 'xlrd'
                xl = pd.ExcelFile(file_path, engine=engine)
        
        # Читаем файл с выбранным движком
        xl = pd.ExcelFile(file_path, engine=engine)
        
        for sheet_name in xl.sheet_names:
            try:
                # Пытаемся прочитать лист
                df = xl.parse(sheet_name, header=None)
                
                # Проверяем наличие ключевых слов в ячейках
                for row in df.values:
                    for cell in row:
                        if pd.isna(cell):
                            continue
                        cell_str = str(cell).lower()
                        if any(keyword in cell_str for keyword in keywords):
                            return True
            except Exception as e:
                # Логируем ошибку чтения листа, но продолжаем проверку других листов
                logging.warning(f"Ошибка чтения листа {sheet_name} в файле {file_path}: {str(e)}")
                continue
                
    except Exception as e:
        # Логируем ошибку, но не прерываем выполнение
        logging.error(f"Ошибка чтения Excel {file_path}: {str(e)}")
    
    return False

# Указываем корневую директорию
root_dir = r'/home/test'  # Указываем путь к корневой директории с папками пользователей

# Проверяем только папки пользователей из файла employees.txt
for dir_name in os.listdir(root_dir):
    dir_path = os.path.join(root_dir, dir_name)
    
    # Проверяем, является ли это директорией и есть ли она в списке логинов
    if os.path.isdir(dir_path) and dir_name in logins:
        logging.info(f"Обрабатываем директорию пользователя: {dir_name}")
        
        # Проверяем только целевые папки внутри директории пользователя
        for folder_name in target_folders:
            folder_path = os.path.join(dir_path, folder_name)
            
            # Проверяем существование папки
            if not os.path.exists(folder_path):
                logging.info(f"Папка {folder_path} не существует, пропускаем")
                continue
                
            # Рекурсивно обходим все файлы в целевой папке
            for root, _, files in os.walk(folder_path):
                for file in files:
                    file_lower = file.lower()
                    ext = os.path.splitext(file_lower)[1]
                    
                    if ext in target_extensions:
                        full_path = os.path.join(root, file)
                        
                        # Проверка имени файла
                        if any(keyword in file_lower for keyword in keywords):
                            files_to_delete.append(full_path)
                            continue
                        
                        # Проверка содержимого для Excel
                        if ext in ['.xls', '.xlsx']:
                            if check_excel_content(full_path, keywords):
                                files_to_delete.append(full_path)

# Удаление файлов
for file_path in files_to_delete:
    try:
        os.remove(file_path)
        print(f"Удален: {file_path}")
        logging.info(f"Удален: {file_path}")
    except Exception as e:
        print(f"Ошибка удаления {file_path}: {str(e)}")
        logging.error(f"Ошибка удаления {file_path}: {str(e)}")