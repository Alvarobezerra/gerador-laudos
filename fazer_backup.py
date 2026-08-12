import os
import shutil
import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_FILE = os.path.join(BASE_DIR, 'app.py')
BACKUP_DIR = os.path.join(BASE_DIR, 'backups')
os.makedirs(BACKUP_DIR, exist_ok=True)

agora = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
backup_filename = f'app_backup_{agora}.py'
backup_filepath = os.path.join(BACKUP_DIR, backup_filename)

try:
    shutil.copy2(APP_FILE, backup_filepath)
    print(f'[SUCESSO] Backup salvo em: {backup_filepath}')
except Exception as e:
    print(f'[ERRO] Falha ao criar backup: {e}')
