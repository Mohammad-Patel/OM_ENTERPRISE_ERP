import os
import shutil
from datetime import datetime

def create_excel_backup():
    source_file = 'path_to_your_excel_file.xlsx' # Apni main Excel file ka path yahan dein
    backup_dir = 'E:\OM_ENTERPRISE_ERP\backups'

    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    if os.path.exists(source_file):
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        backup_file = os.path.join(backup_dir, f'erp_backup_{timestamp}.xlsx')
        shutil.copy(source_file, backup_file)
        print(f"Backup created successfully: {backup_file}")