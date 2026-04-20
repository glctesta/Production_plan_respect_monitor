"""
PlanRespect — Crea pacchetto deploy ZIP
========================================
Uso: python create_deploy_zip.py
"""
import datetime
import os
import zipfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Cartelle da includere (relative a BASE_DIR)
INCLUDE_DIRS = [
    'templates',
    'static',
]

# File singoli da includere
INCLUDE_FILES = [
    'app.py',
    'api_routes.py',
    'app_config.py',
    'config_manager.py',
    'db_connection.py',
    'db_queries.py',
    'email_alerter.py',
    'email_connector.py',
    'excel_parser.py',
    'monitor_engine.py',
    'scheduler.py',
    'utils.py',
    'config.yaml',
    'output_config.json',
    'requirements.txt',
    'start_monitor.bat',
    'Logo.png',
    'claude_code_production_monitor_spec.md',
    'Documentazione_Tecnica_PlanRespect.docx',
    'GUIDA IN ITALIANO.docx',
    'GUIDE IN ENGLISH.docx',
    'GUIDE SVENSKA.docx',
    'Ghid_Utilizator_RO.docx',
    'Gid in Romana.docx',
    'Monitorizarea Respectului Planului de Producție.docx',
    'PlanRespect_UserGuide-Romanian.pdf',
]

EXCLUDE_PATTERNS = [
    '__pycache__',
    '.pyc',
    '.pyo',
    '.git',
    '.env',
    '.claude',
    'node_modules',
    '.vscode',
    '.idea',
    'venv',
    '.venv',
    'logs',
    # Segreti per-installazione: mai nel pacchetto
    'db_config.enc',
    'db_credentials.enc',
    'email_key.key',
    'email_credentials.enc',
    'encryption_key.key',
    '.DS_Store',
    'Thumbs.db',
    '*.zip',
]


def should_exclude(path: str) -> bool:
    for pattern in EXCLUDE_PATTERNS:
        if pattern in path:
            return True
    return False


def create_zip() -> str:
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    zip_name = f'PlanRespect_deploy_{timestamp}.zip'
    zip_path = os.path.join(BASE_DIR, zip_name)
    file_count = 0

    print()
    print('=' * 50)
    print('  PlanRespect — Deploy Package')
    print('=' * 50)
    print()

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for dir_name in INCLUDE_DIRS:
            dir_path = os.path.join(BASE_DIR, dir_name)
            if not os.path.isdir(dir_path):
                print(f'  [SKIP] Cartella non trovata: {dir_name}/')
                continue
            for root, dirs, files in os.walk(dir_path):
                dirs[:] = [d for d in dirs if not should_exclude(d)]
                for fname in files:
                    fpath = os.path.join(root, fname)
                    rel = os.path.relpath(fpath, BASE_DIR)
                    if should_exclude(rel):
                        continue
                    arcname = os.path.join('PlanRespect', rel)
                    zf.write(fpath, arcname)
                    file_count += 1

        for fname in INCLUDE_FILES:
            fpath = os.path.join(BASE_DIR, fname)
            if not os.path.exists(fpath):
                print(f'  [SKIP] File non trovato: {fname}')
                continue
            arcname = os.path.join('PlanRespect', fname)
            zf.write(fpath, arcname)
            file_count += 1

    size_mb = os.path.getsize(zip_path) / (1024 * 1024)
    print()
    print(f'  [OK] Pacchetto creato: {zip_name}')
    print(f'     File inclusi:  {file_count}')
    print(f'     Dimensione:    {size_mb:.2f} MB')
    print(f'     Percorso:      {zip_path}')
    print()
    print('=' * 50)
    print()
    return zip_path


if __name__ == '__main__':
    create_zip()
