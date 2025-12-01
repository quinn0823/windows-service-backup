# Windows Service Backup
# The MIT License (MIT)
# Copyright (c) 2025 Jonathan Chiu

import sys
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime
import zipfile

# ---------------------------------------------
# Globals
# ---------------------------------------------
CONFIG_FILE = 'config.json'
LOG_FILE_HANDLE = None
CURRENT_LOG_PATH = None
LOG_QUEUE = []
USAGE = '''Usage:
    python main.py [command]

Commands:
    backup           Perform backup of services and docker
    help             Show this help message
    prune            Prune old backup versions
    start            Start services and docker containers
    startDocker      Start docker containers only
    startServices    Start services only
    restart          Restart services and docker containers
    restartDocker    Restart docker containers only
    restartServices  Restart services only
    stop             Stop services and docker containers
    stopDocker       Stop docker containers only
    stopServices     Stop services only'''
COMMANDS = ['backup', 'help', 'prune', 'start', 'startDocker', 'startServices', 'stop', 'stopDocker', 'stopServices', 'restart', 'restartDocker', 'restartServices']

# ---------------------------------------------
# Utility: timestamp
# ---------------------------------------------
def now():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def timestamp_folder():
    return datetime.now().strftime('%Y-%m-%d %H%M%S')

# ---------------------------------------------
# Logging
# ---------------------------------------------
def log(msg):
    print(msg)

    global LOG_FILE_HANDLE, LOG_QUEUE

    if LOG_FILE_HANDLE:
        if len(LOG_QUEUE) > 0:
            for line in LOG_QUEUE:
                LOG_FILE_HANDLE.write(line + '\n')
            LOG_QUEUE.clear()
        LOG_FILE_HANDLE.write(msg + '\n')
        LOG_FILE_HANDLE.flush()
    else:
        LOG_QUEUE.append(msg)

def f_log(level, comp, msg):
    line = f'{now()} [{level}] {comp:<7} {msg}'
    log(line)

# ---------------------------------------------
# Load config
# ---------------------------------------------
def load_config():
    f_log('INFO', 'CONFIG', f'Loading config...')
    cfg_path = Path(CONFIG_FILE)
    if not cfg_path.exists():
        f_log('ERRO', 'CONFIG', f'config not found: {cfg_path.resolve()}')
        sys.exit(1)

    try:
        cfg = json.loads(cfg_path.read_text(encoding='utf-8'))
        f_log('INFO', 'CONFIG', f'Loaded {cfg}')
        return cfg
    except Exception as e:
        f_log('ERRO', 'CONFIG', f'Failed to parse config: {e}')
        sys.exit(1)

# ---------------------------------------------
# Run shell command
# ---------------------------------------------
def run_cmd(cmd, timeout=None):
    try:
        p = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=timeout)
        return p.returncode == 0, (p.stdout or '').strip(), (p.stderr or '').strip(), p.returncode
    except Exception as e:
        return False, '', str(e), -1

# ---------------------------------------------
# Service operations
# ---------------------------------------------
def stop_services(cfg):
    for svc in cfg['services']:
        name = svc['name']
        f_log('INFO', 'SERVICE', f'Stopping "{name}"...')

        ok, out, err, code = run_cmd(f'net stop "{name}"')
        if ok:
            f_log('DONE', 'SERVICE', f'Stopped "{name}"')
            continue

        combined = out + err
        if '3521' in combined:
            f_log('DONE', 'SERVICE', f'"{name}" was not running')
            continue

        reason = err or out or f'Return code {code}'
        f_log('ERRO', 'SERVICE', f'Failed stopping "{name}": {reason}')
        sys.exit(1)

def start_services(cfg):
    for svc in cfg['services']:
        name = svc['name']
        f_log('INFO', 'SERVICE', f'Starting "{name}"...')

        ok, out, err, code = run_cmd(f'net start "{name}"')
        if ok:
            f_log('DONE', 'SERVICE', f'Started "{name}"')
            continue

        combined = out + err
        if '2182' in combined:
            f_log('DONE', 'SERVICE', f'"{name}" was already running')
            continue

        reason = err or out or f'Return code {code}'
        f_log('WARN', 'SERVICE', f'Failed starting "{name}": {reason}')

# ---------------------------------------------
# Docker compose operations
# ---------------------------------------------
def stop_docker(cfg):
    docker_root = Path(cfg['docker_root'])

    for name in cfg['docker_compose_names']:
        compose = docker_root / name / 'docker-compose.yml'
        f_log('INFO', 'DOCKER', f'Stopping compose "{name}" -> "{compose}"...')

        if not compose.exists():
            f_log('ERRO', 'DOCKER', f'Compose file not found')
            sys.exit(1)

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" stop')
        if ok:
            f_log('DONE', 'DOCKER', f'Stopped "{name}"')
        else:
            reason = err or out or f'Return code {code}'
            f_log('ERRO', 'DOCKER', f'Failed stopping "{name}": {reason}')
            sys.exit(1)

def start_docker(cfg):
    docker_root = Path(cfg['docker_root'])

    for name in cfg['docker_compose_names']:
        compose = docker_root / name / 'docker-compose.yml'
        f_log('INFO', 'DOCKER', f'Starting compose "{name}" -> "{compose}"...')

        if not compose.exists():
            f_log('WARN', 'DOCKER', f'Compose file missing')
            continue

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" start')
        if ok:
            f_log('DONE', 'DOCKER', f'Started "{name}"')
        else:
            reason = err or out or f'Return code {code}'
            f_log('WARN', 'DOCKER', f'Failed starting "{name}": {reason}')

# ---------------------------------------------
# ZIP helper
# ---------------------------------------------
def zip_folder(src: Path, dst_zip: Path):
    f_log('INFO', 'BACKUP', f'Compressing "{src}"...')

    with zipfile.ZipFile(dst_zip, 'w', compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
        if src.is_file():
            zf.write(src, arcname=src.name)
        else:
            for p in src.rglob('*'):
                if p.is_file():
                    zf.write(p, arcname=str(p.relative_to(src)))

    f_log('DONE', 'BACKUP', f'Created "{dst_zip}"')

# ---------------------------------------------
# Backup paths
# ---------------------------------------------
def backup_all_paths(cfg, timestamp):
    backup_root = Path(cfg['backup_root'])
    ts_folder = backup_root / timestamp

    # Services
    for svc in cfg['services']:
        for src in svc['paths']:
            src_path = Path(src)
            if not src_path.exists():
                f_log('WARN', 'BACKUP', f'Skipped missing path: "{src_path}"')
                continue

            drive = src_path.drive.replace(':', '')
            rel = src_path.relative_to(src_path.anchor)  # full path after drive

            dst_dir = ts_folder / drive / rel.parent
            dst_dir.mkdir(parents=True, exist_ok=True)

            dst_zip = dst_dir / (src_path.name + '.zip')
            zip_folder(src_path, dst_zip)

    # Docker
    docker_root = Path(cfg['docker_root'])
    for name in cfg['docker_compose_names']:
        src = docker_root / name
        if not src.exists():
            f_log('WARN', 'BACKUP', f'Skipped missing path: "{src}"')
            continue

        drive = src.drive.replace(':', '')
        rel = src.relative_to(src.anchor)

        dst_dir = ts_folder / drive / rel.parent
        dst_dir.mkdir(parents=True, exist_ok=True)

        dst_zip = dst_dir / (name + '.zip')
        zip_folder(src, dst_zip)

# ---------------------------------------------
# Retention pruning
# ---------------------------------------------
def prune_versions(cfg):
    root = Path(cfg['backup_root'])
    days = cfg['retention_days']
    min_v = cfg['retention_min_versions']
    max_v = cfg['retention_max_versions']

    # timestamp directories only
    versions = []
    for d in root.iterdir():
        if not d.is_dir():
            continue
        try:
            t = datetime.strptime(d.name, '%Y-%m-%d %H%M%S')
            versions.append((d, t))
        except:
            continue

    if not versions:
        return

    versions.sort(key=lambda x: x[1])  # oldest first
    now = datetime.now()

    to_delete = []

    # delete by days
    for d, t in versions:
        if (now - t).days > days:
            to_delete.append((d, t))

    # enforce min versions
    remain = len(versions) - len(to_delete)
    if remain < min_v:
        need_keep = min_v - remain
        to_delete = to_delete[:-need_keep] if need_keep < len(to_delete) else []

    # enforce max versions
    if max_v > 0 and len(versions) > max_v:
        excess = len(versions) - max_v
        extra = versions[:excess]  # oldest
        extra_dirs = {d for d, t in extra}
        for d, t in versions:
            if d in extra_dirs and (d, t) not in to_delete:
                to_delete.append((d, t))

    # remove
    for d, t in to_delete:
        f_log('INFO', 'PRUNE', f'Deleting old version "{t.strftime('%Y-%m-%d %H%M%S')}"...')
        try:
            shutil.rmtree(d)
            f_log('DONE', 'PRUNE', f'Deleted "{d}"')
        except Exception as e:
            f_log('WARN', 'PRUNE', f'Failed deleting "{d}": {e}')

# ---------------------------------------------
# Full backup process
# ---------------------------------------------
def do_backup(cfg):
    backup_root = Path(cfg['backup_root'])
    ts = timestamp_folder()
    ts_folder = backup_root / ts
    ts_folder.mkdir(parents=True, exist_ok=True)

    # Open log file
    global LOG_FILE_HANDLE, CURRENT_LOG_PATH
    CURRENT_LOG_PATH = ts_folder / 'log.txt'
    LOG_FILE_HANDLE = open(CURRENT_LOG_PATH, 'w', encoding='utf-8')

    # Save config snapshot
    f_log('INFO', 'CONFIG', 'Copying config...')
    try:
        shutil.copy2(CONFIG_FILE, ts_folder / 'config.json')
        f_log('INFO', 'CONFIG', f'Created "{ts_folder / 'config.json'}"')
    except Exception as e:
        f_log('WARN', 'CONFIG', f'Failed creating config: {e}')

    backup_all_paths(cfg, ts)

# ---------------------------------------------
# Main
# ---------------------------------------------
def main():
    log('''Windows Service Backup
The MIT License (MIT)
Copyright (c) 2025 Jonathan Chiu
''')

    if len(sys.argv) == 1:
        f_log('ERRO', 'MAIN', 'Missing command. Run "python main.py help" for usage.')
        sys.exit(1)

    mode = sys.argv[1]
    if sys.argv[1] not in COMMANDS:
        f_log('ERRO', 'MAIN', f'Unknown command "{sys.argv[1]}". Run "python main.py help" for usage.')
        sys.exit(1)

    cfg = {}

    if mode == 'help':
        log(USAGE)
    else:
        cfg = load_config()

    if mode == 'backup' or mode == 'restart' or mode == 'restartServices' or mode == 'stop' or mode == 'stopServices':
        stop_services(cfg)
    if mode == 'backup' or mode == 'restart' or mode == 'restartServices' or mode == 'stop' or mode == 'stopDocker':
        stop_docker(cfg)

    if mode == 'backup':
        do_backup(cfg)

    if mode == 'backup' or mode == 'restart' or mode == 'restartServices' or mode == 'start' or mode == 'startServices':
        start_services(cfg)
    if mode == 'backup' or mode == 'restart' or mode == 'restartDocker' or mode == 'start' or mode == 'startServices':
        start_docker(cfg)

    if mode == 'backup' or mode == 'prune':
        prune_versions(cfg)

    log('''
Run complete.''')

    global LOG_FILE_HANDLE
    if LOG_FILE_HANDLE:
        LOG_FILE_HANDLE.close()
        LOG_FILE_HANDLE = None

if __name__ == '__main__':
    main()
