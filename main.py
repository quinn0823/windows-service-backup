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
CONFIG_FILE = "config.json"
LOG_FILE_HANDLE = None
CURRENT_LOG_PATH = None

# ---------------------------------------------
# Utility: timestamp
# ---------------------------------------------
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def timestamp_folder():
    return datetime.now().strftime("%Y-%m-%d %H%M%S")

# ---------------------------------------------
# Logging
# ---------------------------------------------
def log(level, comp, msg):
    """Print + optional log file write."""
    line = f"{now()} [{level}] {comp:<7} {msg}"
    print(line)

    if LOG_FILE_HANDLE:
        LOG_FILE_HANDLE.write(line + "\n")
        LOG_FILE_HANDLE.flush()

# ---------------------------------------------
# Load config
# ---------------------------------------------
def load_config():
    cfg_path = Path(CONFIG_FILE)
    if not cfg_path.exists():
        print(f"[ERRO] CONFIG config.json not found: {cfg_path.resolve()}")
        sys.exit(1)

    try:
        return json.loads(cfg_path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[ERRO] CONFIG Failed to parse config.json: {e}")
        sys.exit(1)

# ---------------------------------------------
# Run shell command
# ---------------------------------------------
def run_cmd(cmd, timeout=None):
    try:
        p = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=timeout)
        return p.returncode == 0, (p.stdout or "").strip(), (p.stderr or "").strip(), p.returncode
    except Exception as e:
        return False, "", str(e), -1

# ---------------------------------------------
# Service operations
# ---------------------------------------------
def stop_services(cfg):
    for svc in cfg["services"]:
        name = svc["name"]
        log("INFO", "SERVICE", f'Stopping "{name}"')

        ok, out, err, code = run_cmd(f'net stop "{name}"')
        if ok:
            log("DONE", "SERVICE", f'Stopped "{name}"')
            continue

        combined = out + err
        if "3521" in combined:
            log("DONE", "SERVICE", f'"{name}" was not running')
            continue

        reason = err or out or f"Return code {code}"
        log("ERRO", "SERVICE", f'Failed stopping "{name}": {reason}')
        sys.exit(1)

def start_services(cfg):
    for svc in cfg["services"]:
        name = svc["name"]
        log("INFO", "SERVICE", f'Starting "{name}"')

        ok, out, err, code = run_cmd(f'net start "{name}"')
        if ok:
            log("DONE", "SERVICE", f'Started "{name}"')
            continue

        combined = out + err
        if "2182" in combined:
            log("DONE", "SERVICE", f'"{name}" was already running')
            continue

        reason = err or out or f"Return code {code}"
        log("WARN", "SERVICE", f'Failed starting "{name}": {reason}')

# ---------------------------------------------
# Docker compose operations
# ---------------------------------------------
def stop_docker(cfg):
    docker_root = Path(cfg["docker_root"])

    for name in cfg["docker_compose_names"]:
        compose = docker_root / name / "docker-compose.yml"
        log("INFO", "DOCKER", f'Stopping compose "{name}" -> {compose}')

        if not compose.exists():
            log("ERRO", "DOCKER", f'Compose file not found: {compose}')
            sys.exit(1)

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" stop')
        if ok:
            log("DONE", "DOCKER", f'Stopped "{name}"')
        else:
            reason = err or out or f"Return code {code}"
            log("ERRO", "DOCKER", f'Failed stopping "{name}": {reason}')
            sys.exit(1)

def start_docker(cfg):
    docker_root = Path(cfg["docker_root"])

    for name in cfg["docker_compose_names"]:
        compose = docker_root / name / "docker-compose.yml"
        log("INFO", "DOCKER", f'Starting compose "{name}" -> {compose}')

        if not compose.exists():
            log("WARN", "DOCKER", f'Compose file missing: {compose}')
            continue

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" start')
        if ok:
            log("DONE", "DOCKER", f'Started "{name}"')
        else:
            reason = err or out or f"Return code {code}"
            log("WARN", "DOCKER", f'Failed starting "{name}": {reason}')

# ---------------------------------------------
# ZIP helper
# ---------------------------------------------
def zip_folder(src: Path, dst_zip: Path):
    log("INFO", "BACKUP", f'Compressing "{src}"')

    with zipfile.ZipFile(dst_zip, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
        if src.is_file():
            zf.write(src, arcname=src.name)
        else:
            for p in src.rglob("*"):
                if p.is_file():
                    zf.write(p, arcname=str(p.relative_to(src)))

    log("DONE", "BACKUP", f'Created {dst_zip}')

# ---------------------------------------------
# Backup paths
# ---------------------------------------------
def backup_all_paths(cfg, timestamp):
    backup_root = Path(cfg["backup_root"])
    ts_folder = backup_root / timestamp

    # Services
    for svc in cfg["services"]:
        for src in svc["paths"]:
            src_path = Path(src)
            if not src_path.exists():
                log("WARN", "BACKUP", f'Skipped missing path: {src_path}')
                continue

            drive = src_path.drive.replace(":", "")
            rel = src_path.relative_to(src_path.anchor)  # full path after drive

            dst_dir = ts_folder / drive / rel.parent
            dst_dir.mkdir(parents=True, exist_ok=True)

            dst_zip = dst_dir / (src_path.name + ".zip")
            zip_folder(src_path, dst_zip)

    # Docker
    docker_root = Path(cfg["docker_root"])
    for name in cfg["docker_compose_names"]:
        src = docker_root / name
        if not src.exists():
            log("WARN", "BACKUP", f'Skipped missing docker path: {src}')
            continue

        drive = src.drive.replace(":", "")
        rel = src.relative_to(src.anchor)

        dst_dir = ts_folder / drive / rel.parent
        dst_dir.mkdir(parents=True, exist_ok=True)

        dst_zip = dst_dir / (name + ".zip")
        zip_folder(src, dst_zip)

# ---------------------------------------------
# Retention pruning
# ---------------------------------------------
def prune_versions(cfg):
    root = Path(cfg["backup_root"])
    days = cfg["retention_days"]
    min_v = cfg["retention_min_versions"]
    max_v = cfg["retention_max_versions"]

    # timestamp directories only
    versions = []
    for d in root.iterdir():
        if not d.is_dir():
            continue
        try:
            t = datetime.strptime(d.name, "%Y-%m-%d %H%M%S")
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
        log("INFO", "CLEANUP", f"Deleting old version {d}")
        try:
            shutil.rmtree(d)
            log("DONE", "CLEANUP", f"Deleted {d}")
        except Exception as e:
            log("WARN", "CLEANUP", f"Failed deleting {d}: {e}")

# ---------------------------------------------
# Full backup process
# ---------------------------------------------
def do_backup(cfg):

    backup_root = Path(cfg["backup_root"])
    ts = timestamp_folder()
    ts_folder = backup_root / ts
    ts_folder.mkdir(parents=True, exist_ok=True)

    # Open log file
    global LOG_FILE_HANDLE, CURRENT_LOG_PATH
    CURRENT_LOG_PATH = ts_folder / "log.txt"
    LOG_FILE_HANDLE = open(CURRENT_LOG_PATH, "w", encoding="utf-8")

    # Save config snapshot
    try:
        shutil.copy2(CONFIG_FILE, ts_folder / "config.json")
        log("INFO", "CONFIG", "Saved config.json snapshot")
    except Exception as e:
        log("WARN", "CONFIG", f"Failed saving config snapshot: {e}")

    # Stop → Backup → Start
    stop_services(cfg)
    stop_docker(cfg)

    backup_all_paths(cfg, ts)

    start_services(cfg)
    start_docker(cfg)

    # Prune old versions
    prune_versions(cfg)

    LOG_FILE_HANDLE.close()
    LOG_FILE_HANDLE = None

# ---------------------------------------------
# Main
# ---------------------------------------------
def main():
    print("""Windows Service Backup
The MIT License (MIT)
Copyright (c) 2025 Jonathan Chiu
""")

    usage = """Usage:
    python main.py [command]

Commands:
    backup           Perform backup of services and docker
    help             Show this help message
    start            Start services and docker containers
    startDocker      Start docker containers only
    startServices    Start services only
    stop             Stop services and docker containers
    stopDocker       Stop docker containers only
    stopServices     Stop services only"""

    if len(sys.argv) < 2:
        print("[INFO] Run \"python main.py help\" for usage.")
    else:
        mode = sys.argv[1].lower()
        cfg = load_config()

        if mode == "backup":
            do_backup(cfg)

        elif mode == "help":
            print(usage)

        elif mode == "start":
            start_services(cfg)
            start_docker(cfg)
        elif mode == "startDocker":
            start_docker(cfg)
        elif mode == "startServices":
            start_services(cfg)

        elif mode == "stop":
            stop_services(cfg)
            stop_docker(cfg)
        elif mode == "stopDocker":
            stop_docker(cfg)
        elif mode == "stopServices":
            stop_services(cfg)

        else:
            print(f"""[ERRO] Unknown command "{mode}"
    [INFO] Run "python main.py help" for usage.""")

    print("""
Run complete.""")

if __name__ == "__main__":
    main()
