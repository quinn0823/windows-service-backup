# Windows Service Backup
# The MIT License (MIT)
# Copyright (c) 2025 Jonathan Chiu

import os
import sys
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
import zipfile

# ---------------------------------------------
# Globals
# ---------------------------------------------
CONFIG_FILE = "config.json"
LOG_FILE_HANDLE = None   # only set in backup mode
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
    """Print and (if in backup mode) also write to log.txt."""
    line = f"{now()} [{level}] {comp:<7} {msg}"
    print(line)

    global LOG_FILE_HANDLE
    if LOG_FILE_HANDLE:
        LOG_FILE_HANDLE.write(line + "\n")
        LOG_FILE_HANDLE.flush()

# ---------------------------------------------
# Load config
# ---------------------------------------------
def load_config():
    cfg_path = Path(CONFIG_FILE)
    if not cfg_path.exists():
        print(f"[ERRO] CONFIG config.json not found in: {cfg_path.resolve()}")
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
        combined = (out + err)
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
        if not ok:
            reason = err or out or f"Return code {code}"
            log("WARN", "SERVICE", f'Failed starting "{name}": {reason}')
        else:
            log("DONE", "SERVICE", f'Started "{name}"')

# ---------------------------------------------
# Docker compose operations
# ---------------------------------------------
def stop_docker(cfg):
    root = Path(cfg["docker_root"])

    for name in cfg["docker_compose_names"]:
        compose = root / name / "docker-compose.yml"
        log("INFO", "DOCKER", f'Stopping compose "{name}" -> {compose}')

        if not compose.exists():
            log("ERRO", "DOCKER", f'Compose file not found: {compose}')
            sys.exit(1)

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" stop')
        if not ok:
            reason = err or out or f"Return code {code}"
            log("ERRO", "DOCKER", f'Failed stopping "{name}": {reason}')
            sys.exit(1)

        log("DONE", "DOCKER", f'Stopped "{name}"')

def start_docker(cfg):
    root = Path(cfg["docker_root"])

    for name in cfg["docker_compose_names"]:
        compose = root / name / "docker-compose.yml"
        log("INFO", "DOCKER", f'Starting compose "{name}" -> {compose}')

        if not compose.exists():
            log("WARN", "DOCKER", f'Compose file missing, skipped: {compose}')
            continue

        ok, out, err, code = run_cmd(f'docker compose -f "{compose}" start')
        if not ok:
            reason = err or out or f"Return code {code}"
            log("WARN", "DOCKER", f'Failed starting "{name}": {reason}')
        else:
            log("DONE", "DOCKER", f'Started "{name}"')

# ---------------------------------------------
# ZIP (no compression, no top folder)
# ---------------------------------------------
def zip_path(src, dst_zip):
    src = Path(src)
    log("INFO", "BACKUP", f'Compressing "{src}"')

    with zipfile.ZipFile(dst_zip, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
        if src.is_file():
            zf.write(src, arcname=src.name)
        else:
            for root, _, files in os.walk(src):
                r = Path(root)
                for f in files:
                    full = r / f
                    arc = full.relative_to(src)
                    zf.write(full, arcname=str(arc))

    log("DONE", "BACKUP", f'Created {dst_zip}')

# ---------------------------------------------
# Backup paths
# ---------------------------------------------
def backup_all_paths(cfg, ts_folder):
    backup_root = Path(cfg["backup_root"])
    timestamp = ts_folder

    drives_used = set()

    # 1) Services
    for svc in cfg["services"]:
        for src in svc["paths"]:
            src_path = Path(src)
            if not src_path.exists():
                log("WARN", "BACKUP", f'Skipped missing path: {src_path}')
                continue

            drive = src_path.drive.replace(":", "")
            drives_used.add(drive)

            dst_base = backup_root / timestamp / drive
            (dst_base).mkdir(parents=True, exist_ok=True)

            # parent folder name
            parent = src_path.parent.name or "root"
            service_dir = dst_base / parent
            service_dir.mkdir(exist_ok=True)

            zip_name = f"{src_path.name}.zip"
            dst_zip = service_dir / zip_name

            zip_path(src_path, dst_zip)

    # 2) Docker
    docker_root = Path(cfg["docker_root"])
    for name in cfg["docker_compose_names"]:
        src = docker_root / name
        if not src.exists():
            log("WARN", "BACKUP", f'Skipped missing docker path: {src}')
            continue

        drive = src.drive.replace(":", "")
        drives_used.add(drive)

        dst_base = backup_root / timestamp / drive
        dst_base.mkdir(parents=True, exist_ok=True)

        parent = src.parent.name or "docker"
        docker_dir = dst_base / parent
        docker_dir.mkdir(exist_ok=True)

        dst_zip = docker_dir / f"{name}.zip"
        zip_path(src, dst_zip)

    return drives_used

# ---------------------------------------------
# Retention pruning
# ---------------------------------------------
def prune_versions(cfg):
    root = Path(cfg["backup_root"])
    days = cfg["retention_days"]
    min_v = cfg["retention_min_versions"]
    max_v = cfg["retention_max_versions"]

    # check each drive folder
    for drive_folder in root.iterdir():
        if not drive_folder.is_dir():
            continue

        # each entry is timestamp folder -> version
        versions = []
        for d in drive_folder.iterdir():
            if not d.is_dir():
                continue
            try:
                t = datetime.strptime(d.name, "%Y-%m-%d %H%M%S")
                versions.append((d, t))
            except:
                continue

        if not versions:
            continue

        # sort oldest first
        versions.sort(key=lambda x: x[1])

        # determine to delete
        to_delete = []

        # remove by days
        now = datetime.now()
        for d, t in versions:
            if (now - t).days > days:
                to_delete.append((d, t))

        # enforce min versions
        remain = len(versions) - len(to_delete)
        if remain < min_v:
            need_keep = min_v - remain
            if need_keep >= len(to_delete):
                to_delete = []
            else:
                to_delete = to_delete[:-need_keep]

        # enforce max versions
        if max_v > 0:
            # if total versions > max_v
            if len(versions) > max_v:
                excess = len(versions) - max_v
                # delete oldest entries
                extra = versions[:excess]
                for d, t in extra:
                    if d not in [x[0] for x in to_delete]:
                        to_delete.append((d, t))

        # perform deletion
        for d, t in to_delete:
            log("INFO", "CLEANUP", f"Deleting old version {d}")
            try:
                shutil.rmtree(d)
                log("DONE", "CLEANUP", f"Deleted {d}")
            except Exception as e:
                log("WARN", "CLEANUP", f"Failed deleting {d}: {e}")

# ---------------------------------------------
# Backup process
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

    # Save config.json snapshot first
    cfg_path = Path(CONFIG_FILE)
    try:
        shutil.copy2(cfg_path, ts_folder / "config.json")
        log("INFO", "CONFIG", "Saved config.json snapshot")
    except Exception as e:
        log("WARN", "CONFIG", f"Failed saving config snapshot: {e}")

    # Stop services & docker
    stop_services(cfg)
    stop_docker(cfg)

    # Backup all paths
    drives_used = backup_all_paths(cfg, ts)

    # Restart (ok if fails)
    start_services(cfg)
    start_docker(cfg)

    # Retention
    prune_versions(cfg)

    log("DONE", "CORE", "Backup completed")

    LOG_FILE_HANDLE.close()
    LOG_FILE_HANDLE = None

# ---------------------------------------------
# Main
# ---------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py [backup | stop | start]")
        sys.exit(0)

    mode = sys.argv[1].lower()
    cfg = load_config()

    # stop only
    if mode == "stop":
        stop_services(cfg)
        stop_docker(cfg)
        print("Done.")
        return

    # start only
    if mode == "start":
        start_services(cfg)
        start_docker(cfg)
        print("Done.")
        return

    # backup
    if mode == "backup":
        do_backup(cfg)
        return

    print(f"Unknown mode: {mode}")

# ---------------------------------------------
if __name__ == "__main__":
    main()
