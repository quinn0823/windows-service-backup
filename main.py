# Windows Service Backup
# The MIT License (MIT)
# Copyright (c) 2025 Jonathan Chiu

# Run as administrator on Windows.

import os
import sys
import json
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
import zipfile
import ctypes

# -------------------------
# Configuration
# -------------------------
CONFIG_FILE = "config.json"
TOOL_NAME = "windows-service-backup"
LOG_FILENAME = "log.txt"
CONFIG_SNAPSHOT_NAME = "config.json"
TIME_FMT = "%Y-%m-%d %H%M%S"  # e.g., 2025-10-28 171611

# -------------------------
# Logging utilities
# -------------------------
class Logger:
    def __init__(self, console=True):
        self.lines = []
        self.console = console

    def _now(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _fmt(self, level, comp, msg):
        # level and comp padded
        return f"{self._now()} [{level}] {comp:<7} {msg}"

    def info(self, comp, msg):
        line = self._fmt("INFO", comp, msg)
        self.lines.append(line)
        if self.console:
            print(line)

    def okay(self, comp, msg):
        line = self._fmt("OKAY", comp, msg)
        self.lines.append(line)
        if self.console:
            print(line)

    def warn(self, comp, msg):
        line = self._fmt("WARN", comp, msg)
        self.lines.append(line)
        if self.console:
            print(line)

    def erro(self, comp, msg):
        line = self._fmt("ERRO", comp, msg)
        self.lines.append(line)
        if self.console:
            print(line)

    def dump_to_file(self, path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self.lines) + "\n")
        except Exception as e:
            # last resort: print
            print(f"Failed to write log to {path}: {e}")

# -------------------------
# Privilege check
# -------------------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

# -------------------------
# Command runner
# -------------------------
def run_cmd(cmd, timeout=None):
    """
    Run a shell command and return (success: bool, stdout, stderr, returncode)
    """
    try:
        res = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=timeout)
        return (res.returncode == 0, res.stdout or "", res.stderr or "", res.returncode)
    except subprocess.TimeoutExpired as e:
        return (False, "", f"Timeout: {e}", -1)
    except Exception as e:
        return (False, "", f"Exception: {e}", -1)

# -------------------------
# Service operations
# -------------------------
def stop_service(logger, name, wait_sec=30):
    logger.info("SERVICE", f'Stopping "{name}"')
    ok, out, err, code = run_cmd(f'net stop "{name}"', timeout=wait_sec)
    if ok:
        logger.okay("SERVICE", f'Stopped "{name}"')
        # double-check service state
        return wait_for_service_state(logger, name, "STOPPED", timeout=10)
    else:
        # Collect a human-friendly message
        stderr = (err or out).strip()
        if not stderr:
            stderr = f"Return code: {code}"
        logger.erro("SERVICE", f'Failed stopping "{name}": {stderr}')
        return False

def start_service(logger, name, wait_sec=30):
    logger.info("SERVICE", f'Starting "{name}"')
    ok, out, err, code = run_cmd(f'net start "{name}"', timeout=wait_sec)
    if ok:
        logger.okay("SERVICE", f'Started "{name}"')
        return wait_for_service_state(logger, name, "RUNNING", timeout=10)
    else:
        stderr = (err or out).strip()
        if not stderr:
            stderr = f"Return code: {code}"
        logger.erro("SERVICE", f'Failed starting "{name}": {stderr}')
        return False

def wait_for_service_state(logger, name, desired_state, timeout=10):
    """
    desired_state examples: 'STOPPED', 'RUNNING'
    Poll sc query for state.
    """
    end = datetime.now() + timedelta(seconds=timeout)
    while datetime.now() < end:
        ok, out, err, _ = run_cmd(f'sc query "{name}"')
        if ok and out:
            # parse output, look for "STATE"
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("STATE"):
                    # e.g. STATE              : 4  RUNNING
                    parts = line.split()
                    if parts and desired_state in line:
                        return True
        # sleep a short time
        time_sleep(1)
    logger.warn("SERVICE", f'Timeout waiting for "{name}" to reach {desired_state}')
    # final check once
    ok, out, err, _ = run_cmd(f'sc query "{name}"')
    if ok and out and desired_state in out:
        return True
    return False

# -------------------------
# Docker operations
# -------------------------
def stop_docker_compose_project(logger, docker_root, name, wait_sec=60):
    compose_path = Path(docker_root) / name / "docker-compose.yml"
    logger.info("DOCKER", f'Stopping compose project "{name}" using "{compose_path}"')
    if not compose_path.exists():
        logger.erro("DOCKER", f'Docker compose file not found: {compose_path}')
        return False
    ok, out, err, code = run_cmd(f'docker compose -f "{compose_path}" stop', timeout=wait_sec)
    if ok:
        logger.okay("DOCKER", f'Stopped "{name}"')
        return True
    else:
        stderr = (err or out).strip()
        if not stderr:
            stderr = f"Return code: {code}"
        logger.erro("DOCKER", f'Failed stopping "{name}": {stderr}')
        return False

def start_docker_compose_project(logger, docker_root, name, wait_sec=60):
    compose_path = Path(docker_root) / name / "docker-compose.yml"
    logger.info("DOCKER", f'Starting compose project "{name}" using "{compose_path}"')
    if not compose_path.exists():
        logger.warn("DOCKER", f'Docker compose file not found (start skipped): {compose_path}')
        return False
    ok, out, err, code = run_cmd(f'docker compose -f "{compose_path}" start', timeout=wait_sec)
    if ok:
        logger.okay("DOCKER", f'Started "{name}"')
        return True
    else:
        stderr = (err or out).strip()
        if not stderr:
            stderr = f"Return code: {code}"
        logger.erro("DOCKER", f'Failed starting "{name}": {stderr}')
        return False

# -------------------------
# Utilities
# -------------------------
def time_sleep(s):
    try:
        import time
        time.sleep(s)
    except Exception:
        pass

def ensure_dir(path):
    Path(path).mkdir(parents=True, exist_ok=True)

# -------------------------
# ZIP utilities (no compression)
# -------------------------
def zip_path_no_compress(logger, src, dst_zip_path):
    """
    Zip src (file or directory) into dst_zip_path.
    The zip will not include an extra top-level directory: entries are relative to src.
    dst_zip_path must end with .zip
    """
    src = Path(src)
    dst_zip_path = Path(dst_zip_path)
    ensure_dir(dst_zip_path.parent)

    logger.info("BACKUP", f'Compressing "{src}" -> "{dst_zip_path.name}"')
    try:
        with zipfile.ZipFile(dst_zip_path, mode="w", compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
            if src.is_file():
                zf.write(src, arcname=src.name)
            else:
                for root, _, files in os.walk(src):
                    root_path = Path(root)
                    for f in files:
                        file_path = root_path / f
                        # arcname relative to src (so no top-level folder)
                        arcname = file_path.relative_to(src)
                        zf.write(file_path, arcname)
        logger.okay("BACKUP", f'Created: {dst_zip_path}')
        return True
    except Exception as e:
        logger.erro("BACKUP", f'Failed to create zip for "{src}": {e}')
        return False

# -------------------------
# Retention / cleanup
# -------------------------
def cleanup_retention(logger, drive_dir: Path, retention_days: int, min_copies: int):
    """
    drive_dir contains timestamped directories (format TIME_FMT).
    Keep those within retention_days, but at least min_copies.
    Oldest removed first.
    """
    logger.info("CLEANUP", f'Checking retention in "{drive_dir}"')
    if not drive_dir.exists() or not drive_dir.is_dir():
        return

    entries = [d for d in drive_dir.iterdir() if d.is_dir()]
    # filter valid timestamp dirs and sort ascending (oldest first)
    parsed = []
    for d in entries:
        try:
            t = datetime.strptime(d.name, TIME_FMT)
            parsed.append((d, t))
        except Exception:
            # ignore non-timestamp directories
            continue
    parsed.sort(key=lambda x: x[1])

    if not parsed:
        return

    now = datetime.now()
    to_delete = []
    # find those older than retention_days
    for d, t in parsed:
        age_days = (now - t).days
        if age_days > retention_days:
            to_delete.append((d, t))

    # ensure at least min_copies remain
    remaining = len(parsed) - len(to_delete)
    if remaining < min_copies:
        # calculate how many we must keep from to_delete
        need = min_copies - remaining
        # remove newest from to_delete
        if need >= len(to_delete):
            to_delete = []
        else:
            # keep the latest "need" from to_delete, so drop the rest
            to_delete = to_delete[:-need]

    # Delete directories
    for d, t in to_delete:
        logger.info("CLEANUP", f'Deleting old backup: {d.name}')
        try:
            shutil.rmtree(d)
            logger.okay("CLEANUP", f'Deleted {d.name}')
        except Exception as e:
            logger.warn("CLEANUP", f'Failed to delete {d}: {e}')

# -------------------------
# Main flow
# -------------------------
def main():
    logger = Logger(console=True)
    logger.info("CORE", f"{TOOL_NAME} starting")

    # require admin
    if not is_admin():
        logger.erro("CORE", "Administrator privileges required. Please run as Administrator.")
        sys.exit(1)

    # load config
    cfg_path = Path(CONFIG_FILE)
    if not cfg_path.exists():
        logger.erro("CONFIG", f"Config file not found: {cfg_path.resolve()}")
        sys.exit(1)

    try:
        cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    except Exception as e:
        logger.erro("CONFIG", f"Failed to parse config: {e}")
        sys.exit(1)

    backup_root = Path(cfg.get("backup_root", "")).resolve()
    retention_days = int(cfg.get("retention_days", 7))
    retention_min = int(cfg.get("retention_min_copies", cfg.get("retention_min_copies", 3)))
    nssm_list = cfg.get("nssm", [])
    docker_list = cfg.get("docker", [])
    docker_root = Path(cfg.get("docker_root", "C:\\ProgramData\\Docker"))

    if not backup_root:
        logger.erro("CONFIG", "backup_root not set in config.json")
        sys.exit(1)

    # build timestamp and prepare per-drive dirs map
    timestamp = datetime.now().strftime(TIME_FMT)

    # First: stop all NSSM services, abort on any failure
    for item in nssm_list:
        service_name = item.get("service")
        if not service_name:
            continue
        ok = stop_service(logger, service_name)
        if not ok:
            logger.erro("CORE", f'Backup aborted due to failure stopping service "{service_name}"')
            # write a quick log file in current working directory for debugging
            try:
                local_log = Path.cwd() / LOG_FILENAME
                logger.dump_to_file(local_log)
            except Exception:
                pass
            sys.exit(1)

    # Next: stop docker compose projects, abort on any failure
    for name in docker_list:
        ok = stop_docker_compose_project(logger, docker_root, name)
        if not ok:
            logger.erro("CORE", f'Backup aborted due to failure stopping docker "{name}"')
            try:
                local_log = Path.cwd() / LOG_FILENAME
                logger.dump_to_file(local_log)
            except Exception:
                pass
            sys.exit(1)

    # All stopped successfully -> proceed to backup
    logger.info("CORE", "All services and docker projects stopped. Starting backup.")

    # We'll collect which drives we wrote to so we can run retention per-drive
    drives_touched = set()

    # For each NSSM item, compress each path
    for item in nssm_list:
        service_name = item.get("service")
        paths = item.get("paths", [])
        for p in paths:
            src = Path(p)
            if not src.exists():
                logger.warn("BACKUP", f'Source path does not exist (skipped): "{src}"')
                continue

            drive_letter = src.drive.replace(":", "") or "unknown"
            drives_touched.add(drive_letter)

            # target timestamp dir: <backup_root>\<drive_letter>\<timestamp>\
            dst_base = backup_root / drive_letter / timestamp
            # ensure directory exists
            ensure_dir(dst_base)

            # create consistent subfolder: use parent folder name (as earlier spec)
            parent_name = src.parent.name if src.parent.name else "root"
            target_subdir = dst_base / parent_name
            ensure_dir(target_subdir)

            zip_name = f"{src.name}.zip"
            dst_zip = target_subdir / zip_name

            # copy config snapshot into dst_base (once per drive)
            try:
                snapshot_path = dst_base / CONFIG_SNAPSHOT_NAME
                if not snapshot_path.exists():
                    # write the original config file content (not modified)
                    shutil.copy2(cfg_path, snapshot_path)
                    logger.info("CONFIG", f"Wrote config snapshot to {snapshot_path}")
            except Exception as e:
                logger.warn("CONFIG", f"Failed to write config snapshot: {e}")

            # perform zip (no compression)
            ok = zip_path_no_compress(logger, src, dst_zip)
            if not ok:
                logger.erro("CORE", f'Backup failed when compressing "{src}". Aborting.')
                # attempt to write logs locally and exit
                try:
                    logger.dump_to_file(Path.cwd() / LOG_FILENAME)
                except Exception:
                    pass
                sys.exit(1)

    # Docker data: compress each docker project folder under docker_root
    for name in docker_list:
        src = docker_root / name
        if not src.exists():
            logger.warn("BACKUP", f'Docker data path not found (skipped): "{src}"')
            continue

        drive_letter = src.drive.replace(":", "") or "unknown"
        drives_touched.add(drive_letter)

        dst_base = backup_root / drive_letter / timestamp
        ensure_dir(dst_base)

        # use "ProgramData" as parent name to match earlier examples if docker_root includes ProgramData
        parent_name = src.parent.name if src.parent.name else "docker"
        target_subdir = dst_base / parent_name
        ensure_dir(target_subdir)

        zip_name = f"{name}.zip"
        dst_zip = target_subdir / zip_name

        # write config snapshot if not already
        try:
            snapshot_path = dst_base / CONFIG_SNAPSHOT_NAME
            if not snapshot_path.exists():
                shutil.copy2(cfg_path, snapshot_path)
                logger.info("CONFIG", f"Wrote config snapshot to {snapshot_path}")
        except Exception as e:
            logger.warn("CONFIG", f"Failed to write config snapshot: {e}")

        ok = zip_path_no_compress(logger, src, dst_zip)
        if not ok:
            logger.erro("CORE", f'Backup failed when compressing "{src}". Aborting.')
            try:
                logger.dump_to_file(Path.cwd() / LOG_FILENAME)
            except Exception:
                pass
            sys.exit(1)

    # After backing up, attempt to restart services and docker (log errors but do not abort)
    logger.info("CORE", "Backup phase complete. Restarting services and docker projects.")

    for item in nssm_list:
        service_name = item.get("service")
        if not service_name:
            continue
        ok = start_service(logger, service_name)
        if not ok:
            logger.erro("SERVICE", f'Failed to start service "{service_name}" (check manually)')

    for name in docker_list:
        ok = start_docker_compose_project(logger, docker_root, name)
        if not ok:
            logger.erro("DOCKER", f'Failed to start docker project "{name}" (check manually)')

    # Retention cleanup per-drive
    for drive_letter in drives_touched:
        drive_dir = backup_root / drive_letter
        cleanup_retention(logger, drive_dir, retention_days, retention_min)

    # Write log files into each created timestamp directory (and also current working dir)
    for drive_letter in drives_touched:
        dst_base = backup_root / drive_letter / timestamp
        try:
            ensure_dir(dst_base)
            log_path = dst_base / LOG_FILENAME
            logger.dump_to_file(log_path)
            logger.info("CORE", f"Wrote log to {log_path}")
        except Exception as e:
            logger.warn("CORE", f"Failed to write log to {dst_base}: {e}")

    # Also write a local log copy
    try:
        local_log = Path.cwd() / LOG_FILENAME
        logger.dump_to_file(local_log)
    except Exception:
        pass

    logger.okay("CORE", "Backup job finished")

if __name__ == "__main__":
    main()
