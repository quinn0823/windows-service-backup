# Jellytion
# The MIT License (MIT)
# Copyright (c) 2025 Jonathan Chiu

import os
import json
import shutil
import subprocess
import datetime
import sys

# --------------------- 日志函数 ---------------------
def log(msg):
    print(f"[INFO] {msg}")

def log_ok(msg):
    print(f"[OK] {msg}")

def log_warn(msg):
    print(f"[WARN] {msg}")

def log_fatal(msg):
    print(f"[FATAL] {msg}")

# --------------------- 命令执行 ---------------------
def run_cmd(cmd):
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, shell=True
        )
        if result.returncode != 0:
            log_warn(f"Command failed: {result.stdout}{result.stderr}")
            return False
        return True
    except Exception as e:
        log_warn(f"Exception running command: {e}")
        return False

# --------------------- NSSM 服务 ---------------------
def stop_service(service_name):
    log(f"停止 NSSM 服务: {service_name}")
    cmd = f'net stop "{service_name}"'

    ok = run_cmd(cmd)
    if not ok:
        log_fatal(f"无法停止服务：{service_name}")
    return ok

def start_service(service_name):
    log(f"启动 NSSM 服务: {service_name}")
    cmd = f'net start "{service_name}"'
    return run_cmd(cmd)

# --------------------- Docker 操作 ---------------------
def stop_docker(name):
    log(f"停止 Docker 应用: {name}")
    cmd = f"docker stop {name}"

    ok = run_cmd(cmd)
    if not ok:
        log_fatal(f"无法停止 Docker：{name}")
    return ok

def start_docker(name):
    log(f"启动 Docker 应用: {name}")
    cmd = f"docker start {name}"
    return run_cmd(cmd)

# --------------------- 压缩 ---------------------
def zip_path(src, dest_zip):
    log(f"压缩 {src} → {dest_zip}")
    try:
        shutil.make_archive(dest_zip.replace(".zip", ""), "zip", src)
        log_ok("压缩完成")
        return True
    except Exception as e:
        log_warn(f"压缩失败: {e}")
        return False

# --------------------- 目录轮替机制 ---------------------
def rotate_backups(root, retention_days, min_versions):
    entries = sorted(os.listdir(root))
    versions = []

    for e in entries:
        full = os.path.join(root, e)
        if os.path.isdir(full):
            versions.append((e, full))

    # 保持至少 min_versions 份
    if len(versions) <= min_versions:
        return

    now = datetime.datetime.now()
    for dirname, full in versions:
        dt = datetime.datetime.strptime(dirname, "%Y-%m-%d %H%M%S")
        age = (now - dt).days

        if age > retention_days and len(versions) > min_versions:
            log(f"删除过期版本: {dirname}")
            shutil.rmtree(full)
            versions.remove((dirname, full))

# --------------------- 主函数 ---------------------
def main():
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    backup_root = config["backup_root"]
    retention_days = config["retention_days"]
    min_versions = config["min_versions"]

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H%M%S")
    backup_dir = os.path.join(backup_root, timestamp)

    os.makedirs(backup_dir, exist_ok=True)
    os.makedirs(os.path.join(backup_dir, "Program Files"), exist_ok=True)

    # ---------------- NSSM ----------------
    for item in config["nssm"]:
        svc = item["service"]
        paths = item["paths"]

        if not stop_service(svc):
            log_fatal("服务停止失败，终止整个备份流程")
            sys.exit(1)

        for p in paths:
            name = p.replace("C:\\", "").replace("\\", "_")
            dest = os.path.join(backup_dir, name + ".zip")
            zip_path(p, dest)

        start_service(svc)

    # ---------------- Docker ----------------
    for name in config["docker"]:
        if not stop_docker(name):
            log_fatal("Docker 停止失败，终止备份流程")
            sys.exit(1)

        docker_path = f"C:\\ProgramData\\Docker\\{name}"
        dest = os.path.join(backup_dir, f"Docker_{name}.zip")
        zip_path(docker_path, dest)

        start_docker(name)

    # ---------------- 版本轮替 ----------------
    rotate_backups(backup_root, retention_days, min_versions)

    log_ok("备份完成！")

if __name__ == "__main__":
    main()
