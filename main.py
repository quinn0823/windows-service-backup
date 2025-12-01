# Jellytion
# The MIT License (MIT)
# Copyright (c) 2025 Jonathan Chiu

import os
import subprocess
import json
import time
import zipfile
from datetime import datetime, timedelta
from pathlib import Path


# --------------------------------------------------------------
# 工具函数
# --------------------------------------------------------------

def run_cmd(cmd, timeout=60):
    """运行命令并阻塞等待。失败时打印警告但不中断程序。"""
    print(f"[CMD] {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=timeout)
        if result.returncode == 0:
            print("[OK]")
        else:
            print("[WARN] Command failed:", result.stderr.strip())
    except subprocess.TimeoutExpired:
        print(f"[ERROR] Command timeout: {cmd}")


def stop_service(name):
    print(f"[INFO] 停止 NSSM 服务: {name}")
    run_cmd(f'net stop "{name}"', timeout=30)


def start_service(name):
    print(f"[INFO] 启动 NSSM 服务: {name}")
    run_cmd(f'net start "{name}"', timeout=30)


def stop_docker_project(root, name):
    compose_path = os.path.join(root, name, "docker-compose.yml")
    print(f"[INFO] 停止 Docker 项目: {name}")
    run_cmd(f'docker compose -f "{compose_path}" stop', timeout=45)


def start_docker_project(root, name):
    compose_path = os.path.join(root, name, "docker-compose.yml")
    print(f"[INFO] 启动 Docker 项目: {name}")
    run_cmd(f'docker compose -f "{compose_path}" start', timeout=45)


def zip_folder(src_path, dst_zip):
    """压缩 src_path 到 dst_zip，ZIP 内文件不包含多余目录层级。"""
    src_path = Path(src_path)
    dst_zip = Path(dst_zip)
    dst_zip.parent.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] 压缩 {src_path} → {dst_zip}")

    with zipfile.ZipFile(dst_zip, 'w', compression=zipfile.ZIP_STORED) as zipf:
        if src_path.is_file():
            zipf.write(src_path, arcname=src_path.name)
        else:
            for root, _, files in os.walk(src_path):
                root_path = Path(root)
                for file in files:
                    file_path = root_path / file
                    arcname = file_path.relative_to(src_path)
                    zipf.write(file_path, arcname)

    print("[OK] 压缩完成")


def cleanup_old_backups(parent_dir, retention_days, retention_min):
    """保留最近 retention_days 天，但至少保留 retention_min 份"""
    print(f"[INFO] 清理旧备份：{parent_dir}")
    parent = Path(parent_dir)
    if not parent.exists():
        return

    dirs = sorted([d for d in parent.iterdir() if d.is_dir()])

    # 解析日期
    def parse_time(name):
        try:
            return datetime.strptime(name, "%Y-%m-%d %H%M%S")
        except:
            return None

    backup_list = [(d, parse_time(d.name)) for d in dirs]
    backup_list = [(d, t) if t else None for d, t in backup_list if t]

    if not backup_list:
        return

    # 应保留的日期下限
    threshold = datetime.now() - timedelta(days=retention_days)

    # 需要删除的目录
    to_delete = [d for d, t in backup_list if t < threshold]

    # 保证至少 retention_min 份
    remaining = len(backup_list) - len(to_delete)
    if remaining < retention_min:
        to_delete = to_delete[:max(0, len(backup_list) - retention_min)]

    for d in to_delete:
        print(f"[INFO] 删除旧备份: {d}")
        try:
            import shutil
            shutil.rmtree(d)
        except Exception as e:
            print(f"[ERROR] 删除失败: {d} - {e}")


# --------------------------------------------------------------
# 主程序
# --------------------------------------------------------------

def main():
    # 读取配置文件
    with open("config.json", "r", encoding="utf-8") as f:
        cfg = json.load(f)

    backup_root = cfg["backup_root"]
    retention_days = cfg["retention_days"]
    retention_min = cfg["retention_min_copies"]

    timestamp = datetime.now().strftime("%Y-%m-%d %H%M%S")

    # -------------------------------
    # 1. 停止 NSSM 服务
    # -------------------------------
    for item in cfg["nssm"]:
        stop_service(item["service"])

    # -------------------------------
    # 2. 停止 Docker 项目
    # -------------------------------
    docker_root = cfg["docker_root"]
    for name in cfg["docker"]:
        stop_docker_project(docker_root, name)

    # -------------------------------
    # 3. 开始备份各路径（NSSM + Docker）
    # -------------------------------
    for item in cfg["nssm"]:
        svc = item["service"]
        for path in item["paths"]:
            p = Path(path)
            if not p.exists():
                print(f"[WARN] 路径不存在: {path}")
                continue

            drive = p.drive.replace(":", "")
            relative_dir = p.parent.name
            zip_name = p.name + ".zip"

            dst = Path(backup_root) / drive / timestamp / relative_dir / zip_name
            zip_folder(path, dst)

    # Docker 数据备份
    for name in cfg["docker"]:
        path = Path(docker_root) / name
        if path.exists():
            drive = path.drive.replace(":", "")
            relative_dir = path.parent.name
            zip_name = name + ".zip"

            dst = Path(backup_root) / drive / timestamp / relative_dir / zip_name
            zip_folder(path, dst)

    # -------------------------------
    # 4. 重启 NSSM 服务
    # -------------------------------
    for item in cfg["nssm"]:
        start_service(item["service"])

    # -------------------------------
    # 5. 重启 Docker
    # -------------------------------
    for name in cfg["docker"]:
        start_docker_project(docker_root, name)

    # -------------------------------
    # 6. 清理旧备份（按盘符）
    # -------------------------------
    backup_root_path = Path(backup_root)
    for drive_dir in backup_root_path.iterdir():
        if drive_dir.is_dir():
            cleanup_old_backups(drive_dir, retention_days, retention_min)

    print("[DONE] 备份任务完成")


if __name__ == "__main__":
    main()
