@REM Windows Service Backup
@REM The MIT License (MIT)
@REM Copyright (c) 2025 Jonathan Chiu

@echo off

cd /d %~dp0
cd ..

python main.py stop
