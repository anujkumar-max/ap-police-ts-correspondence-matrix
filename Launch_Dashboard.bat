@echo off
title AP Police TS Correspondence Executive Analytics Dashboard
echo =========================================================================
echo  AP POLICE - TECHNICAL SERVICES (PCS&S)
echo  Executive Correspondence & Tappals Analytics Dashboard
echo =========================================================================
echo.
echo Starting Live Web Dashboard Server on http://localhost:8050 ...
echo Tracking Excel File: AP TS-Correspondence MATRIX.xlsx
echo.

cd /d "%~dp0"
start "" http://localhost:8050
python server.py

pause
