@echo off
chcp 65001 >nul
title PDF to Markdown Converter
echo Starting PDF to Markdown Converter...
echo.
cd /d "g:\내 드라이브\RPA\RAG\pdf_to_markdown"
echo Current directory: %CD%
echo.
cmd /k python main.py
