#!/bin/bash
source /home/botuser/.bot_env
exec sudo -u botuser /opt/transcriber/venv/bin/python /opt/transcriber/transcriber.py
