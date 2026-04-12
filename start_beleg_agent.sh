#!/bin/bash
cd "$(dirname "$0")"
.venv/bin/python3 tray_agent.py &
disown
