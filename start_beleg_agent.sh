#!/bin/bash
cd "$(dirname "$0")"
python3 tray_agent.py &
disown
