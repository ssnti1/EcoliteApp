#!/bin/bash
# iniciar servidor VNC + noVNC
Xvfb :0 -screen 0 1024x768x16 &
fluxbox &
x11vnc -display :0 -forever -nopw -rfbport 5900 &
websockify --web=/usr/share/novnc/ 8080 localhost:5900 &

# lanzar tu app en el display virtual
export DISPLAY=:0
python main.py
chmod +x start.sh
