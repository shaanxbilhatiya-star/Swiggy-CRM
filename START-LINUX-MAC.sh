#!/bin/bash

echo ""
echo " ========================================"
echo "  AutoLead Showcaser - Starting Server"
echo " ========================================"
echo ""

# Check Node.js
if ! command -v node &> /dev/null; then
    echo " ERROR: Node.js is not installed!"
    echo " Install it from: https://nodejs.org"
    exit 1
fi

# Install deps if missing
if [ ! -d "node_modules" ]; then
    echo " Installing dependencies..."
    npm install
    echo ""
fi

# Get LAN IP
IP=$(hostname -I 2>/dev/null | awk '{print $1}' || ipconfig getifaddr en0 2>/dev/null || echo "YOUR-LAN-IP")

echo " ========================================"
echo "  Server is RUNNING!"
echo " ========================================"
echo ""
echo "  Admin Panel : http://$IP:3000/admin"
echo "  Agent Panel : http://$IP:3000/agent"
echo ""
echo "  Share the Agent link with your agents."
echo "  Press Ctrl+C to stop."
echo " ========================================"
echo ""

node server.js
