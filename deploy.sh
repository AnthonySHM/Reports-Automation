#!/bin/bash
# Quick deployment script for SHM Report Generator
# Usage: sudo ./deploy.sh your-domain.com your-email@example.com

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

if [ "$EUID" -ne 0 ]; then 
    echo -e "${RED}Please run as root (sudo)${NC}"
    exit 1
fi

if [ $# -ne 2 ]; then
    echo "Usage: $0 <domain> <email>"
    echo "Example: $0 reports.example.com admin@example.com"
    exit 1
fi

DOMAIN=$1
EMAIL=$2
APP_DIR="/var/www/shm-reports"
APP_USER="www-data"

echo -e "${GREEN}🚀 Starting deployment for ${DOMAIN}${NC}"

# 1. Install dependencies
echo -e "${YELLOW}📦 Installing system dependencies...${NC}"
apt update
apt install -y python3 python3-pip python3-venv nginx certbot python3-certbot-nginx

# 2. Setup application directory
echo -e "${YELLOW}📁 Setting up application directory...${NC}"
mkdir -p $APP_DIR
cd $APP_DIR

# 3. Setup Python environment
echo -e "${YELLOW}🐍 Setting up Python environment...${NC}"
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# 4. Configure systemd service
echo -e "${YELLOW}⚙️  Configuring systemd service...${NC}"
cat > /etc/systemd/system/shm-reports.service <<EOF
[Unit]
Description=SHM Report Generator API
After=network.target

[Service]
Type=simple
User=$APP_USER
Group=$APP_USER
WorkingDirectory=$APP_DIR
Environment="PATH=$APP_DIR/venv/bin"
Environment="PYTHONUNBUFFERED=1"
Environment="SHM_GOOGLE_CREDENTIALS=$APP_DIR/Credentials/focal-set-486609-s9-f052f8c1a756.json"
ExecStart=$APP_DIR/venv/bin/python run.py serve --production --port 5000 --workers 4

Restart=always
RestartSec=5
NoNewPrivileges=true
PrivateTmp=true

StandardOutput=journal
StandardError=journal
SyslogIdentifier=shm-reports

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable shm-reports

# 5. Configure nginx
echo -e "${YELLOW}🌐 Configuring nginx...${NC}"
cat > /etc/nginx/sites-available/shm-reports <<EOF
upstream shm_app {
    server 127.0.0.1:5000;
}

server {
    listen 80;
    server_name $DOMAIN;
    
    location / {
        proxy_pass http://shm_app;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        
        proxy_connect_timeout 600;
        proxy_send_timeout 600;
        proxy_read_timeout 600;
        send_timeout 600;
    }
    
    client_max_body_size 100M;
}
EOF

ln -sf /etc/nginx/sites-available/shm-reports /etc/nginx/sites-enabled/
nginx -t
systemctl reload nginx

# 6. Fix permissions
echo -e "${YELLOW}🔒 Setting permissions...${NC}"
chown -R $APP_USER:$APP_USER $APP_DIR
chmod 600 $APP_DIR/Credentials/*.json 2>/dev/null || true

# 7. Start application
echo -e "${YELLOW}▶️  Starting application...${NC}"
systemctl start shm-reports
sleep 2
systemctl status shm-reports --no-pager

# 8. Setup SSL with Let's Encrypt
echo -e "${YELLOW}🔐 Setting up SSL certificate...${NC}"
certbot --nginx -d $DOMAIN --non-interactive --agree-tos --email $EMAIL --redirect

# 9. Setup firewall
echo -e "${YELLOW}🔥 Configuring firewall...${NC}"
ufw allow 'Nginx Full'
ufw --force enable

# 10. Final checks
echo -e "${GREEN}✅ Deployment complete!${NC}"
echo ""
echo -e "${GREEN}Your application is now running at: https://$DOMAIN${NC}"
echo ""
echo "Useful commands:"
echo "  View logs:        sudo journalctl -u shm-reports -f"
echo "  Restart service:  sudo systemctl restart shm-reports"
echo "  Check status:     sudo systemctl status shm-reports"
echo "  Nginx logs:       sudo tail -f /var/log/nginx/shm-reports-*.log"
echo ""
echo "Test your deployment:"
echo "  curl https://$DOMAIN/"
