# Production Deployment Guide

This guide covers deploying the SHM Report Generator API to a production server with HTTPS.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Server Setup](#server-setup)
3. [SSL Certificate Setup](#ssl-certificate-setup)
4. [Deployment Options](#deployment-options)
5. [Testing](#testing)
6. [Maintenance](#maintenance)

---

## Prerequisites

- Ubuntu/Debian server (or Windows Server)
- Domain name pointing to your server
- Root/sudo access
- Python 3.9+

---

## Server Setup

### 1. Update System

```bash
sudo apt update && sudo apt upgrade -y
```

### 2. Install Dependencies

```bash
# Python and essentials
sudo apt install -y python3 python3-pip python3-venv nginx certbot python3-certbot-nginx

# Optional: Git for deployment
sudo apt install -y git
```

### 3. Create Application Directory

```bash
sudo mkdir -p /var/www/shm-reports
sudo chown $USER:$USER /var/www/shm-reports
cd /var/www/shm-reports
```

### 4. Deploy Application

#### Option A: Git Clone (Recommended)

```bash
git clone <your-repo-url> .
```

#### Option B: Upload Files

Upload your application files via SCP/SFTP to `/var/www/shm-reports`

### 5. Setup Python Environment

```bash
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### 6. Configure Credentials

```bash
# Copy your Google service account JSON
mkdir -p Credentials
# Upload your credentials file to Credentials/

# Set environment variable (optional if using default path)
export SHM_GOOGLE_CREDENTIALS=/var/www/shm-reports/Credentials/focal-set-486609-s9-f052f8c1a756.json
```

---

## SSL Certificate Setup

### Option 1: Let's Encrypt (Free, Recommended)

#### Step 1: Install Certbot

```bash
sudo apt install certbot python3-certbot-nginx
```

#### Step 2: Obtain Certificate

```bash
sudo certbot --nginx -d your-domain.com -d www.your-domain.com
```

Follow the prompts:
- Enter email address
- Agree to terms
- Choose whether to redirect HTTP to HTTPS (recommended: Yes)

Certificates will be saved to:
- Certificate: `/etc/letsencrypt/live/your-domain.com/fullchain.pem`
- Private Key: `/etc/letsencrypt/live/your-domain.com/privkey.pem`

#### Step 3: Auto-Renewal

Certbot automatically sets up renewal. Test it:

```bash
sudo certbot renew --dry-run
```

### Option 2: Commercial Certificate

If you purchased a commercial SSL certificate:

1. Save certificate files:
   ```bash
   sudo mkdir -p /etc/ssl/certs
   sudo mkdir -p /etc/ssl/private
   sudo cp your-certificate.crt /etc/ssl/certs/
   sudo cp your-private-key.key /etc/ssl/private/
   sudo chmod 600 /etc/ssl/private/your-private-key.key
   ```

2. Update nginx configuration with your paths

---

## Deployment Options

### Option A: Nginx Reverse Proxy (Recommended)

This is the **recommended approach** for production.

#### 1. Configure Nginx

```bash
sudo cp nginx.conf.example /etc/nginx/sites-available/shm-reports
sudo nano /etc/nginx/sites-available/shm-reports
```

Update the following:
- `server_name your-domain.com` → your actual domain
- SSL certificate paths (if not using default Let's Encrypt paths)
- `proxy_pass http://127.0.0.1:5000` → confirm port matches your app

#### 2. Enable Site

```bash
sudo ln -s /etc/nginx/sites-available/shm-reports /etc/nginx/sites-enabled/
sudo nginx -t  # Test configuration
sudo systemctl reload nginx
```

#### 3. Setup Systemd Service

```bash
sudo cp shm-reports.service.example /etc/systemd/system/shm-reports.service
sudo nano /etc/systemd/system/shm-reports.service
```

Update the following:
- `WorkingDirectory` → your app path
- `User` and `Group` → appropriate user (e.g., `www-data`)
- Environment variables

#### 4. Start Service

```bash
sudo systemctl daemon-reload
sudo systemctl enable shm-reports
sudo systemctl start shm-reports
sudo systemctl status shm-reports
```

#### 5. Check Logs

```bash
# Application logs
sudo journalctl -u shm-reports -f

# Nginx logs
sudo tail -f /var/log/nginx/shm-reports-access.log
sudo tail -f /var/log/nginx/shm-reports-error.log
```

### Option B: Direct HTTPS (No Nginx)

Use built-in SSL support without reverse proxy.

#### 1. Set Certificate Paths

```bash
export SSL_CERT_PATH=/etc/letsencrypt/live/your-domain.com/fullchain.pem
export SSL_KEY_PATH=/etc/letsencrypt/live/your-domain.com/privkey.pem
```

#### 2. Run Production Server

```bash
cd /var/www/shm-reports
source venv/bin/activate
python run.py serve --production --port 443 --ssl-cert $SSL_CERT_PATH --ssl-key $SSL_KEY_PATH --workers 4
```

**Note:** Running on port 443 requires root/sudo. Better to use port 5000 behind nginx.

### Option C: Windows Server (IIS)

For Windows deployment:

#### 1. Install Dependencies

```powershell
pip install -r requirements.txt
```

#### 2. Run with Waitress

```powershell
python run.py serve --production --port 5000
```

#### 3. Configure IIS

- Install IIS
- Install [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite) module
- Add reverse proxy rule to forward requests to port 5000
- Configure SSL certificate in IIS

---

## Testing

### 1. Test Application

```bash
# From server
curl -k https://localhost:5000/

# From external
curl https://your-domain.com/
```

### 2. Generate Test Report

```bash
curl -X POST https://your-domain.com/api/generate \
  -H "Content-Type: application/json" \
  -d '{"client_name": "Test Client"}'
```

### 3. SSL Test

Check SSL configuration: https://www.ssllabs.com/ssltest/

---

## Maintenance

### Update Application

```bash
cd /var/www/shm-reports
git pull  # or upload new files
source venv/bin/activate
pip install -r requirements.txt
sudo systemctl restart shm-reports
```

### View Logs

```bash
# Application logs
sudo journalctl -u shm-reports -f

# Nginx logs
sudo tail -f /var/log/nginx/shm-reports-*.log

# Check service status
sudo systemctl status shm-reports
```

### Backup

```bash
# Backup credentials
sudo tar -czf shm-reports-backup.tar.gz /var/www/shm-reports/Credentials

# Backup generated reports
sudo tar -czf reports-backup.tar.gz /var/www/shm-reports/output
```

### Firewall Configuration

```bash
# Allow HTTP and HTTPS
sudo ufw allow 'Nginx Full'
sudo ufw enable
sudo ufw status
```

---

## Performance Tuning

### Worker Processes

Adjust based on CPU cores:

```bash
# Rule of thumb: (2 × CPU_cores) + 1
python run.py serve --production --workers 8
```

### Nginx Optimization

Add to nginx.conf:

```nginx
worker_processes auto;
worker_connections 1024;
keepalive_timeout 65;
```

---

## Troubleshooting

### Application Won't Start

```bash
# Check logs
sudo journalctl -u shm-reports -n 50

# Test manually
cd /var/www/shm-reports
source venv/bin/activate
python run.py serve --production
```

### SSL Certificate Errors

```bash
# Verify certificates exist
sudo ls -l /etc/letsencrypt/live/your-domain.com/

# Test nginx config
sudo nginx -t

# Renew certificate
sudo certbot renew
```

### Permission Issues

```bash
# Fix ownership
sudo chown -R www-data:www-data /var/www/shm-reports

# Fix credentials permissions
sudo chmod 600 /var/www/shm-reports/Credentials/*.json
```

---

## Security Checklist

- [ ] SSL/TLS enabled (HTTPS)
- [ ] Firewall configured (UFW/iptables)
- [ ] Service runs as non-root user
- [ ] Credentials file permissions set to 600
- [ ] Auto-renewal enabled for Let's Encrypt
- [ ] Security headers configured in nginx
- [ ] Regular backups scheduled
- [ ] System updates automated

---

## Quick Reference Commands

```bash
# Start/Stop/Restart
sudo systemctl start shm-reports
sudo systemctl stop shm-reports
sudo systemctl restart shm-reports

# View logs
sudo journalctl -u shm-reports -f

# Reload nginx
sudo systemctl reload nginx

# Renew SSL
sudo certbot renew

# Check status
sudo systemctl status shm-reports
sudo systemctl status nginx
```

---

## Support

For issues or questions:
1. Check logs first: `sudo journalctl -u shm-reports -n 100`
2. Verify nginx config: `sudo nginx -t`
3. Test SSL: https://www.ssllabs.com/ssltest/
4. Review this guide for common solutions
