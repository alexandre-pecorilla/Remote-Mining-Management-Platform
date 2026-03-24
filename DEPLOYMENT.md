# Deployment Guide

This guide covers deploying Mining Manager on a VPS (public internet) or a local server (LAN). Both scenarios use Nginx as a reverse proxy in front of the Dockerized application.

- **VPS deployment** — uses a domain name and Let's Encrypt (Certbot) for TLS.
- **LAN deployment** — uses a self-signed certificate with OpenSSL.

## Prerequisites

- Debian 13 or Ubuntu 24.04 server
- Root or sudo access
- Docker and Docker Compose installed
- The application repository cloned on the server

### Install Docker

```bash
# Install Docker
curl -fsSL https://get.docker.com | sh

# Add your user to the docker group (log out and back in after)
sudo usermod -aG docker $USER
```

### Install Nginx and Certbot

```bash
sudo apt update
sudo apt install -y nginx certbot python3-certbot-nginx
```

### Clone and configure the application

```bash
git clone <your-repo-url> /opt/mining
cd /opt/mining
cp .env.example .env
```

Edit `.env` with production values:

```env
DJANGO_SECRET_KEY=<generate-a-random-key>
DJANGO_DEBUG=False
DJANGO_ALLOWED_HOSTS=your-domain.com,your-server-ip
CSRF_TRUSTED_ORIGINS=https://your-domain.com,https://your-server-ip
COINMARKETCAP_API_KEY=your-api-key
APP_PASSWORD=your-password
POSTGRES_PASSWORD=a-strong-random-password
GUNICORN_WORKERS=3
WEB_PORT=8000
```

> Generate a secret key with:
> ```bash
> python3 -c "import secrets; print(secrets.token_urlsafe(50))"
> ```

**`DJANGO_ALLOWED_HOSTS` must include every hostname or IP that clients will use to reach the app.** Django rejects requests with unrecognized `Host` headers. Examples:

```env
# VPS with a domain
DJANGO_ALLOWED_HOSTS=mining.example.com
CSRF_TRUSTED_ORIGINS=https://mining.example.com

# LAN with the server's IP and hostname
DJANGO_ALLOWED_HOSTS=localhost,127.0.0.1,192.168.1.50,myserver,myserver.lan
CSRF_TRUSTED_ORIGINS=https://192.168.1.50,https://myserver,https://myserver.lan
```

`CSRF_TRUSTED_ORIGINS` is required when serving over HTTPS behind a reverse proxy (Nginx). Each entry must include the `https://` prefix.

Build and start the application:

```bash
docker compose up -d --build
```

Verify it's running:

```bash
curl -s -o /dev/null -w "%{http_code}" http://localhost:8000
```

A `200` or `302` response means the app is up. A `302` is expected if `APP_PASSWORD` is set (redirect to login).

---

## Option A: VPS Deployment (Public Domain + Let's Encrypt)

This assumes you have a domain name (e.g., `mining.example.com`) with DNS pointing to your server's public IP.

### 1. Configure Nginx

Create the Nginx site configuration:

```bash
sudo nano /etc/nginx/sites-available/mining
```

Paste the following:

```nginx
server {
    listen 80;
    server_name mining.example.com;

    client_max_body_size 20M;

    location /static/ {
        alias /opt/mining/staticfiles/;
    }

    location /media/ {
        alias /opt/mining/media/;
    }

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Replace `mining.example.com` with your actual domain.

Enable the site and test:

```bash
sudo ln -s /etc/nginx/sites-available/mining /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

### 2. Obtain TLS certificate with Certbot

```bash
sudo certbot --nginx -d mining.example.com
```

Certbot will:
- Verify domain ownership
- Obtain a certificate from Let's Encrypt
- Automatically modify the Nginx config to enable HTTPS
- Set up auto-renewal

Verify auto-renewal is working:

```bash
sudo certbot renew --dry-run
```

### 3. Verify

Open `https://mining.example.com` in your browser. You should see a valid HTTPS connection with the Let's Encrypt certificate.

---

## Option B: LAN Deployment (Self-Signed Certificate)

This is for servers on a local network without a public domain name. You'll access the app via IP address (e.g., `https://192.168.1.50`).

### 1. Generate a self-signed certificate

Create a directory for the certificates:

```bash
sudo mkdir -p /etc/nginx/ssl
```

Generate a Certificate Authority (CA) key and certificate. This CA will sign your server certificate, and you'll import the CA into your browsers to trust it.

```bash
# Generate CA private key
sudo openssl genrsa -out /etc/nginx/ssl/ca.key 4096

# Generate CA certificate (valid 10 years)
sudo openssl req -x509 -new -nodes \
    -key /etc/nginx/ssl/ca.key \
    -sha256 -days 3650 \
    -out /etc/nginx/ssl/ca.crt \
    -subj "/C=US/ST=Local/L=Local/O=Mining Manager CA/CN=Mining Manager CA"
```

Generate the server certificate signed by your CA:

```bash
# Generate server private key
sudo openssl genrsa -out /etc/nginx/ssl/server.key 2048

# Create a config file for the certificate with Subject Alternative Names
sudo tee /etc/nginx/ssl/server.cnf > /dev/null << 'EOF'
[req]
default_bits = 2048
prompt = no
distinguished_name = dn
req_extensions = v3_req

[dn]
C = US
ST = Local
L = Local
O = Mining Manager
CN = mining-server

[v3_req]
subjectAltName = @alt_names

[alt_names]
# Add your server's IP address(es) and/or hostname(s) here
IP.1 = 192.168.1.50
IP.2 = 127.0.0.1
DNS.1 = mining-server
DNS.2 = mining-server.local
EOF
```

> **Important:** Edit `server.cnf` and replace `192.168.1.50` with your server's actual LAN IP. Add additional IPs or hostnames as needed.

```bash
# Generate Certificate Signing Request (CSR)
sudo openssl req -new \
    -key /etc/nginx/ssl/server.key \
    -out /etc/nginx/ssl/server.csr \
    -config /etc/nginx/ssl/server.cnf

# Sign the server certificate with your CA (valid 2 years)
sudo openssl x509 -req \
    -in /etc/nginx/ssl/server.csr \
    -CA /etc/nginx/ssl/ca.crt \
    -CAkey /etc/nginx/ssl/ca.key \
    -CAcreateserial \
    -out /etc/nginx/ssl/server.crt \
    -days 730 \
    -sha256 \
    -extensions v3_req \
    -extfile /etc/nginx/ssl/server.cnf

# Lock down permissions
sudo chmod 600 /etc/nginx/ssl/*.key
```

### 2. Configure Nginx

```bash
sudo nano /etc/nginx/sites-available/mining
```

Paste the following:

```nginx
server {
    listen 80;
    server_name _;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name _;

    ssl_certificate /etc/nginx/ssl/server.crt;
    ssl_certificate_key /etc/nginx/ssl/server.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;

    client_max_body_size 20M;

    location /static/ {
        alias /opt/mining/staticfiles/;
    }

    location /media/ {
        alias /opt/mining/media/;
    }

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Enable the site and restart Nginx:

```bash
sudo ln -s /etc/nginx/sites-available/mining /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl reload nginx
```

### 3. Import the CA certificate into browsers

Copy `ca.crt` from the server to your client machine:

```bash
# From your client machine
scp user@192.168.1.50:/etc/nginx/ssl/ca.crt ~/Downloads/mining-ca.crt
```

#### Google Chrome / Chromium / Edge / Brave

These browsers all use the operating system's certificate store.

**Windows:**
1. Double-click `mining-ca.crt`
2. Click **Install Certificate...**
3. Select **Local Machine** (or Current User), click Next
4. Select **Place all certificates in the following store**
5. Click Browse, select **Trusted Root Certification Authorities**
6. Click Next, then Finish
7. Restart Chrome

**macOS:**
1. Double-click `mining-ca.crt` — it opens in Keychain Access
2. It will be added to the **login** keychain
3. Find "Mining Manager CA" in the list, double-click it
4. Expand **Trust**, set **When using this certificate** to **Always Trust**
5. Close the window, enter your password to confirm
6. Restart Chrome

**Linux:**
```bash
# Debian/Ubuntu
sudo cp mining-ca.crt /usr/local/share/ca-certificates/mining-ca.crt
sudo update-ca-certificates

# Then restart Chrome/Chromium
```

#### Mozilla Firefox

Firefox uses its own certificate store and does not use the OS store.

1. Open Firefox, go to `about:preferences#privacy`
2. Scroll down to **Certificates**, click **View Certificates...**
3. Go to the **Authorities** tab
4. Click **Import...**, select `mining-ca.crt`
5. Check **Trust this CA to identify websites**
6. Click OK
7. Restart Firefox

#### Safari (macOS)

Safari uses the macOS system keychain. Follow the same macOS steps described above for Chrome.

### 4. Verify

Open `https://192.168.1.50` (your server's LAN IP) in your browser. After importing the CA certificate, you should see a valid HTTPS connection with no warnings.

---

## Production Hardening Checklist

After deployment, verify these settings:

- [ ] `DJANGO_DEBUG=False` in `.env`
- [ ] `DJANGO_SECRET_KEY` is a unique random value
- [ ] `POSTGRES_PASSWORD` is a strong random password
- [ ] `APP_PASSWORD` is set (if you want login protection)
- [ ] `DJANGO_ALLOWED_HOSTS` includes your domain/IP
- [ ] Firewall allows only ports 80, 443, and SSH (22)

### Firewall setup (UFW)

```bash
sudo ufw allow 22/tcp
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw enable
```

> **Note:** Do not expose port 8000 to the public. Nginx proxies traffic to it internally.

---

## Updating the Application

```bash
cd /opt/mining
git pull
docker compose up -d --build
```

The entrypoint script automatically runs migrations and collects static files on every container start. Static files are bind-mounted to the host, so Nginx picks up changes automatically.

---

## Backup & Restore

### Backup the database

```bash
docker compose exec db pg_dump -U mining mining > backup_$(date +%Y%m%d).sql
```

### Restore from backup

```bash
cat backup_20260323.sql | docker compose exec -T db psql -U mining mining
```

### Backup media files

```bash
tar czf media_backup_$(date +%Y%m%d).tar.gz media/
```
