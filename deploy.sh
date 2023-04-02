#!/bin/bash
#Стопанем службы иначе исполняемые файлы не заменятся новыми
systemctl stop EasySlave

#Копируем файлы в целевой каталог
cp -v -r ./_example /srv/EasySlave/
cp -v ./*.py /srv/EasySlave/
cp -v ./requirements.txt /srv/EasySlave/

#Создаём глобальную переменную с токеном
cd /etc
echo "BOT_API_KEY=$BOT_API_KEY" > environment

# Создаём сервис
cd /etc/systemd/system
echo "[Unit]
Description=EasySlave
After=network.target

[Service]
Environment="BOT_API_KEY=$BOT_API_KEY"
Type=simple
User=root
ExecStart=/srv/EasySlave/venv/bin/python /srv/EasySlave/bot.py
KillMode=process
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target" > EasySlave.service

#Устанавливаем менеджер пакетов и модуль виртуального окружения
apt --yes install python3-venv python3-pip

#Создаём виртуальное окружение и устанавливаем пакеты
python3 -m venv /srv/EasySlave/venv
cd /srv/EasySlave
source venv/bin/activate
pip install -r requirements.txt

#Запускаем сервис
systemctl enable EasySlave
systemctl start EasySlave
