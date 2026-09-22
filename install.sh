#!/usr/bin/env bash
# Установка бота на сервер (Ubuntu/Debian). Запускать из-под root:
#   sudo bash install.sh
set -euo pipefail

ЦЕЛЬ=/opt/prosrochka_bot
ПОЛЬЗОВАТЕЛЬ=prosrochka
ОТКУДА="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "==> Проверяю Python"
if ! command -v python3 >/dev/null; then
  apt-get update -qq && apt-get install -y python3 python3-venv python3-pip
fi
python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3,8) else 1)' || {
  echo "Нужен Python 3.8 или новее"; exit 1; }

echo "==> Создаю пользователя $ПОЛЬЗОВАТЕЛЬ (без права входа)"
id -u "$ПОЛЬЗОВАТЕЛЬ" >/dev/null 2>&1 || useradd --system --create-home --shell /usr/sbin/nologin "$ПОЛЬЗОВАТЕЛЬ"

echo "==> Копирую файлы в $ЦЕЛЬ"
mkdir -p "$ЦЕЛЬ"
cp -r "$ОТКУДА"/bot.py "$ОТКУДА"/engine.py "$ОТКУДА"/requirements.txt "$ЦЕЛЬ"/
mkdir -p "$ЦЕЛЬ/template" "$ЦЕЛЬ/data"
cp -r "$ОТКУДА"/template/. "$ЦЕЛЬ/template"/

if [ ! -f "$ЦЕЛЬ/config.ini" ]; then
  cp "$ОТКУДА"/config.example.ini "$ЦЕЛЬ"/config.ini
  echo "    создан $ЦЕЛЬ/config.ini — его нужно заполнить (токен и список ID)"
else
  echo "    config.ini уже есть, не трогаю"
fi

echo "==> Ставлю зависимости в отдельное окружение"
python3 -m venv "$ЦЕЛЬ/venv"
"$ЦЕЛЬ/venv/bin/pip" install --quiet --upgrade pip
"$ЦЕЛЬ/venv/bin/pip" install --quiet -r "$ЦЕЛЬ/requirements.txt"

chown -R "$ПОЛЬЗОВАТЕЛЬ:$ПОЛЬЗОВАТЕЛЬ" "$ЦЕЛЬ"
chmod 600 "$ЦЕЛЬ/config.ini"

echo "==> Ставлю автозапуск (systemd)"
cp "$ОТКУДА"/prosrochka-bot.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable prosrochka-bot >/dev/null

cat <<TEXT

Готово. Осталось два шага:

  1) Впишите токен и список Telegram ID:
       nano $ЦЕЛЬ/config.ini

  2) Запустите бота:
       systemctl start prosrochka-bot

Полезное:
  systemctl status prosrochka-bot     — работает или нет
  journalctl -u prosrochka-bot -f     — смотреть журнал вживую
  systemctl restart prosrochka-bot    — перезапустить после правки config.ini

Бот стартует сам после перезагрузки сервера.
TEXT
