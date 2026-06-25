#!/bin/bash
# Сторож транскрибера: пингует оператора в Telegram, когда минуты speech2text
# пополнены (лимит больше не на нуле). Срабатывает один раз и удаляет себя из cron.
cd /opt/transcriber || exit 0
read KEY TOKEN PROXY <<< "$(/opt/transcriber/venv/bin/python -c "import json;c=json.load(open('config.json'));print(c['stt_api_key'],c['telegram_bot_token'],c['telegram_proxy'])" 2>/dev/null)"
[ -z "$KEY" ] && exit 0
CHAT=YOUR_TELEGRAM_CHAT_ID   # ваш Telegram chat_id для авто-уведомлений
MIN=$(curl -s -m 20 "https://speech2text.ru/api/user/amounts?api-key=$KEY" \
      | /opt/transcriber/venv/bin/python -c "import sys,json;print(json.load(sys.stdin)['minutes']['available'])" 2>/dev/null)
[ -z "$MIN" ] && exit 0
if [ "$MIN" -gt 20 ]; then
  curl -s -m 30 --proxy "$PROXY" "https://api.telegram.org/bot$TOKEN/sendMessage" \
    --data-urlencode "chat_id=$CHAT" \
    --data-urlencode "text=✅ Транскрибер снова работает: минуты speech2text пополнены (доступно ${MIN} мин). Кидай записи — обработаются. (автоуведомление)" >/dev/null 2>&1
  # одноразово — убираем себя из cron
  crontab -l 2>/dev/null | grep -v 'stt_notify.sh' | crontab - 2>/dev/null
fi
exit 0
