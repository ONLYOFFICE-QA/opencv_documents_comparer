# -*- coding: utf-8 -*-
import requests
from loguru import logger
import os


class Telegram:
    @staticmethod
    def send_message(text: str):
        telegram_token = os.environ.get('TELEGRAM_TOKEN')
        channel_id = os.environ.get('CHANNEL_ID')
        if telegram_token and channel_id:
            url = f"https://api.telegram.org/bot{telegram_token}/sendMessage"
            response = requests.post(url, data={"chat_id": channel_id, "text": text})
            if response.status_code != 200:
                logger.error(f"post_text error text: {text}")
