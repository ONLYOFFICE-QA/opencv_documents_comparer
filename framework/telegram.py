# -*- coding: utf-8 -*-
import requests
from loguru import logger
import os
import subprocess as sb

from config import version


class Telegram:
    telegram_token, channel_id = os.environ.get('TELEGRAM_TOKEN'), os.environ.get('CHANNEL_ID')

    @staticmethod
    def send_message(message: str):
        Telegram.send_text_message(message) if len(message) < 4096 else Telegram.send_long_message_as_document(message)

    @staticmethod
    def send_text_message(text: str):
        if Telegram.telegram_token and Telegram.channel_id:
            url = f"https://api.telegram.org/bot{Telegram.telegram_token}/sendMessage"
            response = requests.post(url, data={"chat_id": Telegram.channel_id, "text": text})
            if response.status_code != 200:
                logger.error(f"post_text error text: {text}")

    @staticmethod
    def send_long_message_as_document(long_message: str):
        if Telegram.telegram_token and Telegram.channel_id:
            with open("./report.txt", "w") as file:
                file.write(long_message)
                file.close()
            response = requests.post(f"https://api.telegram.org/bot{Telegram.telegram_token}/sendDocument",
                                     data={"chat_id": Telegram.channel_id, "caption": f"report.txt\nVersion:{version}"},
                                     files={"document": open("./report.txt", 'rb')})
            if os.path.exists('./report.txt'):
                sb.call(f'powershell.exe rm ./report.txt -Force', shell=True)
            if response.status_code != 200:
                logger.error(f"post_text error text: {long_message}")
