# -*- coding: utf-8 -*-
import requests
from loguru import logger
import os

from configuration import version
from data.StaticData import StaticData
from framework.fileutils import FileUtils


class Telegram:
    telegram_token, channel_id = os.environ.get('TELEGRAM_TOKEN'), os.environ.get('CHANNEL_ID')

    @staticmethod
    def send_message(message: str):
        if Telegram.telegram_token and Telegram.channel_id:
            if len(message) > 4096:
                return Telegram.send_long_message_as_document(message)
            url = f"https://api.telegram.org/bot{Telegram.telegram_token}/sendMessage"
            response = requests.post(url, data={"chat_id": Telegram.channel_id, "text": message})
            if response.status_code != 200:
                logger.error(f"post_text error text: {message}")

    @staticmethod
    def send_document(path_to_document, caption=''):
        if Telegram.telegram_token and Telegram.channel_id:
            response = requests.post(f"https://api.telegram.org/bot{Telegram.telegram_token}/sendDocument",
                                     data={"chat_id": Telegram.channel_id, "caption": caption},
                                     files={"document": open(path_to_document, 'rb')})
            if response.status_code != 200:
                logger.error(f"Error when sending a document")

    @staticmethod
    def send_long_message_as_document(long_message: str):
        if Telegram.telegram_token and Telegram.channel_id:
            report = f"{StaticData.TMP_DIR}/report.txt"
            with open(report, "w") as file:
                file.write(long_message)
                file.close()
            Telegram.send_document(report, caption=f"report.txt\nVersion:{version}")
            FileUtils.delete(report)
