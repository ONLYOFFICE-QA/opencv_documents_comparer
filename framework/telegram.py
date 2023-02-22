# -*- coding: utf-8 -*-
import os
from os.path import join, getsize, basename, isdir

import requests
from loguru import logger

from configurations.project_configurator import ProjectConfig
from framework.FileUtils import FileUtils
from settings import version


class Telegram:
    telegram_token, chat_id = os.environ.get('TELEGRAM_TOKEN'), os.environ.get('CHANNEL_ID')

    @staticmethod
    def send_message(message: str):
        if Telegram.telegram_token and Telegram.chat_id:
            if len(message) > 4096:
                return Telegram.send_long_message_as_document(message)
            url = f"https://api.telegram.org/bot{Telegram.telegram_token}/sendMessage"
            response = requests.post(url, data={"chat_id": Telegram.chat_id, "text": message, "parse_mode": "Markdown"})
            if response.status_code != 200:
                logger.error(f"post_text error text: {message}")

    @staticmethod
    def send_document(document_path, caption=''):
        if Telegram.telegram_token and Telegram.chat_id:
            if getsize(document_path) >= 50_000_000 or isdir(document_path):
                FileUtils.compress_files(document_path, join(ProjectConfig.TMP_DIR, f'{basename(document_path)}.zip'))
                document_path = join(ProjectConfig.TMP_DIR, f'{basename(document_path)}.zip')
            try:
                response = requests.post(f"https://api.telegram.org/bot{Telegram.telegram_token}/sendDocument",
                                         data={"chat_id": Telegram.chat_id, "caption": caption,
                                               "parse_mode": "Markdown"},
                                         files={"document": open(document_path, 'rb')})
                if response.status_code != 200:
                    logger.error(f"Error when sending a document")
            except Exception as e:
                Telegram.send_message(f"|WARNING|Impossible to send: {document_path}. Error: {e}")

    @staticmethod
    def send_long_message_as_document(long_message: str):
        if Telegram.telegram_token and Telegram.chat_id:
            report = join(ProjectConfig.TMP_DIR, 'report.txt')
            with open(report, "w") as file:
                file.write(long_message)
            Telegram.send_document(report, caption=f"report.txt\nVersion:{version}")
            FileUtils.delete(report)
