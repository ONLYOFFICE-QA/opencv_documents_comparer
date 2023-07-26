# -*- coding: utf-8 -*-
import json
import os
from os.path import join, getsize, basename, isdir, isfile, expanduser

import requests
from rich import print
from loguru import logger

from frameworks.decorators import singleton
from frameworks.host_control import FileUtils
import tempfile


@singleton
class Telegram:
    def __init__(
            self,
            token_path: str = None,
            chat_id_path: str = None,
            tmp_dir: str = tempfile.gettempdir()
    ):

        self._telegram_dir = join(expanduser('~'), '.telegram')
        self._telegram_token = self._get_token(token_path)
        self._chat_id = self._get_chat_id(chat_id_path)
        self.tmp_dir = tmp_dir
        FileUtils.create_dir(self.tmp_dir, silence=True)

    def _get_token(self, token_path: str) -> str | None:
        token_file = token_path if token_path else join(self._telegram_dir, 'token')
        if token_file and isfile(token_file):
            return FileUtils.file_reader(token_file).strip()
        elif os.environ.get('TELEGRAM_TOKEN', False):
            return os.environ.get('TELEGRAM_TOKEN')
        print(f"[cyan]|INFO| Telegram token not exists.")

    def _get_chat_id(self, chat_id_path: str) -> str | None:
        chat_id_file = chat_id_path if chat_id_path else join(self._telegram_dir, 'chat')
        if isfile(chat_id_file):
            return FileUtils.file_reader(chat_id_file).strip()
        elif os.environ.get('CHANNEL_ID', False):
            return os.environ.get('CHANNEL_ID')
        print(f"[cyan]|INFO| Telegram chat id not exists.")

    def send_message(self, message: str, out_msg=False) -> None:
        print(message) if out_msg else ...
        if self._access:
            if len(message) > 4096:
                document = self._make_massage_doc(message=message)
                self.send_document(document, caption=self._prepare_caption(message))
                return FileUtils.delete(document)
            self._request(
                f"https://api.telegram.org/bot{self._telegram_token}/sendMessage",
                data={"chat_id": self._chat_id, "text": message, "parse_mode": "Markdown"},
                tg_log=False
            )

    def send_document(self, document_path: str, caption: str = '') -> None:
        if self._access:
            self._request(
                f"https://api.telegram.org/bot{self._telegram_token}/sendDocument",
                data={"chat_id": self._chat_id, "caption": self._prepare_caption(caption), "parse_mode": "Markdown"},
                files={"document": open(self._prepare_documents(document_path), 'rb')}
            )

    def send_media_group(self, document_paths: list, caption: str = None, media_type: str = 'document') -> None:
        if self._access:
            if caption and len(caption) > 200:
                document_paths.append(self._make_massage_doc(caption, 'caption.txt'))
            files, media = {}, []
            for doc_path in document_paths:
                files[basename(doc_path)] = open(self._prepare_documents(doc_path), 'rb')
                media.append(dict(type=media_type, media=f'attach://{basename(doc_path)}'))
            media[-1]['caption'] = self._prepare_caption(caption) if caption is not None else ''
            self._request(
                f'https://api.telegram.org/bot{self._telegram_token}/sendMediaGroup',
                data={'chat_id': self._chat_id, 'media': json.dumps(media), "parse_mode": "Markdown"},
                files=files
            )

    @staticmethod
    def _prepare_caption(caption: str) -> str:
        return caption[:200]

    def _request(self, url: str, data: dict, files: dict = None, tg_log: bool = True) -> None:
        try:
            response = requests.post(url, data=data, files=files)
            if response.status_code != 200:
                logger.error(f"Error when sending a document")
                self.send_message(f"Error when sending to telegram: {response.status_code}") if tg_log else ...
        except Exception as e:
            logger.error(f"|WARNING|Impossible to send: {data}. Error: {e}")
            self.send_message(f"|WARNING|Impossible to send: {data}. Error: {e}") if tg_log else ...

    def _prepare_documents(self, doc_path: str) -> str:
        if getsize(doc_path) >= 50_000_000 or isdir(doc_path):
            FileUtils.compress_files(doc_path, join(self.tmp_dir, f'{basename(doc_path)}.zip'))
            return join(self.tmp_dir, f'{basename(doc_path)}.zip')
        return doc_path

    def _make_massage_doc(self, message: str, name: str = 'report.txt') -> str:
        doc_path = join(self.tmp_dir, name)
        with open(doc_path, "w") as file:
            file.write(message)
        return doc_path

    @property
    def _access(self) -> bool:
        if self._telegram_token and self._chat_id:
            return True
        return False
