# -*- coding: utf-8 -*-
import json
import os
import requests
from tempfile import gettempdir
from os.path import join, getsize, basename, isdir, isfile, expanduser
from rich import print

from frameworks.decorators import singleton
from host_control import FileUtils


@singleton
class Telegram:
    def __init__(
            self,
            token_path: str = None,
            chat_id_path: str = None,
            tmp_dir: str = gettempdir()
    ):

        self._telegram_dir = join(expanduser('~'), '.telegram')
        self._telegram_token = self._get_token(token_path)
        self._chat_id = self._get_chat_id(chat_id_path)
        self.tmp_dir = tmp_dir
        FileUtils.create_dir(self.tmp_dir, stdout=False)

    def send_message(self, message: str, out_msg=False) -> None:
        print(message) if out_msg else ...
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
        self._request(
            f"https://api.telegram.org/bot{self._telegram_token}/sendDocument",
            data={"chat_id": self._chat_id, "caption": self._prepare_caption(caption), "parse_mode": "Markdown"},
            files={"document": open(self._prepare_documents(document_path), 'rb')}
        )

    def send_media_group(self, document_paths: list, caption: str = None, media_type: str = 'document') -> None:
        """
        :param document_paths:
        :param caption:
        :param media_type: types: 'photo', 'video', 'audio', 'document', 'voice', 'animation'
        :return:
        """
        if not document_paths:
            return self.send_message(caption if caption else 'No files to send.', out_msg=True)
        if caption and len(caption) > 200:
            document_paths.append(self._make_massage_doc(caption, 'caption.txt'))
        files, media = {}, []
        for doc_path in document_paths:
            files[basename(doc_path)] = open(self._prepare_documents(doc_path), 'rb')
            media.append(dict(type=media_type, media=f'attach://{basename(doc_path)}'))
        media[-1]['caption'] = self._prepare_caption(caption) if caption is not None else ''
        media[-1]['parse_mode'] = "Markdown"
        self._request(
            f'https://api.telegram.org/bot{self._telegram_token}/sendMediaGroup',
            data={'chat_id': self._chat_id, 'media': json.dumps(media)},
            files=files
        )

    @staticmethod
    def _prepare_caption(caption: str) -> str:
        return caption[:200]

    def _request(self, url: str, data: dict, files: dict = None, tg_log: bool = True) -> None:
        if self._telegram_token and self._chat_id:
            try:
                response = requests.post(url, data=data, files=files)
                if response.status_code != 200:
                    print(f"Error when sending to telegram: {response.status_code}")
                    self.send_message(f"Error when sending to telegram: {response.status_code}") if tg_log else ...
            except Exception as e:
                print(f"|WARNING|Impossible to send: {data}. Error: {e}")
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
