# -*- coding: utf-8 -*-
from os.path import join, dirname, isfile, getsize

from host_tools import File
from host_tools.utils import Dir
from rich.console import Console

from S3Wrapper import S3Wrapper
import concurrent.futures

class S3Downloader:
    def __init__(
            self,
            download_dir: str,
            cores: int = None,
            bucket_name: str = 'conversion-testing-files',
            region: str = 'us-east-1',
            access_key: str = None,
            secret_access_key: str = None,
            check_sha256: bool = False,
            check_size: bool = True
    ):
        self.check_size = check_size
        self.check_sha256 = check_sha256
        self.download_dir = download_dir
        self.errors_log_file = join(self.download_dir, "Failed_download_log.txt")
        self.cores = cores
        self.console = Console()
        self.s3 = S3Wrapper(
            bucket_name=bucket_name,
            region=region,
            access_key=access_key,
            secret_access_key=secret_access_key
        )

    @property
    def cores(self):
        return self.__cores

    @cores.setter
    def cores(self, value):
        try:
            self.__cores = int(value)
        except TypeError:
            self.console.print(f"[red]|ERROR|The value cores mast be integer") if not value is None else ...
            self.__cores = None

    def download_obj(self, obj_key: str):
        download_path = join(self.download_dir, obj_key)

        if self._exists_object(download_path, obj_key, self.check_sha256):
            return f"[cyan]|INFO| Object {obj_key} already exists"

        Dir.create(dirname(download_path), stdout=False)
        self.s3.download(obj_key, download_path, stdout=False)

        if self._exists_object(download_path, obj_key, check_sha256=False):
            self.console.print(f"[green]|INFO| Object [cyan]{obj_key} [green]downloaded")
        else:
            File.write(self.errors_log_file, f"{obj_key}\n", mode='a')
            self.console.print(f"[red]|ERROR| Object [cyan]{obj_key} [red]not downloaded")

    def download_all(self, objects: list = None):
        File.delete(self.errors_log_file, stdout=False, stderr=False)
        object_keys = objects if isinstance(objects, list) else self._get_all_objects()

        with self.console.status('') as status:
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.cores) as executor:
                self.console.print(
                    f"[green]|INFO| Start downloading "
                    f"[red]{len(object_keys)}[/] objects in [red]{executor._max_workers}[/] threads"
                )

                futures = [executor.submit(self.download_obj, obj_key) for obj_key in object_keys]
                for future in concurrent.futures.as_completed(futures):
                    future.add_done_callback(lambda *_: status.update(self._get_thread_result(future)))

            concurrent.futures.wait(futures)
        self._check_download_errors()

    @staticmethod
    def _get_thread_result(future):
        try:
            return future.result()
        except (PermissionError, FileExistsError, NotADirectoryError, IsADirectoryError) as e:
            return f"[red]|ERROR| Exception when getting result {e}"

    def _check_download_errors(self):
        errors = File.read(self.errors_log_file) if isfile(self.errors_log_file) else None
        self.console.print(
            f"[red]|ERROR| Download errors: {errors}" if errors
            else f"[green]|INFO| All objects have been successfully downloaded."
        )

    def _get_all_objects(self):
        self.console.print(f"[green]|INFO| Getting a list of objects from the bucket: [cyan]{self.s3.bucket}")
        return self.s3.get_files()

    def _exists_object(self, download_path: str, obj_key: str, check_sha256: bool) -> bool:
        if not isfile(download_path):
            return False
        if self.check_size and getsize(download_path) != self.s3.get_size(obj_key):
            return False
        if check_sha256 and File.get_sha256(download_path) != self.s3.get_sha256(obj_key):
            return False
        return True
