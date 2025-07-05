# -*- coding: utf-8 -*-
from os.path import join, dirname, isfile, getsize
from typing import Optional

from host_tools import File
from host_tools.utils import Dir
from rich.console import Console

from s3wrapper import S3Wrapper
import concurrent.futures

class S3Downloader:
    """
    A class for downloading objects from Amazon S3.
    This class allows downloading objects from an Amazon S3 bucket with support for parallel downloading.

    Attributes:
        download_dir (str): The directory where files will be downloaded.
        cores (int): The number of CPU cores to use for parallel downloading.
        bucket_name (str): The name of the S3 bucket.
        region (str): The AWS region of the S3 bucket.
        access_key (str): The AWS access key.
        secret_access_key (str): The AWS secret access key.
        check_sha256 (bool): Whether to check SHA256 hash of downloaded files.
        check_size (bool): Whether to check file size of downloaded files.
    """
    def __init__(
            self,
            download_dir: str,
            cores: Optional[int] = None,
            bucket_name: str = 'conversion-testing-files',
            region: str = 'us-east-1',
            access_key: Optional[str] = None,
            secret_access_key: Optional[str] = None,
            check_sha256: bool = False,
            check_size: bool = True
    ):
        """
        Initializes an S3Downloader object.
        :param download_dir: The directory where files will be downloaded.
        :param cores: The number of CPU cores to use for parallel downloading. (default: None)
        :param bucket_name: The name of the S3 bucket. (default: 'conversion-testing-files')
        :param region: The AWS region of the S3 bucket. (default: 'us-east-1')
        :param access_key: The AWS access key. (default: None)
        :param secret_access_key: The AWS secret access key. (default: None)
        :param check_sha256: Whether to check SHA256 hash of downloaded files. (default: False)
        :param check_size: Whether to check file size of downloaded files. (default: True)
        """
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
        """
        Downloads a single object from S3.
        :param obj_key: The key of the object in the S3 bucket.
        """
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

    def download_all(self, objects: Optional[list] = None):
        """
        Downloads multiple objects from S3.
        :param objects: List of object keys to download. (default: None)
        """
        File.delete(self.errors_log_file, stdout=False, stderr=False)
        object_keys = objects if isinstance(objects, list) else self._get_all_files()

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
        """
        Gets the result of a thread execution.
        :param future: The future object representing the result of a thread.
        :return: The result of the thread execution.
        """
        try:
            return future.result()
        except (PermissionError, FileExistsError, NotADirectoryError, IsADirectoryError) as e:
            return f"[red]|ERROR| Exception when getting result {e}"

    def _check_download_errors(self):
        """
        Checks for any download errors and prints the result.
        """
        errors = File.read(self.errors_log_file) if isfile(self.errors_log_file) else None
        self.console.print(
            f"[red]|ERROR| Download errors: {errors}" if errors
            else f"[green]|INFO| All objects have been successfully downloaded."
        )

    def _get_all_files(self):
        """
        Retrieves a list of all objects from the S3 bucket.
        :return: List of object keys in the S3 bucket.
        """
        self.console.print(f"[green]|INFO| Getting a list of files from the bucket: [cyan]{self.s3.bucket}")
        return self.s3.get_files()

    def _exists_object(self, download_path: str, obj_key: str, check_sha256: bool) -> bool:
        """
        Checks if the downloaded object exists and optionally verifies its integrity.
        :param download_path: The path where the object is downloaded.
        :param obj_key: The key of the object in the S3 bucket.
        :param check_sha256: Whether to check the SHA256 hash of the downloaded object.
        :return: True if the object exists and passes integrity checks, False otherwise.
        """
        if not isfile(download_path):
            return False
        if self.check_size and getsize(download_path) != self.s3.get_size(obj_key):
            return False
        if check_sha256 and File.get_sha256(download_path) != self.s3.get_sha256(obj_key):
            return False
        return True
