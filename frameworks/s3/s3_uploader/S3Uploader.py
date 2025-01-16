# -*- coding: utf-8 -*-
from os.path import basename
from rich import print
import concurrent.futures

from host_tools import File
from s3wrapper import S3Wrapper


class S3Uploader:
    _exceptions: list = ['.DS_Store']
    _checked_dirs: dict = {}

    def __init__(
            self,
            bucket_name: str = 'conversion-testing-files',
            region: str = 'us-east-1',
            check_duplicates: bool = True,
            cores: int = None
    ):
        """
        Initialize the S3Uploader instance.

        :param bucket_name: The name of the S3 bucket to upload files to.
        :param region: The AWS region where the S3 bucket is located.
        :param check_duplicates: A flag to determine whether to check for duplicate files.
        :param cores: The number of cores to use for concurrent uploads.
        """
        self.cores = int(cores)
        self.check_duplicates = check_duplicates
        self.s3 = S3Wrapper(bucket_name=bucket_name, region=region)
        self.all_s3_files = self._fetch_all_files()

    def upload_file(self, file_path: str) -> bool:
        """
        Upload a single file to the S3 bucket.

        :param file_path: The local file path to upload.
        :return: True if the file was successfully uploaded, False otherwise.
        """
        s3_dir = file_path.split('.')[-1].lower()
        object_key = f"{s3_dir}/{basename(file_path)}"
        file_sha256 = File.get_sha256(file_path)

        if object_key in self.all_s3_files:
            s3_object_sha256 = self.s3.get_sha256(object_key)

            if file_sha256 == s3_object_sha256:
                print(
                    f'[cyan] File [magenta]{basename(file_path)}[/] already exists in: '
                    f'[magenta]{self.s3.bucket}/{object_key}[/]'
                )
            else:
                print(
                    f'[bold red] File conflict in [magenta]{self.s3.bucket}/{object_key}[/]\n'
                    f'SHA256 mismatch:\nLocal: [cyan]{file_sha256}[/]\nS3: [magenta]{s3_object_sha256}[/]'
                )
            return False

        if self.check_duplicates:
            s3_sha256_files = self._fetch_s3_files_sha256(s3_dir)
            if file_sha256 in s3_sha256_files:
                print(
                    f"[bold red] File [magenta]{basename(file_path)}[/] has the same hash as an existing file: "
                    f"[magenta]{s3_sha256_files[file_sha256]}[/] in S3."
                )
                return False

        self.s3.upload(file_path, object_key)
        return True

    def upload_from_dir(self, dir_path: str) -> None:
        """
        Upload all files from a given directory to the S3 bucket.

        :param dir_path: The directory path to upload files from.
        """
        for file_path in File.get_paths(dir_path, exceptions_files=self._exceptions):
            self.upload_file(file_path)

    def _fetch_all_files(self) -> list:
        """
        Fetch the list of all files currently stored in the S3 bucket.

        :return: A list of all files in the S3 bucket.
        """
        print(f"[green]|INFO| Fetching file list from S3 bucket: [magenta]{self.s3.bucket}[/]")
        return self.s3.get_files()

    def _fetch_s3_files_sha256(self, s3_dir: str) -> dict:
        """
        Fetch SHA256 hashes for files in a specific S3 directory.

        :param s3_dir: The directory in the S3 bucket to fetch SHA256 hashes from.
        :return: A dictionary mapping SHA256 hashes to lists of object keys.
        """
        if s3_dir in self._checked_dirs:
            return self._checked_dirs[s3_dir]

        print(f"[green]|INFO| Fetching SHA256 hashes from S3 directory: {s3_dir}")
        sha256_data = self._get_objects_sha256([file for file in self.all_s3_files if file.startswith(s3_dir)])
        self._checked_dirs[s3_dir] = sha256_data
        return sha256_data

    def _get_objects_sha256(self, object_keys: list) -> dict:
        """
        Fetch SHA256 hashes for multiple S3 objects.

        :param object_keys: A list of object keys to fetch SHA256 hashes for.
        :return: A dictionary mapping SHA256 hashes to lists of object keys.
        """
        sha256_results = {}

        def process_object(object_key: str) -> None:
            sha256 = self.s3.get_sha256(object_key)
            sha256_results.setdefault(sha256, []).append(object_key)

        with concurrent.futures.ThreadPoolExecutor(max_workers=self.cores) as executor:
            futures = [executor.submit(process_object, object_key) for object_key in object_keys]
            self._handle_futures(futures)

        return sha256_results

    @staticmethod
    def _handle_futures(futures: list) -> None:
        """
        Handle the completion of all future tasks in the ThreadPoolExecutor.

        :param futures: A list of futures representing asynchronous tasks.
        """
        for future in concurrent.futures.as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"[bold red] Exception occurred: {e}")
