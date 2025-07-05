# -*- coding: utf-8 -*-
from os.path import basename, splitext
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
        self.cores = int(cores) if cores else None
        self.check_duplicates = check_duplicates
        self.s3 = S3Wrapper(bucket_name=bucket_name, region=region)
        self.all_s3_files = self._fetch_all_files()

    def generate_unique_object_key(self, file_name: str, s3_dir: str, all_s3_files_lower: list) -> str:
        """
        Generate a unique object key for a file to be uploaded to S3.

        :param file_path: The local file path to upload.
        :param s3_dir: The S3 directory where the file will be uploaded.
        :param all_s3_files: List of all files currently in S3.
        :return: A unique object key for the file.
        """
        base, ext = splitext(file_name)
        new_object_key = f"{s3_dir}/{base}{ext}"
        counter = 1
        while new_object_key.lower() in all_s3_files_lower:
            new_object_key = f"{s3_dir}/{base}_{counter}{ext}"
            counter += 1
        return new_object_key

    def upload_file(self, file_path: str) -> bool:
        """
        Upload a single file to the S3 bucket.

        :param file_path: The local file path to upload.
        :return: True if the file was successfully uploaded, False otherwise.
        """
        # Extract file name and directory
        file_name = basename(file_path)
        s3_dir = splitext(file_name)[1][1:].lower()
        object_key = f"{s3_dir}/{file_name}"
        file_sha256 = File.get_sha256(file_path)
        all_s3_files_lower = [file_name.lower() for file_name in self.all_s3_files]

        # Check if file already exists in S3
        if object_key.lower() in all_s3_files_lower:
            file_in_s3 = self.get_s3_file_path(object_key)
            s3_object_sha256 = self.s3.get_sha256(file_in_s3)

            if file_sha256 == s3_object_sha256:
                self._print_file_exists(file_name, file_in_s3)
                return False

            self._print_file_conflict(file_name, file_in_s3, file_sha256, s3_object_sha256)
            object_key = self.generate_unique_object_key(file_name, s3_dir, all_s3_files_lower)
            print(f'[yellow]File {file_name} uploaded as: [magenta]{object_key}[/]')


        # Check for duplicate hashes in S3
        if self.check_duplicates:
            s3_sha256_files = self._fetch_s3_files_sha256(s3_dir)
            if file_sha256 in s3_sha256_files:
                self._print_duplicate_hash(file_name, s3_sha256_files[file_sha256])
                return False

        # Upload the file
        self.s3.upload(file_path, object_key)
        return True

    def _print_file_exists(self, file_name: str, file_in_s3: str) -> None:
        """
        Print a message indicating that the file already exists in S3.

        :param file_name: The name of the file.
        :param file_in_s3: The path of the file in S3.
        """
        print(f'[cyan]File [magenta]{file_name}[/] already exists in: [magenta]{self.s3.bucket}/{file_in_s3}[/]')

    def _print_file_conflict(self, file_name: str, file_in_s3: str, local_sha256: str, s3_sha256: str) -> None:
        """
        Print a message indicating a file conflict due to SHA256 mismatch.

        :param file_name: The name of the file.
        :param file_in_s3: The path of the file in S3.
        :param local_sha256: The SHA256 of the local file.
        :param s3_sha256: The SHA256 of the file in S3.
        """
        print(
            f'[bold red]File {file_name} conflict in [magenta]{self.s3.bucket}/{file_in_s3}[/]\n'
            f'SHA256 mismatch:\n{file_name} Local: [cyan]{local_sha256}[/]\n'
            f'{file_in_s3} S3: [magenta]{s3_sha256}[/]'
        )

    def _print_duplicate_hash(self, file_name: str, s3_file: str) -> None:
        """
        Print a message indicating a duplicate hash in S3.

        :param file_name: The name of the file.
        :param s3_file: The path of the file in S3 with the same hash.
        """
        print(
            f"[bold red] File [magenta]{file_name}[/] has the same hash as an existing file: "
            f"[magenta]{s3_file}[/] in S3."
        )

    def get_s3_file_path(self, object_key: str) -> str:
        """
        Get the path of a file in the S3 bucket with the same case as the local file.

        :param object_key: The object key of the file to get the path for.
        :return: The path of the file in the S3 bucket with the same case as the local file.
        """
        for s3_file in self.all_s3_files:
            if s3_file.lower() == object_key.lower():
                return s3_file
        return None

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

        print(f"[green]|INFO| Fetching SHA256 hashes from S3 directory: [magenta]{s3_dir}[/]")
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
