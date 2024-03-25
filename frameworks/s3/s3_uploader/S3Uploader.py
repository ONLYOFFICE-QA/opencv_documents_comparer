# -*- coding: utf-8 -*-
from os.path import basename, join, isfile
from typing import Optional

from rich import print
import concurrent.futures

from host_tools import File

from s3wrapper import S3Wrapper


class S3Uploader:
    _checked_dirs = {}

    def __init__(
            self,
            bucket_name='conversion-testing-files',
            region='us-east-1',
            check_duplicates: bool = True,
            cache_results_dir: str | None = None
    ):
        self.cache_results_dir = cache_results_dir
        self.cache_file = join(self.cache_results_dir, 'cached_s3_files_sha256.json') if cache_results_dir else ...
        self.check_duplicates = check_duplicates
        self.s3 = S3Wrapper(bucket_name=bucket_name, region=region)
        self.all_s3_files = self._get_all_files()

    def upload_from_dir(self, dir_path: str) -> None:
        for file_path in File.get_paths(dir_path, exceptions_files=['.DS_Store']):
            s3_dir = file_path.split('.')[-1].lower()
            object_key = f"{s3_dir}/{basename(file_path)}"
            file_sha256 = File.get_sha256(file_path)

            if object_key in self.all_s3_files:
                print(f'[red] File already exists in: {self.s3.bucket}/{object_key}')
                continue

            if self.check_duplicates:
                s3_sha256_files = self._get_s3_files_sha256(s3_dir, cores=200)
                if s3_sha256_files.get(file_sha256):
                    print(f"[red] File {basename(file_path)} has the same hash as\n{s3_sha256_files[file_sha256]}")
                    continue

            self.s3.upload(file_path, object_key)

    def _get_all_files(self) -> list:
        print(f"[green]|INFO| Getting information about files from s3 bucket: {self.s3.bucket}")
        return self.s3.get_files()

    def _get_cache_results(self) -> Optional[dict]:
        if isfile(self.cache_file):
            return File.read_json(self.cache_file)
        print(f"[red]|ERROR| Cache file not exists: {self.cache_file}")

    def cache_results(self, cores: int = None):
        File.write_json(self.cache_file, self.get_object_info(cores=cores))

    def _get_s3_files_sha256(self, s3_dir: str, cores: int = None) -> dict:
        if s3_dir in self._checked_dirs:
            return self._checked_dirs[s3_dir]

        print(f"[green]|INFO| Getting s3 files sha256 from {s3_dir} dir")
        result = self.get_objects_sha256([file for file in self.all_s3_files if file.startswith(s3_dir)], cores)
        self._checked_dirs[s3_dir] = result
        return result

    def get_object_info(self, cores: int = None) -> dict:
        results = {}
        def process_object(object_key):
            results[object_key]['sha256'] = self.s3.get_sha256(object_key)

        with concurrent.futures.ThreadPoolExecutor(max_workers=cores) as executor:
            futures = [executor.submit(process_object, object_key) for object_key in self.all_s3_files]
            self.process_concurrent_result(futures)

        return results

    @staticmethod
    def process_concurrent_result(futures: list):
        for future in concurrent.futures.as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f'Exception occurred: {e}')

    def get_objects_sha256(self, object_keys: list, cores: int = None) -> dict:
        sha256_results = {}
        def process_object(object_key):
            sha256 = self.s3.get_sha256(object_key)
            if sha256 in sha256_results:
                sha256_results[sha256].append(object_key)
            else:
                sha256_results[sha256] = [object_key]

        with concurrent.futures.ThreadPoolExecutor(max_workers=cores) as executor:
            self.process_concurrent_result([executor.submit(process_object, object_key) for object_key in object_keys])

        return sha256_results
