from pprint import pprint

from tests.conversion_tests import TestConfig



config = TestConfig(
    input_dir='1',
    cores=1,
    result_dir='1',
    core_dir='12',
    reports_dir='1234',
    errors_only=True,
    version="8.2.2.22",
    tmp_dir='234234',
    fonts_dir='234',
timestamp=True
)



print(config.fonts_dir)
