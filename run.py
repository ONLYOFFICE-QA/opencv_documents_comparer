from tqdm import tqdm
from multiprocessing import Process

from libs.get_error import run_get_pr, ggg
from libs.word import WordCompareImg, Word
from var import *

if __name__ == "__main__":
    for i in tqdm(range(1)):
        p1 = Process(target=run_get_pr)
        p1.start()
        WordCompareImg()
        p1.terminate()

