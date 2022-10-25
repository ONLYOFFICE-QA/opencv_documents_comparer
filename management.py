# -*- coding: utf-8 -*-
import pyperclip as pc
from invoke import task
import os
import json
import random
import shutil
import subprocess as sb
import sys
import codecs
import psutil
import pyautogui as pg
from loguru import logger
from rich import print
from rich.progress import track
from multiprocessing import Process
from time import sleep
import win32con
import win32gui
from win32com.client import Dispatch
import csv
import io
