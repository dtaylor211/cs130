'''
Context

context file for testing, provides correct path to sheets package

'''


import os
import sys


sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '..')))
