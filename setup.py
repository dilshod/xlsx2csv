
import os
import shutil
from distutils.core import setup

if not os.path.exists('scripts'):
    os.makedirs('scripts')
shutil.copyfile('xlsx2csv.py', 'scripts/xlsx2csv')

scripts = ["scripts/xlsx2csv"]

name = "xlsx2csv"
version = "0.7.2"
author = "Dilshod Temirkhdojaev"
author_email = "tdilshod@gmail.com"
desc = "xlsx to csv converter"
long_desc = "CherryPy is a pythonic, object-oriented HTTP framework"
url = "http://github.com/dilshod/xlsx2csv"
classifiers=[
    "Development Status :: 5 - Production/Stable",
    "Environment :: Console",
    "Intended Audience :: End Users/Desktop",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: GNU General Public License (GPL)",
    "Operating System :: OS Independent",
    "Programming Language :: Python",
    "Programming Language :: Python :: 2",
    "Programming Language :: Python :: 2.4",
    "Programming Language :: Python :: 2.5",
    "Programming Language :: Python :: 2.6",
    "Programming Language :: Python :: 2.7",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.0",
    "Programming Language :: Python :: 3.1",
    "Programming Language :: Python :: 3.2",
    "Programming Language :: Python :: 3.3",
    "Programming Language :: Python :: 3.4",
    "Topic :: Office/Business",
    "Topic :: Utilities"
]
data_files=[
    ('test', []),
]

setup(
    name='xlsx2csv',
    version='0.7.2',
    description=desc,
    author=author,
    author_email=author_email,
    classifiers=classifiers,
    py_modules=['xlsx2csv'],
    data_files=data_files,
    url=url,
    scripts=scripts
)
