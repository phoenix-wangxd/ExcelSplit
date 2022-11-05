#!/usr/bin/env python

from setuptools import setup, find_packages

mode_descript = "If there are a large number of data records in Excel, " \
                "this tool can easily split the data into different sheets"

setup(
    name='excel_split_large_records',
    version='0.0.1',
    description=mode_descript,
    author='Phoenix Wang',
    author_email='phoenix.wangxd@icloud.com',
    python_requires=">=3.8",
    packages=find_packages(
        include=['*'],
    ),
    install_requires=[
        'requests',
        'openpyxl',
    ],
)
