# -*- coding:utf-8 -*-

from setuptools import setup
from setuptools import find_packages

VERSION = '1.1.0'

setup(
    name='pythonexcel',
    description='',
    long_description='',
    classifiers=[],
    keywords='',
    author='Lawes',
    author_email='haiou_chen@sina.cn',
    url='https://github.com/MrLawes/pythonexcel',
    license='MIT',
    packages=find_packages(),
    include_package_data=True,
    zip_safe=True,
    install_requires=[
        'xlsxwriter==0.9.6',
    ],
    version=VERSION,
)