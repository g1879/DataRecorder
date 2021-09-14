# -*- coding:utf-8 -*-
from setuptools import setup, find_packages

with open("README.md", "r", encoding='utf-8') as fh:
    long_description = fh.read()

setup(
    name="DataRecorder",
    version="1.1.0",
    author="g1879",
    author_email="g1879@qq.com",
    description="用于记录数据的模块。",
    long_description=long_description,
    long_description_content_type="text/markdown",
    license="MIT",
    keywords="DataRecorder",
    url="https://gitee.com/g1879/DataRecorder",
    include_package_data=True,
    packages=find_packages(),
    install_requires=[
        "openpyxl"
    ],
    classifiers=[
        "Programming Language :: Python :: 3.6",
        "Development Status :: 4 - Beta",
        "Topic :: Utilities",
        "License :: OSI Approved :: BSD License",
    ],
    python_requires='>=3.6'
)
