import os
import sys
from setuptools import setup, find_packages

# Read requirements from file
def read_requirements():
    with open('requirements.txt', 'r') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="pdf-travel-expense-extractor",
    version="1.0.0",
    packages=find_packages(),
    install_requires=read_requirements(),
    entry_points={
        'console_scripts': [
            'pdf-extractor=main:main',
        ],
    },
    author="Assistant",
    author_email="assistant@example.com",
    description="Portable PDF travel expense extractor with OCR support",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/example/pdf-extractor",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
