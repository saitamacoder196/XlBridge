from setuptools import setup, find_packages

setup(
    name="xlbridge",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "openpyxl>=3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "xlbridge=xlbridge.cli:main",
        ],
    },
    python_requires=">=3.8",
)
