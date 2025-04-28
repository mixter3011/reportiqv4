from setuptools import setup, find_packages

setup(
    name="portfolio_review",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "PyQt5>=5.15.0",
        "pandas>=1.3.0",
        "numpy>=1.20.0",
        "selenium>=4.0.0",
        "openpyxl>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "portfolio_review=main:main",
        ],
    },
    author="Finance Team",
    description="Portfolio Review Application",
)