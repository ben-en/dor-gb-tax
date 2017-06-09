from setuptools import setup

import dor_gb_tax as mod

setup(
    name='dor-gb-tax',
    version='0.1',
    author='Ben Ennis',
    author_email='pnw.ben.ennis@gmail.com',
    description='A script for use with greenbits to create tax reports',
    packages=[mod.__name__],
    entry_points={
        'console_scripts': [
            'dor-gb-tax = dor_gb_tax.main:main',
        ],
    },
)
