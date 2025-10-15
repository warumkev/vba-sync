from setuptools import setup, find_packages

setup(
    name='vba-sync',
    version='0.1.0',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'click',
        'olefile',
    ],
    entry_points={
        'console_scripts': [
            'vba-sync = vba_sync.main:cli',
        ],
    },
)