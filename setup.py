from setuptools import setup, find_packages

setup(
    name='setuporders',
    description='bot to automate work',

    version='1.0',
    author='Urszula Maciaszek',
    autohor_email='maciaszek89@gmail.com',
    
    packages=find_packages(where='src'),
    package_dir={'':'src'},
    )
