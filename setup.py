from setuptools import setup

setup(
    name='easypyxl',
    version='0.1.0',
    descripton='This python package is a wrapper of OpenPyXL for easy usage.',
    url='https://github.com/CauchyComplete/EasyPyXL',
    author='CauchyComplete',
    author_email='corundum240@gmail.com',
    license='MIT',
    packages=['easypyxl'],
    install_requires=[
        'openpyxl>=3.0.9'
    ],
)


