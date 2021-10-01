from setuptools import setup


with open("README.md", "r") as f:
    long_description = f.read()

setup(
    name='easypyxl',
    version='0.4.0',
    descripton='This python package is a wrapper of OpenPyXL for easy usage.',
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/CauchyComplete/EasyPyXL',
    author='CauchyComplete',
    author_email='corundum240@gmail.com',
    license='MIT',
    packages=['easypyxl'],
    install_requires=[
        'openpyxl>=3.0.9'
    ],
)


