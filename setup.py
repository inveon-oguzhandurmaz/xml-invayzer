from setuptools import setup, find_packages
import sys

# Python sürümü kontrolü
if sys.version_info < (3, 7):
    sys.exit('Python 3.7 veya daha yüksek bir sürüm gereklidir.')

setup(
    name="xml-invayzer",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'tkinter',
        'requests',
        'pandas',
        'openpyxl',
        'aiohttp',
    ],
    entry_points={
        'console_scripts': [
            'xml-invayzer=xml_analyzer:main',
        ],
    },
    author="Oğuzhan Durmaz",
    author_email="oguzhan.durmaz@inveon.com",
    description="XML dosyalarını analiz eden ve URL'leri kontrol eden bir uygulama",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    keywords="xml analyzer url checker",
    url="https://github.com/inveon-oguzhandurmaz/xml-invayzer",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
) 