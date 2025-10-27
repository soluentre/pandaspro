from setuptools import setup, find_packages

setup(
    name='pandaspro',
    version='0.10.5',
    description='Upgraded pandas package for easier dataframe operations',
    packages=find_packages(),
    install_requires=[
        'numpy>=1.26.0',
        'pandas>=2.2.1',
        'openpyxl>=3.1.5',
        'xlwings>=0.31.10',
        'jinja2>=3.1.4',
        'termcolor>=2.4.0',
        'tabulate>=0.9.0',
        'python-dateutil>=2.8.2',
        'PyPDF2>=3.0.1',
    ],
    python_requires='>=3.8',
    py_modules=[],
    author='Shiyao Wang and Xueqi Li',
    author_email='sw1016@georgetown.edu',
    url='https://github.com/soluentre/pandaspro',
    long_description='''
pandaspro is the upgrade version of pandas Python pacakge with more user-friendly dataframe manipulating methods and 
better Excel-based exporting and formatting apis 
''',
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ]
)