from setuptools import setup

from xlrd.info import __VERSION__

setup(
    name='xlrd',
    version=__VERSION__,
    author='John Machin',
    author_email='sjmachin@lexicon.net',
    url='http://www.python-excel.org/',
    packages=['xlrd'],
    scripts=[
        'scripts/runxlrd.py',
    ],
    description=(
        'Library for developers to extract data from '
        'Microsoft Excel (tm) spreadsheet files'
    ),
    long_description=(
        "Extract data from Excel spreadsheets "
        "(.xls and .xlsx, versions 2.0 onwards) on any platform. "
        "Pure Python (3.7, 3.8). "
        "Strong support for Excel dates. Unicode-aware."
    ),
    platforms=["Any platform -- don't need Windows"],
    license='BSD',
    keywords=['xls', 'excel', 'spreadsheet', 'workbook'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Operating System :: OS Independent',
        'Topic :: Database',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires=">=3.7.0",
)
