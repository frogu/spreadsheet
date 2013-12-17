from setuptools import setup, find_packages
import sys, os

version = '1.1.1'
def read(fname):
    try:
        return open(os.path.join(os.path.dirname(__file__), fname)).read()
    except:
        return ""

setup(name='spreadsheet',
      setup_requires=["setuptools_hg"],
      version=version,
      description="Universal Spreadsheet data access",
      long_description=read('README.txt'),
      classifiers=[
                   "Development Status :: 3 - Alpha",
                   'Environment :: Console',
                   'Natural Language :: English',
                   'Operating System :: Microsoft :: Windows :: Windows NT/2000',
                   'Operating System :: POSIX :: Linux',
                   'Programming Language :: Python :: 2.7',
                   'Topic :: Software Development :: Libraries :: Python Modules',
                   ],  # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='spreadsheet csv xls xlsx',
      author='\xc5\x81ukasz Proszek',
      author_email='proszek@gmail.com',
      packages=find_packages('src', exclude=['ez_setup', 'examples', 'tests']),
      package_dir={'': 'src'},
      package_data={
                      '': ['*.txt', ],
                      },
      include_package_data=True,
      zip_safe=True,
      install_requires=[
                        'openpyxl>=1.1.7',
                        'xlrd>=0.7.1',
                        'xlwt>=0.7.2',
      ],

      )
