import os
import sys
from setuptools import setup, find_packages
from fnmatch import fnmatchcase
from distutils.util import convert_path

standard_exclude = ('*.pyc', '*~', '.*', '*.bak', '*.swp*')
standard_exclude_directories = ('.*', 'CVS', '_darcs', './build', './dist', 'EGG-INFO', '*.egg-info')
def find_package_data(where='.', package='', exclude=standard_exclude, exclude_directories=standard_exclude_directories):
    out = {}
    stack = [(convert_path(where), '', package)]
    while stack:
        where, prefix, package = stack.pop(0)
        for name in os.listdir(where):
            fn = os.path.join(where, name)
            if os.path.isdir(fn):
                bad_name = False
                for pattern in exclude_directories:
                    if (fnmatchcase(name, pattern)
                        or fn.lower() == pattern.lower()):
                        bad_name = True
                        break
                if bad_name:
                    continue
                if os.path.isfile(os.path.join(fn, '__init__.py')):
                    if not package:
                        new_package = name
                    else:
                        new_package = package + '.' + name
                        stack.append((fn, '', new_package))
                else:
                    stack.append((fn, prefix + name + '/', package))
            else:
                bad_name = False
                for pattern in exclude:
                    if (fnmatchcase(name, pattern)
                        or fn.lower() == pattern.lower()):
                        bad_name = True
                        break
                if bad_name:
                    continue
                out.setdefault(package, []).append(prefix+name)
    return out

setup(name='docassemble.microsoftgraph',
      version='0.0.6',
      description=('Docassemble interface for Microsoft Graph API.'),
      long_description='# Microsoft Graph API (Office 365) for Docassemble\r\n\r\nProvides Graph API access to Docassemble, returning Docassemble objects. E.g.,\r\nIndividual, Address, etc.\r\n\r\nWorks with application level permissions, not user permissions.\r\n\r\n## Administrator Setup\r\n\r\nCreate an application in Azure Portal following instructions [here](https://docs.microsoft.com/en-us/graph/auth-v2-service?view=graph-rest-1.0)\r\nand provide it with a set of credentials.\r\n\r\nAdd a new section to your Docassemble configuration that looks like this, replacing\r\nwith the details from your new application:\r\n\r\n```\r\nmicrosoft graph:\r\n  tenant id: xxxxxxx\r\n  client id: xxxxxxxx\r\n  client secret: xxxxxxxxx\r\n```\r\n\r\n## Implemented APIs\r\n* Get user information (with user principal name, typically email address)\r\n* Get user contacts (with upn)\r\n\r\n## Usage\r\nSee example interview, `msgraph_example.yml`',
      long_description_content_type='text/markdown',
      author='Quinten Steenhuis',
      author_email='admin@admin.com',
      license='The MIT License (MIT)',
      url='https://docassemble.org',
      packages=find_packages(),
      namespace_packages=['docassemble'],
      install_requires=[],
      zip_safe=False,
      package_data=find_package_data(where='docassemble/microsoftgraph/', package='docassemble.microsoftgraph'),
     )

