# -*- coding: utf-8 -*-

# A very simple setup script to create a single executable
#
# hello.py is a very simple 'Hello, world' type script which also displays the
# environment in which the script runs
#
# Run the build process by running the command 'python setup.py build'
#
# If everything works well you should find a subdirectory in the build
# subdirectory that contains the files needed to run the script without Python

from cx_Freeze import setup, Executable

executables = [
    Executable('ZteDumpsToCsv.py')
]

buildOptions = dict(create_shared_zip=False,include_msvcr=True)

setup(name='ZteDumpsToCsv',
      version='0.1',
      description='cx_Freeze script',
      options=dict(build_exe=buildOptions),
      executables=executables
      )
