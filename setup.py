from setuptools import setup

setup(name='microsoftgraph',
      version='0.1',
      description='API wrapper for Microsoft Graph written in Python',
      url='https://github.com/GearPlug/microsoftgraph-python',
      author='Miguel Ferrer, Nerio Rincon',
      author_email='ingferrermiguel@gmail.com',
      license='GPL',
      packages=['microsoftgraph'],
      install_requires=[
          'requests',
      ],
      zip_safe=False)
