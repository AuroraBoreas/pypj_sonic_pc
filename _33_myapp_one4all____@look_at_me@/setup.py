from setuptools import setup, find_namespace_packages

with open("README.md") as f:
    long_description = f.readlines()

setup(
    name='ssv_pc_office_auto_package',

    version='0.0.4',

    description='',
    long_description=long_description,
    long_description_content_type='text/markdown',

    author='ZL',
    author_email='liang.zhang@sony.com',

    license='Apach2',

    packages=find_namespace_packages(include=['ssv_pc_office_auto_pkg']),
    zip_safe=False,
)