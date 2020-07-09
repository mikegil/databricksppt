#!/usr/bin/env python

"""The setup script."""

from setuptools import setup, find_packages

with open('README.rst') as readme_file:
    readme = readme_file.read()

with open('HISTORY.rst') as history_file:
    history = history_file.read()

requirements = ['Click>=7.0', ]

setup_requirements = []

test_requirements = []

setup(
    author="Michael R. Gilbert",
    author_email='michael.r.gilbert@me.com',
    python_requires='>=3.6',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
    ],
    description="Python package to allow creation of Think-Cell PPT charts from Databricks",
    entry_points={
        'console_scripts': [
            'databricksppt=databricksppt.cli:main',
        ],
    },
    install_requires=requirements,
    license="MIT license",
    long_description=readme + '\n\n' + history,
    include_package_data=True,
    keywords='databricksppt',
    name='databricksppt',
    packages=find_packages(include=['databricksppt', 'databricksppt.*']),
    package_data={'databricksppt': ['data/template.html']},
    scripts=['databricksppt/bin/databricksppt_script.py'],
    setup_requires=setup_requirements,
    test_suite='tests',
    tests_require=test_requirements,
    url='https://github.com/mikegil/databricksppt',
    version='0.1.0-dev12',
    zip_safe=False,
)
