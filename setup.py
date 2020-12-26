from setuptools import setup, find_packages

setup(
    name='metadata_names',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'python-magic',
        'metadata_extractor @ git+ssh://git@github.com/vphpersson/metadata_extractor.git#egg=metadata_extractor',
        'terminal_utils @ git+ssh://git@github.com/vphpersson/terminal_utils.git#egg=terminal_utils',
    ]
)
