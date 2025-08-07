from setuptools import setup, find_packages

setup(
    name="vpo",
    version="0.1",
    packages=find_packages(where="module"),
    package_dir={"": "module"},
)