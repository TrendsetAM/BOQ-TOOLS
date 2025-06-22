from setuptools import setup, find_packages

setup(
    name="boq-tools",
    version="0.1.0",
    description="Bill of Quantities Excel Processor",
    author="Your Name",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "pathlib2; python_version < '3.4'"
    ],
    entry_points={
        "console_scripts": [
            "boq-tools=main:main"
        ]
    },
) 