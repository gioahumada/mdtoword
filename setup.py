from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="mdtoword",
    version="0.1.0",
    author="Gio Ahumada",
    author_email="giovanni.ahumada.t@gmail.com",
    description="A Python package to convert Markdown files to Word documents with advanced formatting",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/gioahumada/mdtoword",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=[
        "python-docx",
        "markdown",
        "beautifulsoup4",
    ],
)
