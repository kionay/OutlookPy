import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="OutlookPy",
    version="0.1.0",
    author="Patrick Childers",
    author_email="patrickchildersit@gmail.com",
    description="A Python wrapper for Outlook's COM API",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/kionay/OutlookPy",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows :: Windows 7",
        "Operating System :: Microsoft :: Windows :: Windows 8",
        "Operating System :: Microsoft :: Windows :: Windows 8.1",
        "Operating System :: Microsoft :: Windows :: Windows 10",
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "Natural Language :: English",
        "Topic :: Software Development :: Libraries"
    ],
    install_requires=[
        "pypiwin32 == 223",
        "pywin32 == 224"
    ],
)