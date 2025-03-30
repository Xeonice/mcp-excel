from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="mcp-excel",
    version="0.1.1",
    author="Eric Julianto",
    author_email="",
    description="MCP server to give client the ability to read Excel files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ericjulianto/mcp-excel",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.12 :: Only",
    ],
    python_requires=">=3.12",
    install_requires=[
        "mcp[cli]>=1.3.0",
        "openpyxl>=3.1.5",
        "pandas>=2.2.3",
        "py>=1.11.0",
    ],
    entry_points={
        "console_scripts": [
            "mcp-excel=mcp_excel.main:main",
        ],
    },
) 