from setuptools import setup, find_packages

setup(
    name="CompareResults",
    version="0.1.0",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "Flask>=2.3.2",
        "pandas>=1.5.3",
        "openpyxl>=3.1.2",
        "python-pptx>=0.6.21",
    ],
    entry_points={
        "console_scripts": [
            "compare-results=compare_results.core:main",
        ],
    },
    author="Your Name",
    author_email="your.email@example.com",
    description="A web app to compare Excel results and generate analysis summaries in PowerPoint.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    # url="https://github.com/yourusername/CompareResults",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)

