from setuptools import setup, find_packages

setup(
    name="onesharepy",
    version="0.0.2",
    author="xiongzhen (公众号: 正月十九)",
    description="Download shared personal onedrive files.",
    packages=find_packages(),
    python_requires=">=3.13",
    install_requires=['aiohttp', 'requests'],
)
