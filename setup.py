from pathlib import Path

from setuptools import find_packages, setup

version = "0.0.1"

with open('README.md', encoding='utf-8') as readme_file:
    readme = readme_file.read()

setup(
    name='excel-comparison-bot',
    version=version,
    description=(
        'A telegram bot that compares two excel files '
        'and returns difference.'
    ),
    long_description=readme,
    long_description_content_type='text/markdown',
    author='Maksym Ivanchenko',
    author_email='github@ivanchenko.io',
    url='https://github.com/vanchaxy/excel-comparison-bot',
    packages=find_packages(include=['excel_comparison_bot', 'excel_comparison_bot.*']),
    entry_points={'console_scripts': ['excel-comparison-bot = excel_comparison_bot.bot:main']},
    include_package_data=True,
    python_requires='<3.10',
    install_requires=[
        line.strip()
        for line in Path('requirements.txt').read_text().split('\n')
    ],
    zip_safe=False,
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Natural Language :: English",
        "Programming Language :: Python :: 3 :: Only",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python",
    ],
)
