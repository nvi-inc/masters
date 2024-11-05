from setuptools import setup

setup(
    name='masters',
    packages=['masters'],
    description='Master Schedule generator',
    version='2.5.0',
    url='https://github.com/nvi-inc/masters.git',
    author='Mario',
    author_email='mario.berube@nviinc.com',
    keywords=['vlbi', 'master', 'ivs'],
    install_requires=['toml', 'paramiko', 'openpyxl', 'python-docx',
                      'pywin32 ; platform_system=="Windows"', 'appscript ; sys_platform=="darwin"'],
    include_package_data=False,
    package_data={'': ['data/fs-10.toml', 'data/types.json']},
    entry_points={
        'console_scripts': [
            'make_master=masters.make_master:main',
            'backup=masters.backup:main',
            'reqsched=masters.reqsched:main',
            'make_notes=masters.notes:main',
            'make_xlsx=masters.make_xlsx:main'
        ]
    },
)
