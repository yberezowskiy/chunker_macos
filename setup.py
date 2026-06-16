from setuptools import setup

APP = ['chunker_ra_app.py']
DATA_FILES = ['cat_icon.icns']
OPTIONS = {
    'argv_emulation': True,
    'iconfile': 'cat_icon.icns',
    'packages': [
        'langchain_text_splitters',
        'docx',
        'pdfplumber',
        'reportlab',
        'langchain',
        'langchain_core'
    ],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
