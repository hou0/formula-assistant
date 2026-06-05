import PyInstaller.__main__
import os
import shutil

base_dir = os.path.dirname(os.path.abspath(__file__))

for d in ['build', 'dist']:
    p = os.path.join(base_dir, d)
    if os.path.exists(p):
        shutil.rmtree(p, ignore_errors=True)
spec_path = os.path.join(base_dir, '备案配方表助手.spec')
if os.path.exists(spec_path):
    try:
        os.remove(spec_path)
    except PermissionError:
        pass

PyInstaller.__main__.run([
    'main.py',
    '--name=备案配方表助手',
    '--onedir',
    '--windowed',
    '--add-data=crawl_catalog_i.py;.',
    '--add-data=crawl_catalog_ii.py;.',
    '--add-data=userdata/cosing_data.json;userdata',
    '--add-data=userdata/ingredient_catalog.json;userdata',
    '--add-data=userdata/ingredient_catalog_ii.json;userdata',
    '--hidden-import=database',
    '--hidden-import=catalog_mgr',
    '--hidden-import=crawl_catalog_i',
    '--hidden-import=crawl_catalog_ii',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=requests',
    '--hidden-import=selenium',
    '--hidden-import=selenium.webdriver.edge.options',
    '--hidden-import=selenium.webdriver.edge.webdriver',
    '--hidden-import=selenium.webdriver.common.by',
    '--hidden-import=selenium.webdriver.support.ui',
    '--hidden-import=selenium.webdriver.support.expected_conditions',
    '--hidden-import=bs4',
    '--hidden-import=PySide6.QtNetwork',
    '--exclude-module=matplotlib',
    '--exclude-module=scipy',
    '--exclude-module=torch',
    '--exclude-module=torchvision',
    '--exclude-module=tensorflow',
    '--exclude-module=PIL',
    '--exclude-module=sympy',
    '--exclude-module=pydub',
    '--exclude-module=wandb',
    '--exclude-module=sentry_sdk',
    '--exclude-module=uvicorn',
    '--exclude-module=jsonschema',
])

print("\n打包完成！EXE 位于: dist/备案配方表助手/备案配方表助手.exe")
