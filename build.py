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
    '--add-data=crawlers/crawl_catalog_i.py;crawlers',
    '--add-data=crawlers/crawl_catalog_ii.py;crawlers',
    '--add-data=crawlers/crawl_cosing.py;crawlers',
    '--add-data=src/app_qt.py;src',
    '--add-data=src/database.py;src',
    '--add-data=src/catalog_mgr.py;src',
    '--add-data=src/safety_engine.py;src',
    '--add-data=src/report_generator.py;src',
    '--add-data=src/product_classifier.py;src',
    '--add-data=src/project_manager.py;src',
    '--add-data=src/market_config.py;src',
    '--add-data=src/review_workflow.py;src',
    '--add-data=userdata/cosing_data.json;userdata',
    '--add-data=userdata/ingredient_catalog.json;userdata',
    '--add-data=userdata/ingredient_catalog_ii.json;userdata',
    '--hidden-import=src.database',
    '--hidden-import=src.catalog_mgr',
    '--hidden-import=src.app_qt',
    '--hidden-import=src.safety_engine',
    '--hidden-import=src.report_generator',
    '--hidden-import=src.product_classifier',
    '--hidden-import=src.project_manager',
    '--hidden-import=src.market_config',
    '--hidden-import=src.review_workflow',
    '--hidden-import=crawlers.crawl_catalog_i',
    '--hidden-import=crawlers.crawl_catalog_ii',
    '--hidden-import=crawlers.crawl_cosing',
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
