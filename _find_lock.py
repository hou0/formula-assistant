import psutil, os
fp = os.path.join('data', '安全评估数据库.xlsx')
for proc in psutil.process_iter(['pid','name']):
    try:
        for f in proc.open_files():
            if '安全评估数据库' in f.path:
                print(f"PID:{proc.info['pid']}, Name:{proc.info['name']}, File:{f.path}")
    except:
        pass
