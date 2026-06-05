import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import database as db  # noqa: E402

conn = db.get_db()

print('=== 数据来源分布 ===')
rows = conn.execute('SELECT source, COUNT(*) as cnt FROM exposure_daily_usage GROUP BY source ORDER BY cnt DESC').fetchall()
for r in rows:
    print(f'  {r["source"]}: {r["cnt"]}')

print()
print('=== 儿童产品记录 ===')
rows = conn.execute("SELECT * FROM exposure_daily_usage WHERE application_site LIKE '%儿童%' OR product_category LIKE '%儿童%'").fetchall()
for r in rows:
    print(f'  {r["application_site"]} | {r["product_category"]} | {r["daily_amount_g"]}g | F={r["retention_factor"]}')

print()
print('=== 年龄段体重默认值 ===')
rows = conn.execute('SELECT * FROM exposure_body_weight ORDER BY default_weight_kg').fetchall()
for r in rows:
    print(f'  {r["age_group"]}: {r["default_weight_kg"]}kg ({r["source"]})')

print()
total = conn.execute('SELECT COUNT(*) as c FROM exposure_daily_usage').fetchone()['c']
print(f'日均使用量记录数: {total}')
total2 = conn.execute('SELECT COUNT(*) as c FROM exposure_body_weight').fetchone()['c']
print(f'年龄段体重记录数: {total2}')

conn.close()
