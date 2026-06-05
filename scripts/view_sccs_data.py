import sqlite3

conn = sqlite3.connect('userdata/safety.db')
conn.row_factory = sqlite3.Row

rows = conn.execute('SELECT * FROM exposure_daily_usage WHERE source LIKE "%SCCS%" ORDER BY product_category').fetchall()

print('=== SCCS 数据详情 ===')
print(f'共 {len(rows)} 条记录\n')
print(f'{"产品类别":<20} {"使用部位":<15} {"日均使用量(g)":<15} {"驻留因子":<10}')
print('-' * 70)
for row in rows:
    print(f'{row["product_category"]:<20} {row["application_site"]:<15} {str(row["daily_amount_g"]):<15} {str(row["retention_factor"]):<10}')

conn.close()
