import sqlite3, os
import pandas as pd

BASE=os.path.dirname(os.path.abspath(__file__))
DB=os.path.join(BASE,'news_monitor.db')
OUT=os.path.join(BASE,'news_monitor_export.xlsx')
con=sqlite3.connect(DB)
df=pd.read_sql_query('SELECT * FROM articles ORDER BY published_time DESC, id DESC',con)
con.close()
if df.empty:
    print('数据库暂无新闻。')
else:
    df.to_excel(OUT,index=False)
    print(f'已导出 {len(df)} 条新闻：{OUT}')
