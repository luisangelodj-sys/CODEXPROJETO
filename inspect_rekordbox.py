import sqlite3
import sys


db = sys.argv[1]
con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
for (name,) in con.execute("select name from sqlite_master where type='table' order by name"):
    print(name)
