import sqlite3
sql = sqlite3.connect("./test.db")
sql.execute("""create table if not exists languages (
key text primary key,
urtext text,
zh_CN text)""")


def insert_data(key, value):
    sql.execute(
        "INSERT OR REPLACE into multilingual (key,urtext) values ('%s','%s')" % (key, value))
    sql.commit()
