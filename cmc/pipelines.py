# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


# useful for handling different item types with a single interface
from itemadapter import ItemAdapter
import sqlite3


class CmcPipeline:
    # def process_item(self, item, spider):
    #     return item

    def __init__(self):
        self.con =sqlite3.connect('cmc.db')
        self.cur = self.con.cursor()
        self.cur.execute(""" 
        CREATE TABLE IF NOT EXISTS cmc(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pais TEXT,
            variable TEXT, 
            fecha DATE, 
            valor NUMERIC
        )
        """)
    def process_item(self, item, spider):
        # check if exixt a item and updated 
        self.cur.execute("SELECT * FROM cmc WHERE pais =? AND variable = ? AND fecha = ?", (item['pais'], item['variable'], item['fecha']))
        result = self.cur.fetchone()

        if result:
            self.cur.execute("SELECT id FROM cmc WHERE pais =? AND variable = ? AND fecha = ? AND valor=? ", (item['pais'], item['variable'], item['fecha'], item['valor']))
            id = self.cur.fetchone()
            self.cur.execute('''UPDATE cmc SET valor = ? WHERE id= ?''', (item['valor'], id[0]))
        
        else:
            self.cur.execute("""
                INSERT INTO cmc (pais, variable, fecha, valor) VALUES (?,?,?,?)
            """,
            (
                item['pais'],
                item['variable'],
                item['fecha'],
                item['valor']
            ))
        
        self.con.commit()
        return item


        
