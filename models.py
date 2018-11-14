from pony.orm import *

db = Database()

class Stud(db.Entity):
    seq = Required(str)
    name = Required(str)
    addr = Required(str)

# set_sql_debug(True)
db.bind(provider='sqlite', filename='my.db', create_db=True)
db.generate_mapping(create_tables=True)
