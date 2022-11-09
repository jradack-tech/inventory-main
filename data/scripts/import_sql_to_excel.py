import pandas as pd
from sqlalchemy import create_engine
import sys

hostname="localhost"
dbname="test"
uname="root"
pwd=""

# df = pd.read_excel("20210219-R3.2.19.xlsx", "customer master")
#
# engine = create_engine("mysql+pymysql://{user}:{pw}@{host}/{db}".format(host=hostname, db=dbname, user=uname, pw=pwd))
# df.to_sql('customer', engine, index=False)


df = pd.read_excel("20210219-R3.2.19.xlsx", "product name")

engine = create_engine("mysql+pymysql://{user}:{pw}@{host}/{db}".format(host=hostname, db=dbname, user=uname, pw=pwd))
df.to_sql('product', engine, index=False)
