import sqlite3
from pathlib import Path
db=Path(__file__).with_name("inventory.db")
con=sqlite3.connect(db)
con.executescript("""
CREATE TABLE IF NOT EXISTS host(host_id INTEGER PRIMARY KEY,host_name TEXT,host_type TEXT,manufacturer TEXT,model TEXT,datacenter TEXT,status TEXT);
CREATE TABLE IF NOT EXISTS application(application_id INTEGER PRIMARY KEY,application_name TEXT,environment TEXT,availability TEXT,status TEXT);
CREATE TABLE IF NOT EXISTS service_endpoint(endpoint_id INTEGER PRIMARY KEY,application_id INTEGER,endpoint_name TEXT,protocol TEXT,port INTEGER,vip_address TEXT);
INSERT INTO host VALUES(1,'DCG-ESX-01','ESXi Host','Dell','PowerEdge R760','DC-01','Active');
INSERT INTO host VALUES(2,'DCG-ESX-02','ESXi Host','Dell','PowerEdge R760','DC-01','Active');
INSERT INTO application VALUES(1,'Vyuha SQLite Patch Manager','Demo','HA','Active');
INSERT INTO application VALUES(2,'Vyuha SQLite Viewer','Demo','HA','Active');
INSERT INTO service_endpoint VALUES(1,1,'Patch Manager Web','HTTPS',443,'10.10.10.50');
""")
con.commit();con.close();print(f"Created {db}")
