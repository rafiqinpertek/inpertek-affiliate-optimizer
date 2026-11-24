"""PostgreSQL storage for the INSERT pilot. Secrets never logged or returned to UI.

The original engine uses SQLite for its proven Excel parser and diff algorithm. We
load only the latest PostgreSQL run into an isolated temporary SQLite database,
calculate the next run, and atomically persist the new run to PostgreSQL.
No master Excel bytes are stored in PostgreSQL.
"""
import os
import sqlite3
import tempfile

import pandas as pd
import psycopg

from engine import init as init_sqlite, sync as sync_sqlite

SCHEMA_PG = '''
CREATE SCHEMA IF NOT EXISTS inpertek_internal;
CREATE TABLE IF NOT EXISTS sync_runs (
  run_id BIGINT PRIMARY KEY, file_hash TEXT NOT NULL, source_name TEXT,
  sync_time TEXT NOT NULL, sheets_green INTEGER, source_rows INTEGER,
  listing_rows INTEGER, issues INTEGER, notes TEXT
);
CREATE TABLE IF NOT EXISTS records (
  run_id BIGINT NOT NULL, source_sheet TEXT NOT NULL, source_row INTEGER NOT NULL,
  sku TEXT, name TEXT, marketplace TEXT NOT NULL, price DOUBLE PRECISION,
  affiliate_nominal DOUBLE PRECISION, commission_pct DOUBLE PRECISION,
  margin_after DOUBLE PRECISION, margin_pct DOUBLE PRECISION, hpp DOUBLE PRECISION,
  row_hash TEXT, validation TEXT,
  PRIMARY KEY (run_id,source_sheet,source_row,marketplace)
);
CREATE INDEX IF NOT EXISTS idx_records_key ON records(sku,marketplace);
CREATE TABLE IF NOT EXISTS changes (
  run_id BIGINT, marketplace TEXT, source_sheet TEXT, sku TEXT, field TEXT,
  previous_value TEXT, new_value TEXT, event_type TEXT
);
CREATE TABLE IF NOT EXISTS issues (
  run_id BIGINT, source_sheet TEXT, source_row INTEGER, marketplace TEXT,
  sku TEXT, issue TEXT
);
CREATE TABLE IF NOT EXISTS mappings (
  source_sheet TEXT PRIMARY KEY, mapping_json TEXT,
  validated INTEGER, updated_at TEXT
);
'''

TABLE_COLUMNS = {
 'sync_runs': ['run_id','file_hash','source_name','sync_time','sheets_green','source_rows','listing_rows','issues','notes'],
 'records': ['run_id','source_sheet','source_row','sku','name','marketplace','price','affiliate_nominal','commission_pct','margin_after','margin_pct','hpp','row_hash','validation'],
 'changes': ['run_id','marketplace','source_sheet','sku','field','previous_value','new_value','event_type'],
 'issues': ['run_id','source_sheet','source_row','marketplace','sku','issue'],
 'mappings': ['source_sheet','mapping_json','validated','updated_at'],
}


def _connect(url):
    if not url or not url.startswith(('postgresql://','postgres://')):
        raise ValueError('DATABASE_URL PostgreSQL belum dikonfigurasi.')
    return psycopg.connect(url, connect_timeout=12, sslmode='require', options='-c search_path=inpertek_internal')


def init_cloud(url):
    with _connect(url) as pg:
        with pg.cursor() as cur:
            cur.execute(SCHEMA_PG, prepare=False)


def read_query(url, sql, params=()):
    """Read-only SQL query. Only application-authored SELECTs should be passed."""
    if not sql.lstrip().upper().startswith('SELECT'):
        raise ValueError('Hanya query SELECT yang diizinkan untuk dashboard.')
    with _connect(url) as pg:
        with pg.cursor() as cur:
            cur.execute(sql.replace('?', '%s'), params)
            names=[desc.name for desc in cur.description]
            return pd.DataFrame(cur.fetchall(),columns=names)


def _latest_to_sqlite(pg, tempdb, last_run):
    sq=init_sqlite(tempdb)
    if last_run is None:
        sq.close()
        return
    for tab in ('sync_runs','records'):
        cols=TABLE_COLUMNS[tab]
        with pg.cursor() as cur:
            cur.execute(f'SELECT {",".join(cols)} FROM {tab} WHERE run_id=%s',(last_run,))
            while rows:=cur.fetchmany(500):
                sq.executemany(f'INSERT INTO {tab} ({",".join(cols)}) VALUES ({",".join("?" for _ in cols)})', rows)
    sq.commit()
    sq.close()


def _upload_new_run(pg, sq, source_run_id, pg_run_id):
    # Explicit column allowlists (no arbitrary SQL identifiers supplied by users).
    for tab in ('sync_runs','records','changes','issues'):
        cols=TABLE_COLUMNS[tab]
        rows=sq.execute(f'SELECT {",".join(cols)} FROM {tab} WHERE run_id=?',(source_run_id,)).fetchall()
        if not rows: continue
        with pg.cursor() as cur:
            cmd=f'INSERT INTO {tab} ({",".join(cols)}) VALUES ({",".join("%s" for _ in cols)})'
            chunk=300
            for off in range(0,len(rows),chunk):
                prepared=[]
                for r in rows[off:off+chunk]:
                    r=list(r);r[0]=pg_run_id;prepared.append(tuple(r))
                cur.executemany(cmd,prepared)
    for row in sq.execute('SELECT source_sheet,mapping_json,validated,updated_at FROM mappings'):
        pg.execute('''INSERT INTO mappings (source_sheet,mapping_json,validated,updated_at)
                      VALUES (%s,%s,%s,%s)
                      ON CONFLICT(source_sheet) DO UPDATE SET
                      mapping_json=EXCLUDED.mapping_json,validated=EXCLUDED.validated,
                      updated_at=EXCLUDED.updated_at''',row)


def sync_cloud(path,url):
    """Atomic import under a PostgreSQL advisory transaction lock.

    Avoid collisions between simultaneous users by serializing imports and
    assigning the run_id from the last committed run.
    """
    with _connect(url) as pg:
        with pg.transaction():
            pg.execute('SELECT pg_advisory_xact_lock(%s)',(61002411,))
            last=pg.execute('SELECT run_id FROM sync_runs ORDER BY run_id DESC LIMIT 1').fetchone()
            last_id=int(last[0]) if last else None
            with tempfile.TemporaryDirectory(prefix='inpertek_sync_') as tmp:
                local=os.path.join(tmp,'latest.sqlite')
                _latest_to_sqlite(pg,local,last_id)
                result=sync_sqlite(path,local)
                if result['status']=='UNCHANGED':
                    return dict(result,storage='PostgreSQL')
                new_id=(last_id or 0)+1
                sq=sqlite3.connect(local)
                try:
                    _upload_new_run(pg,sq,result['run_id'],new_id)
                finally:
                    sq.close()
                return dict(result,run_id=new_id,storage='PostgreSQL')
