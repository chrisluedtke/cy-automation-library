import datetime
import os
from configparser import ConfigParser
from pathlib import Path

import pysftp

# this script appends a "YYYY.MM.DD" subdirectory to this path
OUTDIR = "Z:/ChiPrivate/Chicago Data and Evaluation/SY20/CPS SFTP"

def get_creds(key_name, creds_path = 'credentials.ini'):
    config = ConfigParser()

    try:
        open(creds_path)
    except FileNotFoundError:
        raise FileNotFoundError(
            'Before you can use this library, you must create '
            'a credentials.ini file. See the README for '
            'details.'
        )
    else:
        config.read_file(open(creds_path))

    HOST = config[key_name]['host']
    USER = config[key_name]['username']
    PASS = config[key_name]['password']
    
    return HOST, USER, PASS

def get_srv(key_name):
    HOST, USER, PASS = get_creds(key_name)
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None  # returns error without this line
    srv = pysftp.Connection(host=HOST, username=USER, 
                            password=PASS, cnopts=cnopts)
    return srv

def export_cps_to_cyc(read_dir="Outgoing/2019-2020", write_dir=OUTDIR):
    # create YYYY.MM.DD subdirectory
    dt = datetime.datetime.now()
    dt_dir = (str(dt.year) + '.' + str(dt.month).zfill(2) + '.' +
              str(dt.day).zfill(2))
    write_dir = os.path.join(write_dir, dt_dir)

    if not os.path.exists(write_dir):
        os.mkdir(write_dir)

    srv = get_srv(key_name='cps')
    srv.get_d(read_dir, write_dir, preserve_mtime=True)

    return None

if __name__=="__main__":
    export_cps_to_cyc()
