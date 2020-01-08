import datetime
import os
import tempfile
import zipfile
import shutil

import requests


def get_binary(url):
    r = requests.get(url, allow_redirects=True)
    assert r.status_code == 200, "Can't get a file from server"
    return r.content


def get_next_monday():
    today = datetime.date.today()
    next_monday = today + datetime.timedelta(days=-today.weekday(), weeks=1)
    return next_monday


def fix_xlsx(in_file):
    zin = zipfile.ZipFile(in_file, "r")
    if "xl/SharedStrings.xml" in zin.namelist():
        tmpfd, tmp = tempfile.mkstemp(dir=os.path.dirname(in_file))
        os.close(tmpfd)

        with zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                if item.filename == "xl/SharedStrings.xml":
                    zout.writestr("xl/sharedStrings.xml", zin.read(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))

        zin.close()
        shutil.move(tmp, in_file)
