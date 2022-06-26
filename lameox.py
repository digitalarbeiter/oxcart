#!/usr/bin/env python3
import colorful as cf
import datetime
import json
import os

import oxcart

if __name__ == "__main__":
    ox = oxcart.OX(
        "https://office.mailbox.org/ajax",
        os.environ.get("OX_USERNAME"),
        os.environ.get("OX_PASSWORD"),
        debug=True,
    )
    ox.debug = False
    for type_ in ["calendar", "contacts", "mail", "tasks"]:
        all_folders = ox.folders.all_folders(type_)
        print(cf.cyan(f"all {type_}s:"))
        for folder in all_folders:
            print(cf.green(folder["title"]), json.dumps(folder, indent=4, sort_keys=True))
    ox.logout()
