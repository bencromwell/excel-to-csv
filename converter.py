#!/usr/bin/python3

import pyinotify
import subprocess
import os

XLS_DIRECTORY = "/input-files/excel"

SUFFIXES = {".xlsx", ".xls"}

wm = pyinotify.WatchManager()
mask = pyinotify.IN_CREATE

def should_process(event):
    return os.path.splitext(event.name)[1] in SUFFIXES

class EventHandler(pyinotify.ProcessEvent):
    def process_IN_CREATE(self, event):
        if should_process(event):
            print ("Converting:", event.pathname)
            subprocess.call(['xlsx2csv', event.pathname, XLS_DIRECTORY + "/out/" + event.name + ".csv"])
            os.remove(event.pathname)

handler = EventHandler()
notifier = pyinotify.Notifier(wm, handler)
wdd = wm.add_watch(XLS_DIRECTORY, mask, rec=True)

notifier.loop()

