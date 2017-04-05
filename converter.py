#!/usr/bin/python3

import pyinotify
import subprocess
import os
import logging
from logging.handlers import SysLogHandler

XLS_DIRECTORY = "/input-files/excel"
SUFFIXES = {".xlsx", ".xls"}
LOG_LEVEL = logging.INFO


class EventHandler(pyinotify.ProcessEvent):
    def __init__(self, suffixes, directory, logger, **kargs):
        super().__init__(**kargs)
        self.suffixes = suffixes
        self.directory = directory
        self.logger = logger

    def process_IN_CREATE(self, event):
        if self.should_process(event):
            self.logger.info("Converting %s" % event.pathname)
            converted = self.convert_file(event)
            if converted:
                self.logger.info(
                    "Converted {in_file} to {out_file}".format(
                        in_file=event.pathname,
                        out_file=self.output_file_path(event)
                    )
                )
                self.remove_source_file(event)
            else:
                self.logger.error("Error converting %s" % event.pathname)

    def should_process(self, event):
        return os.path.splitext(event.name)[1] in self.suffixes

    def convert_file(self, event):
        return_code = subprocess.call(['xlsx2csv', event.pathname, self.output_file_path(event)])
        return return_code == 0

    def output_file_path(self, event):
        return "{dir}/out/{file_name}.csv".format(dir=self.directory, file_name=event.name)

    def remove_source_file(self, event):
        os.remove(event.pathname)


log = logging.getLogger(__name__)
log.setLevel(LOG_LEVEL)
log_handler = SysLogHandler(address="/dev/log")
log_handler.setFormatter(logging.Formatter('%(module)s.%(funcName)s: %(message)s'))
log.addHandler(log_handler)

handler = EventHandler(suffixes=SUFFIXES, directory=XLS_DIRECTORY, logger=log)
watch_manager = pyinotify.WatchManager()
notifier = pyinotify.Notifier(watch_manager, handler)
watch_descriptors = watch_manager.add_watch(XLS_DIRECTORY, mask=pyinotify.IN_CREATE, rec=False)

for path, result in watch_descriptors.items():
    if result > 0:
        log.info("Added watch for path %s" % path)
    else:
        log.error("Error adding watch for path %s" % path)

notifier.loop()
