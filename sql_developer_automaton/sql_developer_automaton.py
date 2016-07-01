from __future__ import print_function
import time
from random import randint
import logging
import datetime
import json

from email_module import send_mail

__NAME__ = "sql_developer_automaton"
__author__ = "joydeep bhattacharjee"
__version__ = "1.0"

# import the required libraries
try:
    from pywinauto import application
except ImportError:
    import os.path
    pywinauto_path = os.path.abspath(__file__)
    pywinauto_path = os.path.split(os.path.split(pywinauto_path)[0])[0]
    import sys
    sys.path.append(pywinauto_path)
    from pywinauto import application

from pywinauto import findwindows
from pywinauto.controls.HwndWrapper import HwndWrapper
from pywinauto import WindowAmbiguousError

#vars

config_file = "config.json"

def get_date():
    mylist = []
    today = datetime.date.today()
    mylist.append(today)
    return mylist[0]


logFormatter = logging.Formatter("%(asctime)s [%(threadName)-12.12s] [%(levelname)-5.5s]  %(message)s")
rootLogger = logging.getLogger()

fileName = "automaton_%s_%s" % (get_date(), randint(0, 50))
fileHandler = logging.FileHandler("{0}.log".format(fileName))
fileHandler.setFormatter(logFormatter)
rootLogger.addHandler(fileHandler)

consoleHandler = logging.StreamHandler()
consoleHandler.setFormatter(logFormatter)
rootLogger.addHandler(consoleHandler)

def present_active_windows():
    global rootLogger
    handles = findwindows.find_windows()
    for w_handle in handles:
        wind = app.window_(handle=w_handle)
        rootLogger.info(wind.Texts())

def load_config_file():
    with open(config_file) as config_fh:
        config = json.load(config_fh)
        return config

def map_space_typekey(query):
    query = query.split(" ")
    runnable_query = []
    for indx, item  in enumerate(query):
        runnable_query.append(item)
        if indx < len(query) - 1:
            runnable_query.append("{SPACE}")
    return runnable_query

def handle_special_chars(item, char):
    return list(item.split(char)[0]) + \
        ["{%s}" % char] + list(item.split(char)[1])

def handle__rem_special_chars(item):
    if item == "(":
        return "{(}"
    elif item == ")":
        return "{)}"
    else:
        return item

def map_plus_typekey(elem):
    runnable_query = []
    for indx, item  in enumerate(elem):
        if "(+)" in item:
            map(runnable_query.append, handle_special_chars(item, "+"))
        elif "(" in item:
            map(runnable_query.append, handle_special_chars(item, "("))
        elif ")" in item:
            map(runnable_query.append, handle_special_chars(item, ")"))
        else:
            runnable_query.append(item)

    return map(handle__rem_special_chars, runnable_query)


def get_query():
    query = load_config_file()["task1"]["query"]
    query = str(query)
    query = map_space_typekey(query)
    return map_plus_typekey(query)

def open_app():
    global rootLogger
    config = load_config_file()
    app = application.Application().start(r"%s" % config["task1"]["path"])
    rootLogger.info("sql developer starting ..")
    time.sleep(25)

def close_tip_of_the_day():
    global rootLogger
    app = application.Application().Connect(title=u'Tip of the Day', class_name='SunAwtDialog')
    rootLogger.info("will close the Tip of the day window")
    sunawtdialog = app.SunAwtDialog
    sunawtdialog.Close()
    rootLogger.info("Tip of the day window closed")


def open_conections():
    global rootLogger
    app = application.Application().Connect(title_re=u'Oracle SQL Developer.*', class_name='oracle.ideimpl.MainWindowImpl')
    oracleideimplmainwindowimpl = app[u'Oracle SQL Developer']
    rootLogger.info("will open application")
    oracleideimplmainwindowimpl.TypeKeys("%v")
    rootLogger.info("menu item view opened")
    time.sleep(1)
    oracleideimplmainwindowimpl.ClickInput(coords=(114, 40))
    rootLogger.info("menu item Connections clicked")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{RIGHT}")
    rootLogger.info("Connections drop down opened.")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{RIGHT}")
    time.sleep(2)
    oracleideimplmainwindowimpl.TypeKeys("{DOWN}")
    oracleideimplmainwindowimpl.TypeKeys("{DOWN}")
    oracleideimplmainwindowimpl.TypeKeys("{DOWN}")
    rootLogger.info("went to required connection")
    oracleideimplmainwindowimpl.TypeKeys("{RIGHT}")
    rootLogger.info("opening connection")
    time.sleep(5)
    oracleideimplmainwindowimpl.Wait('ready')
    oracleideimplmainwindowimpl.ClickInput(coords=(641, 362))
    rootLogger.info("click on the developer console area.")
    time.sleep(1)

def run_queries():
    global rootLogger
    app = application.Application().Connect(title_re=u'Oracle SQL Developer.*', class_name='oracle.ideimpl.MainWindowImpl')
    oracleideimplmainwindowimpl = app[u'oracle.ideimpl.MainWindowImpl']
    oracleideimplmainwindowimpl.ClickInput()
    map(oracleideimplmainwindowimpl.TypeKeys, get_query())
    rootLogger.info("query has been typed out")
    time.sleep(1)
    oracleideimplmainwindowimpl.RightClickInput()
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{ENTER}")
    rootLogger.info("query has been executed")
    time.sleep(10)

def save_query_output():
    global rootLogger
    app = application.Application().Connect(title_re=u'Oracle SQL Developer.*', class_name='oracle.ideimpl.MainWindowImpl')
    oracleideimplmainwindowimpl = app[u'Oracle SQL Developer']
    oracleideimplmainwindowimpl.ClickInput(coords=(553, 800))
    rootLogger.info("clicked the results page")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("+{F10}")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{UP}")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{RIGHT}")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{UP}")
    oracleideimplmainwindowimpl.TypeKeys("{UP}")
    time.sleep(1)
    oracleideimplmainwindowimpl.TypeKeys("{ENTER}")
    rootLogger.info("export and save as xls")
    time.sleep(5)

def handle_export_window():
    global rootLogger
    app = application.Application().Connect(title=u'Export Data', class_name='SunAwtDialog')
    sunawtdialog = app[u'Export Data']
    rootLogger.info("linked export data window")
    sunawtdialog.ClickInput(coords=(423, 450))
    rootLogger.info("click yes")
    time.sleep(3)

def emailling():
    global rootLogger
    rootLogger.info("sending email...")
    text="attached the report..."
    subject="give the subject of the email"
    files=[r"path/to/saved/file"]
    send_mail(send_from="give from email", send_to=["person1@domain.com", "person2@domain.com"],text = text, subject = subject, files = files)
    rootLogger.info("email sent")


def close_app():
    app = application.Application().Connect(title_re=u'Oracle SQL Developer.*', class_name='oracle.ideimpl.MainWindowImpl')
    oracleideimplmainwindowimpl = app[u'Oracle SQL Developer']
    rootLogger.info("will be closing the app now...")
    oracleideimplmainwindowimpl.TypeKeys("%f")
    time.sleep(1)
    oracleideimplmainwindowimpl.ClickInput(coords=(106, 319))
    time.sleep(1)

    app = application.Application().Connect(title=u'Save Files', class_name='SunAwtDialog')
    sunawtdialog = app[u'Save Files']
    rootLogger.info("not saving any files")
    sunawtdialog.Click()
    sunawtdialog.TypeKeys("{RIGHT}")
    sunawtdialog.TypeKeys("{ENTER}")
    rootLogger.info("Work seems to be done. Please check your email.")
    rootLogger.info("App Closed.")



if __name__ == '__main__':
    open_app()
    close_tip_of_the_day()
    open_conections()
    run_queries()
    save_query_output()
    handle_export_window()
    emailling()
    close_app()
