import pywinauto
from pywinauto import application

app = application.Application().start(r"C:\Program Files\Calibre2\calibre.exe")

mainwin = app.window_(title_re = u'.*calibre.*')
mainwin.TypeKeys("{DOWN}")
mainwin.TypeKeys("{DOWN}")
mainwin.TypeKeys("{DOWN}")
mainwin.TypeKeys("v")