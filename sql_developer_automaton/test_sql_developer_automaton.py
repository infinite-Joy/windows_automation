import time

from lfe_task import load_config_file, \
                        open_app, \
                        close_tip_of_the_day, \
                        open_conections, \
                        run_queries, \
                        save_query_output, \
                        handle_export_window, \
                        get_query, \
                        map_plus_typekey, \
                        emailling, \
                        close_app
import lfe_task
from nose.tools import nottest

# helper functions
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

def present_active_window(window_name):
    app = application.Application()
    handles = findwindows.find_windows()
    for w_handle in handles:
        wind = app.window_(handle=w_handle)
        if window_name == wind.Texts():
            return True

def skipIf(skip_condition=True, msg=None):
    ''' Make function raise SkipTest exception if skip_condition is true

    Parameters
    ----------
    skip_condition : bool or callable.
        Flag to determine whether to skip test.  If the condition is a
        callable, it is used at runtime to dynamically make the decision.  This
        is useful for tests that may require costly imports, to delay the cost
        until the test suite is actually executed.
    msg : string
        Message to give on raising a SkipTest exception

   Returns
   -------
   decorator : function
       Decorator, which, when applied to a function, causes SkipTest
       to be raised when the skip_condition was True, and the function
       to be called normally otherwise.

    Notes
    -----
    You will see from the code that we had to further decorate the
    decorator with the nose.tools.make_decorator function in order to
    transmit function name, and various other metadata.
    '''
    if msg is None:
        msg = 'Test skipped due to test condition'
    def wrapper(f):
        # Local import to avoid a hard nose dependency and only incur the 
        # import time overhead at actual test-time. 
        import nose
        def skipper(*args, **kwargs):
            if skip_condition:
                raise nose.SkipTest, msg
            else:
                return f(*args, **kwargs)
        return nose.tools.make_decorator(f)(skipper)
    return wrapper

def runWhen(run_condition=True):
    ''' Make function raise SkipTest exception if run_condition is true

    Parameters
    ----------
    run_condition : bool or callable.
        Flag to determine whether to skip test.  If the condition is a
        callable, it is used at runtime to dynamically make the decision.  This
        is useful for tests that may require costly imports, to delay the cost
        until the test suite is actually executed.
    msg : string
        Message to give on raising a SkipTest exception

   Returns
   -------
   decorator : function
       Decorator, which, when applied to a function, causes SkipTest
       to be raised when the run_condition was True, and the function
       to be called normally otherwise.

    Notes
    -----
    You will see from the code that we had to further decorate the
    decorator with the nose.tools.make_decorator function in order to
    transmit function name, and various other metadata.
    '''
    msg = 'Test skipped due to test condition not met'
    def wrapper(f):
        # Local import to avoid a hard nose dependency and only incur the 
        # import time overhead at actual test-time. 
        import nose
        def runnner(*args, **kwargs):
            try:
                if run_condition:
                    return f(*args, **kwargs)
            except:
                raise nose.SkipTest, msg
        return nose.tools.make_decorator(f)(runnner)
    return wrapper

## end of helper functions

## tests

def test_load_config_file():
    res = load_config_file()
    print(res)
    assert res == {u'task1': {u'path': u'path/to.sql/developer', u'query': u"sleect query"}}

#@skipIf(skip_condition=present_active_window([u'Oracle SQL Developer']))#
@nottest
def test_open_app():
    open_app()
    time.sleep(10)
    assert present_active_window([u'Oracle SQL Developer']) == 2

#@runWhen(run_condition=present_active_window([u'Tip of the Day']))
@nottest
def test_close_tip_of_the_day():
    close_tip_of_the_day()
    time.sleep(2)
    assert present_active_window([u'Tip of the Day']) == True

@nottest
def test_open_conections():
    open_conections()
    assert True

@nottest
def test_run_queries():
    try:
        run_queries()
        assert True
    except Exception, e:
        print(e)
        assert False

@nottest
def test_save_query_output():
    save_query_output()
    assert True

@nottest
def test_handle_export_window():
    try:
        handle_export_window()
        assert True
    except Exception, e:
        print(e)
        assert False

#@nottest 
def test_get_query():
    res = get_query()
    print(res)
    assert res == [select query list]

def test_map_plus_typekey1():
    res = map_plus_typekey(["(+)"])
    assert res == ['{(}', '{+}', '{)}']


def test_map_plus_typekey2():
    res = map_plus_typekey(['(+)', '{SPACE}', '{SPACE}', 'desc;'])
    print(res)
    assert res == 1

@nottest
def test_emailling():
    emailling()
    assert True

@skipIf(skip_condition=present_active_window([u'give window name']))
def test_close_app():
    close_app()
    assert True
