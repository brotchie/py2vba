def pytest_funcarg__xl(request):
    from win32com.client import Dispatch
    xl = Dispatch('Excel.Application')

    def finalize():
        xl.DisplayAlerts = 0
#        xl.visible = 1
        xl.Quit()

    request.addfinalizer(finalize)
    return xl

def pytest_funcarg__workbook(request):
    xl = request.getfuncargvalue('xl')
    wb = xl.Workbooks.Add()

    def finalize():
        wb.Close(False)
        pass

    request.addfinalizer(finalize)
    return wb
