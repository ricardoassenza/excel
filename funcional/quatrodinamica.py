import win32com.client


def atualiza_dinamicas(caminho):

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(caminho)

    for sheet in wb.Sheets:
        for pivot in sheet.PivotTables():
            pivot.RefreshTable()

    wb.Save()
    wb.Close()
    excel.Quit()

    print('dinamica atualizada')