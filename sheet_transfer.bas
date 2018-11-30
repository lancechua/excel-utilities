Attribute VB_Name = "sheet_transfer"
Option Explicit
Option Base 0

Public Sub TransferSheets()
    ' Parameters will be read from 'INPUT' sheet
    Dim src_WB As Workbook
    Dim dest_WB As Workbook
    Dim src_WS As Worksheet, dest_WS As Worksheet, temp_WS As Worksheet
    Dim cur_src As String, cur_dest As String
    Dim cur_row As Integer, copy_idx As Integer
    Dim asu As Boolean, ada As Boolean, aee As Boolean, calc As Variant, aas As Variant
    
    Const sf_col = 1
    Const ss_col = 2
    Const df_col = 3
    Const ds_col = 4
    Const del_col = 5
    
    asu = Application.ScreenUpdating
    ada = Application.DisplayAlerts
    aee = Application.EnableEvents
    calc = Application.Calculation
    aas = Application.AutomationSecurity
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    
    
    cur_src = ""
    cur_dest = ""
    cur_row = 2
    
    
    With ThisWorkbook.Worksheets("INPUT")
        Do Until .Cells(cur_row, 1).Value2 = ""
            
            If cur_dest <> .Cells(cur_row, df_col).Value2 Then
                If Not dest_WB Is Nothing Then
                    dest_WB.Close savechanges:=True
                End If
                Application.StatusBar = "Opening " & .Cells(cur_row, df_col).Value2
                Set dest_WB = Workbooks.Open(.Cells(cur_row, df_col).Value2)
                cur_dest = .Cells(cur_row, df_col).Value2
            End If
            
            If cur_src <> .Cells(cur_row, sf_col).Value2 Then
                If Not src_WB Is Nothing Then
                    src_WB.Close savechanges:=False
                End If
                Application.StatusBar = "Opening " & .Cells(cur_row, sf_col).Value2
                Set src_WB = Workbooks.Open(.Cells(cur_row, sf_col).Value2)
                cur_src = .Cells(cur_row, sf_col).Value2
            End If

            If sheetExists(.Cells(cur_row, ss_col).Value2, src_WB) Then
                
                ' check if destination sheet to be deleted
                If .Cells(cur_row, del_col).Value2 Or LCase(Left(.Cells(cur_row, del_col).Value2, 1)) = "Y" Then
                    If sheetExists(.Cells(cur_row, ds_col).Value2, dest_WB) Then
                        copy_idx = dest_WB.Worksheets(.Cells(cur_row, ds_col).Value2).Index
                        Call deleteSheet(.Cells(cur_row, ds_col).Value2, dest_WB)
                    Else
                        copy_idx = dest_WB.Worksheets.Count + 1
                    End If
                    
                End If
                
                Set src_WS = src_WB.Worksheets(.Cells(cur_row, ss_col).Value2)
                ' remove filter
                If src_WS.AutoFilterMode Then
                    If src_WS.FilterMode Then src_WS.ShowAllData
                End If

                Application.StatusBar = "Checking if " & .Cells(cur_row, ds_col).Value2 & " exists in " & dest_WB.Name
                If sheetExists(.Cells(cur_row, ds_col).Value2, dest_WB) Then
                    Set dest_WS = dest_WB.Worksheets(.Cells(cur_row, ds_col).Value2)
                    
                    Application.StatusBar = "pasting over data in" & dest_WS.Name
                    
                    dest_WS.UsedRange.Clear
                    src_WS.UsedRange.Copy _
                        Destination:=dest_WS.Range(src_WS.UsedRange.Cells(1, 1).Address)
                Else
                    Application.StatusBar = "Inserting " & .Cells(cur_row, ds_col).Value2
                
                    src_WS.Copy after:=dest_WB.Worksheets(dest_WB.Worksheets.Count)
                    dest_WB.Worksheets(dest_WB.Worksheets.Count).Name = .Cells(cur_row, ds_col).Value2
                End If
                        
            Else
                MsgBox .Cells(cur_row, ss_col).Value2 & " not found in " & .Cells(cur_row, sf_col).Value2
            End If
            
            cur_row = cur_row + 1
        
        Loop
        
        src_WB.Close savechanges:=False
        dest_WB.Close savechanges:=True
        
    End With

    Application.StatusBar = False
    Application.ScreenUpdating = asu
    Application.DisplayAlerts = ada
    Application.EnableEvents = aee
    Application.Calculation = calc
    Application.AutomationSecurity = aas

    MsgBox "done!"

End Sub


Function sheetExists(sheetToFind As String, Optional wb As Workbook) As Boolean
    ' returns True or False if the sheet exists in a workbook
    
    Dim sht As Variant
    sheetExists = False
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    wb.Activate
    For Each sht In wb.Sheets
        If sheetToFind = sht.Name Then
            sheetExists = True
            Exit Function
        End If
    Next sht

End Function


Sub deleteSheet(sht_name As String, Optional wb As Workbook)
    ' procedure deletes sheet if it exists with no prompt
    Dim ada_setting As Boolean
    
    ada_setting = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    If sheetExists(sht_name, wb) Then
        wb.Worksheets(sht_name).Delete
    End If

    Application.DisplayAlerts = ada_setting

End Sub

