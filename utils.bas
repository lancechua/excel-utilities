' Miscellaneous functions and subs to make VBA a bit less painful
' Note: requires "Microsoft Scripting Runtime" reference to be selected
Attribute VB_Name = "utils"
Option Base 0
Option Explicit


Function GetFP()
    ' returns the path of the file selected using a dialog box
    
    Dim intChoice As Variant
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        GetFP = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Else
        MsgBox "No File Selected."
        GetFP = Empty
        Exit Function
    End If

End Function


Function sheetExists(sheetToFind As String, Optional wb As Workbook) As Boolean
    ' returns True or False if the sheet exists in a workbook
    
    Dim sht As Variant
    sheetExists = False
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    ' .Sheets used to include chart sheets
    For Each sht In wb.Sheets
        If sheetToFind = sht.Name Then
            sheetExists = True
            Exit Function
        End If
    Next sht

End Function


Function append(arr As Variant, item As Variant) As Variant
    ' function that returns the input array with the item appended at the end
    ' handles empty arrays, but assumes option base 0
        
    On Error GoTo emptyarr
    
    ReDim Preserve arr(UBound(arr) + 1) As Variant
    arr(UBound(arr)) = item
    append = arr
    Exit Function

emptyarr:
    On Error GoTo -1
    ReDim arr(0) As Variant
    arr(0) = item
    append = arr

End Function


Function ColumnLetter(ColumnNumber As Long) As String
    ' Converts an integer column index to its string representation
    ' source: https://stackoverflow.com/a/15366979

    Dim n As Long
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = Trim(s)
End Function


Public Function PosFormat(ParamArray arr() As Variant) As String
    ' Primitive "{0} {1}..."" style formatting using positions
    ' positions, starting from 0, are required
    ' https://stackoverflow.com/a/31730589

    Dim i As Long
    Dim temp As String

    temp = CStr(arr(0))
    For i = 1 To UBound(arr)
        temp = Replace(temp, "{" & i - 1 & "}", CStr(arr(i)))
    Next

    PosFormat = temp
End Function


Public Function DictFormat(mystr As String, my_dict As Scripting.Dictionary) As String
    ' Primitive "{key0} {key1}..."" style formatting using labels instead of positions
    ' https://stackoverflow.com/a/31730589

    Dim ctr As Variant
    Dim temp As String

    temp = mystr
    For Each ctr In my_dict
        temp = Replace(temp, "{" & ctr & "}", my_dict(ctr))
    Next ctr

    DictFormat = temp
End Function


Function parseR1C1(r1c1_ad)
    ' get the string address and returns an array for coordinates Array(row_ind, col_ind)

    Dim temp
    
    temp = Split(Right(r1c1_ad, Len(r1c1_ad) - 1), "C")
    parseR1C1 = Array(CInt(temp(0)), CInt(temp(1)))
    
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


Public Sub Update_Module(strModuleName as String, Optional modulePath as String)
    ' Notes:
    '     - requires "Microsoft Visual Basic for Applications Extensibility" reference
    '     - Macro Settings must be set to "Enable all macros"; with "Trust access to the VBA project object model" checked
    ' sources:
    ' https://www.mrexcel.com/forum/excel-questions/150819-import-module-into-vba-using-vba-macro.html
    ' https://answers.microsoft.com/en-us/office/forum/office_2007-access/using-vba-to-check-if-a-module-exists/82483c2c-406b-4b2b-882f-96e4612ef6fb

    Dim VBProj As Object
    Dim myFileName As String
    Dim mdl As Variant
    
    if modulePath is Nothing Then modulePath = ActiveWorkbook.Path

    myFileName = modulePath & "\" & strModuleName & ".bas"
    
    Set VBProj = Nothing
    On Error Resume Next
    Set VBProj = ActiveWorkbook.VBProject
    On Error GoTo 0

    If VBProj Is Nothing Then
        MsgBox "Update_Module FAILED! -- Workbook is probably not trusted!" & Chr(10) & "Please update module manually."
        Exit Sub
    End If

    If  Dir(myFileName, vbDirectory) = vbNullString Then
        Msgbox myFileName & " does not exist!"
        Exit Sub
    End If 

    With VBProj
        For Each mdl In .vbcomponents
            If mdl.Name = strModuleName And mdl.Type <> vbext_ct_Document Then
                .vbcomponents.Remove mdl
                DoEvents
                Exit For
            End If
        Next mdl
        
        Application.StatusBar = "Importing " & myFileName
        .vbcomponents.Import myFileName
        Application.StatusBar = ""

    End With
    
End Sub