Attribute VB_Name = "Menu_Combine_Inputs"
Option Explicit

'******************************************************************************
'* Menu_Glands_Inputs
'* This file contains procedures which deals with user inputs and interacts
'* with 'Combine lines' menu
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'******************************************************************************
'* Procedure which stores selected range address to textBoxCombineRange value
'*
'* Params: -
'*
'* Return: -
'******************************************************************************
Public Sub addSelectedRange()
    
    'We are interesed only in range selection
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection
    
        Dim lngRows As Long
        Dim lngColumns As Long
        
        lngRows = .Rows.Count
        lngColumns = .Columns.Count
        'we need at least 2 rows and 2 cols
        If lngRows < 2 Or lngColumns < 2 Then
            MsgBox "Pasirinkta per maza sritis"
            Exit Sub
        End If
    
        Dim strWorkbook As String
        Dim strWorksheet As String
        Dim strRange As String
    
        'set text box value with combinet worksheet, workbook and range values
        strWorksheet = .Parent.Name
        strWorkbook = .Parent.Parent.Name
        strRange = .Cells(1, 1).Address(False, False) & ":" & _
            .Cells(1, 1).Offset(lngRows - 1, lngColumns - 1).Address(False, False)
    
        Menu_Form.textBoxCombineRange = "[" & strWorkbook & "]" & strWorksheet _
            & "!" & strRange
        
        gstrWorkbookName = strWorkbook
        gstrWorksheetName = strWorksheet
        gstrCombineRange = strRange
        
    End With

End Sub

'******************************************************************************
'* Procedure which stores selected column letter to textbox
'*
'* Params: -
'*
'* Return: -
'******************************************************************************
Public Sub addSelectedCodeColumn()

    'We are interesed only in range selection
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection

        If .Columns.Count > 1 Then
            MsgBox "Pasirink tik viena stulpely"
            Exit Sub
        End If

        'we have to be sure, that we are on the same workbook and sheet.
        If gstrWorkbookName <> .Parent.Parent.Name _
            Or gstrWorksheetName <> .Parent.Name Then
            
            MsgBox "Pasirinktas stulpelis yra ne tame paciame lape ar faile"
            Exit Sub
        End If

        'get selected column letter
        Dim strLetter As String
        strLetter = getColumnLetter(.Column)
        
        'range must intersect with selected column
        If Application.Intersect(Range(gstrCombineRange), _
            Range(strLetter & ":" & strLetter)) Is Nothing Then
            
            MsgBox "Pasirinktas stulpelis nepapuola y pasirinkta srity"
            Exit Sub
        End If
        
        Menu_Form.textBoxCombineCode = strLetter
        glngCombineCodeColumn = .Column
        
    End With

End Sub

'******************************************************************************
'* Procedure which stores selected column letter to textbox
'*
'* Params: -
'*
'* Return: -
'******************************************************************************
Public Sub addSelectedValueColumn()

    'We are interesed only in range selection
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    With Selection

        If .Columns.Count > 1 Then
            MsgBox "Pasirink tik viena stulpely"
            Exit Sub
        End If

        'we have to be sure, that we are on the same workbook and sheet.
        If gstrWorkbookName <> .Parent.Parent.Name _
            Or gstrWorksheetName <> .Parent.Name Then
            
            MsgBox "Pasirinktas stulpelis yra ne tame paciame lape ar faile"
            Exit Sub
        End If

        'get selected column letter
        Dim strLetter As String
        strLetter = getColumnLetter(.Column)
        
        'range must intersect with selected column
        If Application.Intersect(Range(gstrCombineRange), _
            Range(strLetter & ":" & strLetter)) Is Nothing Then
            
            MsgBox "Pasirinktas stulpelis nepapuola y pasirinkta srity"
            Exit Sub
        End If
        
        'value and code columns can't be the same
        If strLetter = Menu_Form.textBoxCombineCode Then
            MsgBox "Kodo ir reiksmes stulpelis negali buti tas pats"
            Exit Sub
        End If
        
        Menu_Form.textBoxCombineValue = strLetter
        glngCombineValueColumn = .Column
        
    End With

End Sub

'******************************************************************************
'* Procedure which stores selected range to array
'*
'* Params: -
'*
'* Return: -
'******************************************************************************
Public Sub prepareCombineResultSheet()

    'create new workbook to copy range to
    

End Sub

