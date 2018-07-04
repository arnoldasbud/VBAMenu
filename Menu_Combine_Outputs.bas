Attribute VB_Name = "Menu_Combine_Outputs"
Option Explicit
'******************************************************************************
'* Menu_Glands_Outputs
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'******************************************************************************
'* Procedure which stores selected range to new workbook
'*
'* Params: -
'*
'* Return: -
'******************************************************************************
Public Sub prepareCombineResultSheet()

    'check if there are any empty textboxes
    If Menu_Form.textBoxCombineRange = vbNullString _
        Or Menu_Form.textBoxCombineCode = vbNullString _
        Or Menu_Form.textBoxCombineValue = vbNullString Then
        
        MsgBox "Visi laukeliai yra privalomi"
        Exit Sub
    End If
    
    'create new workbook to copy range to
    Dim wkbBook As Workbook
    Set wkbBook = Workbooks.Add(xlWBATWorksheet)

    Dim rngSelectedRange As Range
    Dim lngColumns As Long
    Dim lngRows As Long
    
    Set rngSelectedRange = Workbooks(gstrWorkbookName). _
            Worksheets(gstrWorksheetName).Range(gstrCombineRange)

    'copy selected range and paste it to the first sheet A2 field (second row
    'because we might want to add header to our 'table'
    With rngSelectedRange
        
        lngColumns = .Columns
        lngRows = .Rows
        
        .Copy (wkbBook.Sheets(1).Range("A2"))
    
    End With
    
End Sub
