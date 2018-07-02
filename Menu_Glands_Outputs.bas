Attribute VB_Name = "Menu_Glands_Outputs"
Option Explicit

'******************************************************************************
'* Menu_Glands_Outputs
'* This file contains procedures which prepares and shows output sheets,
'* updates form items values according to user selection
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'******************************************************************************
'* Procedure that removes selected item from listbox
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub listBoxRemoveCable()
    Dim lngRow As Long
    
    With Menu_Form
        lngRow = .listBoxCables.ListIndex
        
        If lngRow < 0 Then Exit Sub
        
        .listBoxCables.RemoveItem (lngRow)
    
    End With
    
    glngListBoxItems = glngListBoxItems - 1
End Sub

'******************************************************************************
'* Procedure that creates new workbook with info from array
'*
'* Params:
'*      arr()   - array of found suitable glands as string
'* Return: -
'******************************************************************************
Public Sub showGlandsResult(strArray() As String)
    On Error GoTo HandleErrors
    
    Application.DisplayAlerts = False
    
    Dim lngLength As Long
    Dim wkbBook As Workbook
    Dim i As Long
    Dim j As Long
    Dim lngStartRow As Long
    Dim lngRow As Long
    
    lngLength = UBound(strArray)
    lngRow = 2
    
    Set wkbBook = Workbooks.Add(xlWBATWorksheet)
    
    With wkbBook.Worksheets(1)
        .Name = "Sandarikliai"
        
        .Cells(1, CellResult.CableDescription) = "Kabelis"
        .Cells(1, CellResult.GlandDescription) = "Sandariklis"
        .Cells(1, CellResult.Manufacturer) = "Gamintojas"
        .Cells(1, CellResult.Code) = "Kodas"
        .Cells(1, CellResult.Quantity) = "Kiekis"
        
        'CableDescription must be first constant in enum and CellResult - last,
        'otherwise code will not work.
        With .Range(.Cells(1, CellResult.CableDescription), _
                .Cells(1, CellResult.Quantity))
            
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        
        End With
        
        For i = 1 To UBound(strArray)
            lngStartRow = lngRow
            
            For j = 1 To UBound(strArray, 2)
                If strArray(i, j, CellResult.CableDescription) <> vbNullString Then
                
                    .Cells(lngRow, CellResult.CableDescription) = _
                        strArray(i, j, CellResult.CableDescription)
                    .Cells(lngRow, CellResult.GlandDescription) = _
                        strArray(i, j, CellResult.GlandDescription)
                    .Cells(lngRow, CellResult.Manufacturer) = _
                        strArray(i, j, CellResult.Manufacturer)
                    .Cells(lngRow, CellResult.Code) = _
                        strArray(i, j, CellResult.Code)
                    .Cells(lngRow, CellResult.Quantity) = _
                        strArray(i, j, CellResult.Quantity)
                    
                    lngRow = lngRow + 1
                End If
            Next j
            
            'merge description cells if there are multiple glands found
            .Range(.Cells(lngStartRow, CellResult.CableDescription), _
                .Cells(lngRow - 1, CellResult.CableDescription)).Merge
        Next i
            
        With .Range(.Cells(2, CellResult.CableDescription), _
                .Cells(lngRow - 1, CellResult.CableDescription))
            
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        
        End With
        With .Range(.Cells(2, CellResult.GlandDescription), _
                .Cells(lngRow - 1, CellResult.Quantity))
            
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        
        End With
        .Range(.Cells(1, CellResult.CableDescription), _
            .Cells(lngRow - 1, CellResult.Quantity)) _
            .Columns.EntireColumn.AutoFit
        
    End With
    
    wkbBook.Activate
HandleErrors:
    Application.DisplayAlerts = True
End Sub
