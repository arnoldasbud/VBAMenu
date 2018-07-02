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
    Dim selectedRow As Long
    
    With Menu_Form
        selectedRow = .listBoxCables.ListIndex
        
        If selectedRow < 0 Then Exit Sub
        
        .listBoxCables.RemoveItem (selectedRow)
    
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
Public Sub showGlandsResult(arr() As String)
    On Error GoTo HandleErrors
    
    Application.DisplayAlerts = False
    
    Dim arrLength As Long
    Dim newWBook As Workbook
    Dim i As Long
    Dim j As Long
    Dim startRow As Long
    Dim curRow As Long
    
    arrLength = UBound(arr)
    curRow = 2
    
    Set newWBook = Workbooks.Add(xlWBATWorksheet)
    
    With newWBook.Worksheets(1)
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
        
        For i = 1 To UBound(arr)
            startRow = curRow
            
            For j = 1 To UBound(arr, 2)
                If arr(i, j, CellResult.CableDescription) <> vbNullString Then
                
                    .Cells(curRow, CellResult.CableDescription) = _
                        arr(i, j, CellResult.CableDescription)
                    .Cells(curRow, CellResult.GlandDescription) = _
                        arr(i, j, CellResult.GlandDescription)
                    .Cells(curRow, CellResult.Manufacturer) = _
                        arr(i, j, CellResult.Manufacturer)
                    .Cells(curRow, CellResult.Code) = _
                        arr(i, j, CellResult.Code)
                    .Cells(curRow, CellResult.Quantity) = _
                        arr(i, j, CellResult.Quantity)
                    
                    curRow = curRow + 1
                End If
            Next j
            
            'merge description cells if there are multiple glands found
            .Range(.Cells(startRow, CellResult.CableDescription), _
                .Cells(curRow - 1, CellResult.CableDescription)).Merge
        Next i
            
        With .Range(.Cells(2, CellResult.CableDescription), _
                .Cells(curRow - 1, CellResult.CableDescription))
            
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        
        End With
        With .Range(.Cells(2, CellResult.GlandDescription), _
                .Cells(curRow - 1, CellResult.Quantity))
            
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        
        End With
        .Range(.Cells(1, CellResult.CableDescription), _
            .Cells(curRow - 1, CellResult.Quantity)) _
            .Columns.EntireColumn.AutoFit
        
    End With
    
    newWBook.Activate
HandleErrors:
    Application.DisplayAlerts = True
End Sub
