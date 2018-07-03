Attribute VB_Name = "Menu_Main"
Option Explicit

'******************************************************************************
'* Menu_Main
'* Purpose: Setup global variables and constants for entire project. Prepare
'* main procedure showing
'* user form.
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'Constants
Public Const FILEPATH As String = _
    "C:\Users\Arnoldas\AppData\Roaming\Microsoft\Excel\XLSTART\Failai\"

Public Const GLANDSFILE As String = FILEPATH & "Sandarikliai.xlsx"

Public Const CONDUCTOR_CU As String = "Varis"
Public Const CONDUCTOR_AL As String = "Aliuminis"

'Enumerations
'All members represents excel sheet column number and array index
Public Enum CellCable
    Material = 1
    Cable           'cable type
    Cores
    Cross
    Diameter
End Enum

Public Enum CellGland
    Gland = 1       'gland type
    GlandName
    Code
    Manufacturer
    MinDiameter
    MaxDiameter
End Enum

Public Enum CellResult
    CableDescription = 1
    GlandDescription
    Manufacturer
    Code
    Quantity
End Enum

'Arrays
Public gstrArrCables() As String
Public gstrArrGlands() As String

'variable for setting list box item values
Public glngListBoxItems As Long

'global variables for combine function
Public gstrWorkbookName As String
Public gstrWorksheetName As String
Public gstrCombineRange As String
Public glngCombineCodeColumn As Long     'range column index, not sheet column!
Public glngCombineValueColumn As Long     'range column index, not sheet column!

'Application.EnableEvents = false does not work with userforms!!!
'Variable for custom events blocking
Public gblnSkipEvents As Boolean

'******************************************************************************
'* Procedure which shows user form
'*
'* Params: -
'* Return: -
'******************************************************************************
Sub showFormMenu()
    'Load formMenu
    'vbModeless param is required for .Show, as we want to allow changes to
    'worksheet, while user form is open.
    Menu_Form.Show vbModeless
End Sub

'******************************************************************************
'* Procedure which initializes user form and collects data from other workbooks
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub formMenuInitialize()
    'On Error GoTo HandleErrors
    
    gblnSkipEvents = False
    Application.ScreenUpdating = False
    
    glngListBoxItems = 0
    
    Dim objCurrentBook As Object
    Dim lngLastRow As Long
    Dim objCurrentSheet As Object
    
    'collect info from glands file
    Set objCurrentBook = Workbooks.Open(Filename:=GLANDSFILE, ReadOnly:=True)
    Set objCurrentSheet = objCurrentBook.Worksheets("Kabeliai")
    
    lngLastRow = findEmptyCell(objCurrentSheet) - 1
    prepareCableArray objCurrentSheet, lngLastRow
    
    Set objCurrentSheet = objCurrentBook.Worksheets("Sandarikliai")
    
    lngLastRow = findEmptyCell(objCurrentSheet) - 1
    prepareGlandsArray objCurrentSheet, lngLastRow

    With Menu_Form
        .comboBoxCableTypes.Enabled = False
        .comboBoxCableCores.Enabled = False
        .comboBoxCableCross.Enabled = False
    End With
        
HandleErrors:
    objCurrentBook.Close SaveChanges:=False
    Application.ScreenUpdating = True

End Sub

'******************************************************************************
'* Procedure which collects data from given sheet and stores required values
'* into array
'*
'* Params:
'*      shtSheet    - cables worksheet object
'*      lngRows     - number of rows, filled with data
'* Return: -
'******************************************************************************
Private Sub prepareCableArray(ByVal shtSheet As Worksheet, _
        ByVal lngRows As Long)

    Dim i As Long
    
    ReDim gstrArrCables(1 To lngRows, 1 To 5)
    
    With shtSheet
        For i = 2 To lngRows
        
            gstrArrCables(i - 1, CellCable.Material) = _
                .Cells(i, CellCable.Material)
            gstrArrCables(i - 1, CellCable.Cable) = _
                .Cells(i, CellCable.Cable)
            gstrArrCables(i - 1, CellCable.Cores) = _
                .Cells(i, CellCable.Cores)
            gstrArrCables(i - 1, CellCable.Cross) = _
                .Cells(i, CellCable.Cross)
            gstrArrCables(i - 1, CellCable.Diameter) = _
                .Cells(i, CellCable.Diameter)
            
        Next i
    
    End With
End Sub

'******************************************************************************
'* Procedure which collects data from given sheet and stores required values
'* into array
'*
'* Params:
'*      shtSheet    - cables worksheet object
'*      lngRows     - number of rows, filled with data
'* Return: -
'******************************************************************************
Private Sub prepareGlandsArray(ByVal shtSheet As Object, ByVal lngRows As Long)

    Dim i As Long
    
    ReDim gstrArrGlands(1 To lngRows, 1 To 6)
    
    With shtSheet
        For i = 2 To lngRows
            
            gstrArrGlands(i - 1, CellGland.Gland) = _
                .Cells(i, CellGland.Gland)
            gstrArrGlands(i - 1, CellGland.GlandName) = _
                .Cells(i, CellGland.GlandName)
            gstrArrGlands(i - 1, CellGland.Code) = _
                .Cells(i, CellGland.Code)
            gstrArrGlands(i - 1, CellGland.Manufacturer) = _
                .Cells(i, CellGland.Manufacturer)
            gstrArrGlands(i - 1, CellGland.MinDiameter) = _
                .Cells(i, CellGland.MinDiameter)
            gstrArrGlands(i - 1, CellGland.MaxDiameter) = _
                .Cells(i, CellGland.MaxDiameter)
            
        Next i
        
    End With
End Sub
