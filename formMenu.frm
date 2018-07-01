VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMenu 
   Caption         =   "Meniu"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   OleObjectBlob   =   "formMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'**************************************************************************************************
'* formMenu
'* Purpose: Control UserForm events
'*
'* Bugs: -
'*
'* To do: -
'*
'**************************************************************************************************

'**************************************************************************************************
'* Event that loads userform
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub UserForm_Initialize()
    formMenuInitialize
End Sub

'**************************************************************************************************
'* Event that fires up when user selects option
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub optionCableCu_Click()
    prepareCableTypesBox CONDUCTOR_CU
End Sub

'**************************************************************************************************
'* Event that fires up when user selects option
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub optionCableAl_Click()
    prepareCableTypesBox CONDUCTOR_AL
End Sub

'**************************************************************************************************
'* Event that fires up when user changes combobox selection
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub comboBoxCableTypes_Change()
    'On Error GoTo HandleErrors
    
    If gblnSkipEvents Then Exit Sub
    
    gblnSkipEvents = True
    Application.ScreenUpdating = False
    
    'isvalom gyslu ir skerspjuviu bosus
    comboBoxCableCores.Clear
    comboBoxCableCross.Clear
    
    comboBoxCableCross.Enabled = False
    
    'surasom kabelio gyslu skaiciu y komboboxa
    Dim i As Long
    Dim arrSize As Long
    Dim conductor As String
    
    If optionCableCu Then
        conductor = CONDUCTOR_CU
    Else
        conductor = CONDUCTOR_AL
    End If
    
    arrSize = UBound(gstrArrCables)
    
    For i = 1 To arrSize
        If gstrArrCables(i, CellCable.Material) = conductor And gstrArrCables(i, CellCable.Cable) = comboBoxCableTypes Then
            comboBoxCableCores = gstrArrCables(i, CellCable.Cores)
            
            If Not comboBoxCableCores.MatchFound Then
                comboBoxCableCores.AddItem gstrArrCables(i, CellCable.Cores)
            End If
        End If
    Next i
    
    comboBoxCableCores = ""
    comboBoxCableCores.Enabled = True
    
HandleErrors:
    gblnSkipEvents = False
    Application.ScreenUpdating = True
End Sub

'jungiam skerspjuvio pasirinkima
Private Sub comboBoxCableCores_Change()
    'On Error GoTo HandleErrors
    
    If gblnSkipEvents Then Exit Sub
    
    Application.ScreenUpdating = False
    
    'isvalom skerspjuviu comboboxa
    comboBoxCableCross.Clear
    
    'surasom kabelio gyslu skaiciu y komboboxa
    Dim i As Long
    Dim arrSize As Long
    Dim conductor As String
    
    If optionCableCu Then
        conductor = CONDUCTOR_CU
    Else
        conductor = CONDUCTOR_AL
    End If
    
    arrSize = UBound(gstrArrCables)
    
    For i = 1 To arrSize
        If gstrArrCables(i, CellCable.Material) = conductor And gstrArrCables(i, CellCable.Cable) = comboBoxCableTypes _
            And gstrArrCables(i, CellCable.Cores) = comboBoxCableCores Then
            comboBoxCableCross = gstrArrCables(i, CellCable.Cross)
            
            If Not comboBoxCableCross.MatchFound Then
                comboBoxCableCross.AddItem gstrArrCables(i, CellCable.Cross)
            End If
        End If
    Next i
    
    comboBoxCableCross = ""
    comboBoxCableCross.Enabled = True
    
HandleErrors:
    Application.ScreenUpdating = True
End Sub

Private Sub buttonCableAdd_Click()
    'patikrinam ar pasirinktos reiksmes teisingos
    If Not comboBoxCableTypes.MatchFound Then
        MsgBox "Neteisingas kabelio tipas"
        Exit Sub
    ElseIf Not comboBoxCableCores.MatchFound Then
        MsgBox "Neteisingas kabelio gyslu skaicius"
        Exit Sub
    ElseIf Not comboBoxCableCross.MatchFound Then
        MsgBox "Neteisingas kabelio skerspjuvis"
        Exit Sub
    ElseIf Not IsNumeric(textBoxQuantity) Then
        MsgBox "Neteisingas kabeliu kiekis"
        Exit Sub
    End If
    
    'Surenkam surasyta informacija apie kabely
    Dim cableType As String
    Dim cableCores As String
    Dim cableCross As String
    Dim cablesQuant As Long
    
    cableType = comboBoxCableTypes
    cableCores = comboBoxCableCores
    cableCross = comboBoxCableCross
    cablesQuant = textBoxQuantity
    
    If cablesQuant < 1 Then
        MsgBox "Neteisingas kiekis"
        Exit Sub
    End If
    
    Dim arrElements As Long
    arrElements = UBound(gstrArrCables)
    
    'ieskom pasirinkto kabelio masyve
    Dim i As Long
    Dim cableFound As Boolean
    cableFound = False
    For i = 1 To arrElements
        If cableType = gstrArrCables(i, CellCable.Cable) And cableCores = gstrArrCables(i, CellCable.Cores) _
            And cableCross = gstrArrCables(i, CellCable.Cross) Then
        
            'kabelis rastas, ytraukiam nauja elementa y kabeliu sarasa
            listBoxCables.AddItem
            listBoxCables.List(glngListBoxItems, 0) = "Kabelis " + cableType + " " + cableCores + "x" + cableCross
            listBoxCables.List(glngListBoxItems, 1) = gstrArrCables(i, CellCable.Diameter)
            listBoxCables.List(glngListBoxItems, 2) = cablesQuant
            cableFound = True
            
            glngListBoxItems = glngListBoxItems + 1
            Exit For
        End If
    Next i
    
    If Not cableFound Then
        MsgBox "Toks kabelis nerastas"
    End If
    
End Sub

Private Sub buttonCableRemove_Click()
    
    Dim selectedRow As Long
    selectedRow = listBoxCables.ListIndex
    
    If selectedRow < 0 Then Exit Sub
    
    'istrinam kabely is kolekcijos ir lenteles
    listBoxCables.RemoveItem (selectedRow)
    
    glngListBoxItems = glngListBoxItems - 1
    
End Sub

Private Sub buttonSearchGlands_Click()
    
    Dim listItems As Long
    listItems = listBoxCables.ListCount
    
    'patikrinam ar yra ko ieskoti
    If listItems < 1 Then Exit Sub
    
    Dim i As Long
    Dim j As Long
    Dim cableDiam As Double
    Dim cableMinD As Double
    Dim cableMaxD As Double
    Dim curGlands As Long
    Dim maxGlands As Long
    
    maxGlands = 10
    
    Dim arrayResult() As String
    ReDim arrayResult(1 To listItems, 1 To maxGlands, 1 To 5)
    'ieskom tinkamo sandariklio kabeliui
    For i = 0 To listItems - 1
    
        curGlands = 0
        
        'issisaugom diametra
        cableDiam = CDbl(listBoxCables.List(i, 1))
    
        'paieskom sandarikliu masyve tinkamo
        For j = 1 To UBound(gstrArrGlands)
            
            If gstrArrGlands(j, CellGland.MinDiameter) = vbNullString Or gstrArrGlands(j, CellGland.MaxDiameter) = vbNullString Then
                Exit For
            End If
            
            cableMinD = CDbl(gstrArrGlands(j, CellGland.MinDiameter))
            cableMaxD = CDbl(gstrArrGlands(j, CellGland.MaxDiameter))
            
            If cableDiam < cableMaxD And cableDiam > cableMinD Then
                curGlands = curGlands + 1
                
                'jei tinkantys sandarikliai nebetelpa y masyva, padidinam jy
                If curGlands > maxGlands Then
                    maxGlands = maxGlands + 1
                    
                    ReDim Preserve arrayResult(1 To listItems, 1 To maxGlands, 1 To 5)
                End If
                
                'issaugojam sandarikly y masyva
                arrayResult(i + 1, curGlands, CellResult.CableDescription) = listBoxCables.List(i, 0) + ", " + CStr(cableDiam) + "mm"
                arrayResult(i + 1, curGlands, CellResult.GlandDescription) = gstrArrGlands(j, CellGland.TypeName) + _
                    " (" + CStr(cableMinD) + "mm-" + CStr(cableMaxD) + "mm)"
                arrayResult(i + 1, curGlands, CellResult.Manufacturer) = gstrArrGlands(j, CellGland.Manufacturer)
                arrayResult(i + 1, curGlands, CellResult.Code) = gstrArrGlands(j, CellGland.Code)
                arrayResult(i + 1, curGlands, CellResult.Quantity) = listBoxCables.List(i, 2)
                
            End If
        Next j
    
        ' jei neradom ytraukiam y masyva ir prirasom, kad reikiamas sandariklis nerastas.c
        If curGlands < 1 Then
            arrayResult(i + 1, 1, CellResult.CableDescription) = listBoxCables.List(i, 0) + ", " + CStr(cableDiam) + "mm"
            arrayResult(i + 1, 1, CellResult.GlandDescription) = "Nerastas"
        End If
        
    Next i
    
    'surasom viska y nauja faila
    showGlandsResult arrayResult
    
End Sub

'funkcija paruosia kabeliu tipu comboboxa
Private Function prepareCableTypesBox(ByVal cond As String)
    'On Error GoTo HandleErrors

    Application.ScreenUpdating = False
    gblnSkipEvents = True

    'isjungiam ir isvalom kitus pasirinkimo laukelius
    comboBoxCableTypes.Clear
    comboBoxCableCores.Clear
    comboBoxCableCross.Clear
    
    comboBoxCableCores.Enabled = False
    comboBoxCableCross.Enabled = False
    
    'surasom tipus y komboboxa
    Dim i As Long
    Dim arrSize As Long
    
    arrSize = UBound(gstrArrCables)
    
    For i = 1 To arrSize
        If gstrArrCables(i, CellCable.Material) = cond Then
            comboBoxCableTypes = gstrArrCables(i, CellCable.Cable)
            
            If Not comboBoxCableTypes.MatchFound Then
                comboBoxCableTypes.AddItem gstrArrCables(i, CellCable.Cable)
            End If
        End If
    Next i
    
    'yjungiam combobox
    comboBoxCableTypes = ""
    comboBoxCableTypes.Enabled = True
    
HandleErrors:
    Application.ScreenUpdating = True
    gblnSkipEvents = False
End Function

'funkcija, kuri masyva suraso y nauja faila
Private Sub showGlandsResult(arr() As String)
    'On Error GoTo HandleErrors
    
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
    
    'Surasom masyvo reiksmes y nauja knyga
    With newWBook.Worksheets(1)
        .Name = "Sandarikliai"
        
        .Cells(1, CellResult.CableDescription) = "Kabelis"
        .Cells(1, CellResult.GlandDescription) = "Sandariklis"
        .Cells(1, CellResult.Manufacturer) = "Gamintojas"
        .Cells(1, CellResult.Code) = "Kodas"
        .Cells(1, CellResult.Quantity) = "Kiekis"
        'reikia ysitikinti, kad RES_CALL_DESC reiksme yra maziausia, o CellResult.Quantity - didziausia,
        'antraip kodas normaliai neveiks.
        With .Range(.Cells(1, CellResult.CableDescription), .Cells(1, CellResult.Quantity))
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        For i = 1 To UBound(arr)
            'kad zinotume, kurias eilutes sujungti
            startRow = curRow
            
            For j = 1 To UBound(arr, 2)
                If arr(i, j, CellResult.CableDescription) <> vbNullString Then
                    .Cells(curRow, CellResult.CableDescription) = arr(i, j, CellResult.CableDescription)
                    .Cells(curRow, CellResult.GlandDescription) = arr(i, j, CellResult.GlandDescription)
                    .Cells(curRow, CellResult.Manufacturer) = arr(i, j, CellResult.Manufacturer)
                    .Cells(curRow, CellResult.Code) = arr(i, j, CellResult.Code)
                    .Cells(curRow, CellResult.Quantity) = arr(i, j, CellResult.Quantity)
                    
                    curRow = curRow + 1
                End If
            Next j
            
            'sujungiam eilutes, kuriose pasikartoja kabelio pavadinimas
            .Range(.Cells(startRow, CellResult.CableDescription), .Cells(curRow - 1, CellResult.CableDescription)).Merge
        Next i
            
        'sulygiuojam ir sutalpinam viska y stulpelius
        With .Range(.Cells(2, CellResult.CableDescription), .Cells(curRow - 1, CellResult.CableDescription))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        With .Range(.Cells(2, CellResult.GlandDescription), .Cells(curRow - 1, CellResult.Quantity))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Range(.Cells(1, CellResult.CableDescription), .Cells(curRow - 1, CellResult.Quantity)).Columns.EntireColumn.AutoFit
        
    End With
    
    newWBook.Activate
HandleErrors:
    Application.DisplayAlerts = True
End Sub

'testavimo funkcija, atspausdina visas masyvo reiksmes
Private Function printArray(arr() As String)

    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next i
    
End Function
