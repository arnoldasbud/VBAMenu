Attribute VB_Name = "Menu_Glands_Inputs"
Option Explicit

'******************************************************************************
'* Menu_Glands_Inputs
'* This file contains procedures which deals with user inputs and interacts
'* with 'Glands' menu
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'******************************************************************************
'* Procedure which adds cable types combo box items, according to selected
'* cable conductor material
'*
'* Params:
'*      cond    - conductor material as String (CONDUCTOR_CU or CONDUCTOR_AL)
'* Return: -
'******************************************************************************
Public Sub prepareCableTypesBox(ByVal cond As String)
    On Error GoTo HandleErrors

    Application.ScreenUpdating = False
    gblnSkipEvents = True

    'isjungiam ir isvalom kitus pasirinkimo laukelius
    With Menu_Form
        .comboBoxCableTypes.Clear
        .comboBoxCableCores.Clear
        .comboBoxCableCross.Clear
    
        .comboBoxCableCores.Enabled = False
        .comboBoxCableCross.Enabled = False
    
        'surasom tipus y komboboxa
        Dim i As Long
        Dim arrSize As Long
        
        arrSize = UBound(gstrArrCables)
        
        For i = 1 To arrSize
            If gstrArrCables(i, CellCable.Material) = cond Then
                .comboBoxCableTypes = gstrArrCables(i, CellCable.Cable)
                
                If Not .comboBoxCableTypes.MatchFound Then
                    .comboBoxCableTypes.AddItem gstrArrCables(i, CellCable.Cable)
                End If
            End If
        Next i
        
        'yjungiam combobox
        .comboBoxCableTypes = ""
        .comboBoxCableTypes.Enabled = True
    
    End With
    
HandleErrors:
    Application.ScreenUpdating = True
    gblnSkipEvents = False
End Sub

'******************************************************************************
'* Procedure which adds cable cores combo box items, according to selected
'* cable type
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub prepareCableCoresBox()
    On Error GoTo HandleErrors
    
    If gblnSkipEvents Then Exit Sub
    
    gblnSkipEvents = True
    Application.ScreenUpdating = False
    
    'isvalom gyslu ir skerspjuviu bosus
    With Menu_Form
        .comboBoxCableCores.Clear
        .comboBoxCableCross.Clear
        
        .comboBoxCableCross.Enabled = False
        
        'surasom kabelio gyslu skaiciu y komboboxa
        Dim i As Long
        Dim arrSize As Long
        Dim conductor As String
        
        If .optionCableCu Then
            conductor = CONDUCTOR_CU
        Else
            conductor = CONDUCTOR_AL
        End If
        
        arrSize = UBound(gstrArrCables)
        
        For i = 1 To arrSize
            If gstrArrCables(i, CellCable.Material) = conductor And gstrArrCables(i, CellCable.Cable) = .comboBoxCableTypes Then
                .comboBoxCableCores = gstrArrCables(i, CellCable.Cores)
                
                If Not .comboBoxCableCores.MatchFound Then
                    .comboBoxCableCores.AddItem gstrArrCables(i, CellCable.Cores)
                End If
            End If
        Next i
        
        .comboBoxCableCores = ""
        .comboBoxCableCores.Enabled = True
    
    End With
    
HandleErrors:
    gblnSkipEvents = False
    Application.ScreenUpdating = True
End Sub

'******************************************************************************
'* Procedure which adds cable cross combo box items, according to selected
'* cable cores
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub prepareCableCrossBox()
    On Error GoTo HandleErrors
    
    If gblnSkipEvents Then Exit Sub
    
    Application.ScreenUpdating = False
    
    'isvalom skerspjuviu comboboxa
    With Menu_Form
        .comboBoxCableCross.Clear
        
        'surasom kabelio gyslu skaiciu y komboboxa
        Dim i As Long
        Dim arrSize As Long
        Dim conductor As String
        
        If .optionCableCu Then
            conductor = CONDUCTOR_CU
        Else
            conductor = CONDUCTOR_AL
        End If
        
        arrSize = UBound(gstrArrCables)
        
        For i = 1 To arrSize
            If gstrArrCables(i, CellCable.Material) = conductor And gstrArrCables(i, CellCable.Cable) = .comboBoxCableTypes _
                And gstrArrCables(i, CellCable.Cores) = .comboBoxCableCores Then
                .comboBoxCableCross = gstrArrCables(i, CellCable.Cross)
                
                If Not .comboBoxCableCross.MatchFound Then
                    .comboBoxCableCross.AddItem gstrArrCables(i, CellCable.Cross)
                End If
            End If
        Next i
        
        .comboBoxCableCross = ""
        .comboBoxCableCross.Enabled = True
    
    End With
    
HandleErrors:
    Application.ScreenUpdating = True
End Sub

'******************************************************************************
'* Procedure which adds cable description, diameter and quantity values
'* to cable list box
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub prepareCableListBox()
    'patikrinam ar pasirinktos reiksmes teisingos
    With Menu_Form
        If Not .comboBoxCableTypes.MatchFound Then
            MsgBox "Neteisingas kabelio tipas"
            Exit Sub
        ElseIf Not .comboBoxCableCores.MatchFound Then
            MsgBox "Neteisingas kabelio gyslu skaicius"
            Exit Sub
        ElseIf Not .comboBoxCableCross.MatchFound Then
            MsgBox "Neteisingas kabelio skerspjuvis"
            Exit Sub
        ElseIf Not IsNumeric(.textBoxQuantity) Then
            MsgBox "Neteisingas kabeliu kiekis"
            Exit Sub
        End If
        
        'Surenkam surasyta informacija apie kabely
        Dim cableType As String
        Dim cableCores As String
        Dim cableCross As String
        Dim cablesQuant As Long
        
        cableType = .comboBoxCableTypes
        cableCores = .comboBoxCableCores
        cableCross = .comboBoxCableCross
        cablesQuant = .textBoxQuantity
        
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
                .listBoxCables.AddItem
                .listBoxCables.List(glngListBoxItems, 0) = "Kabelis " + cableType + " " + cableCores + "x" + cableCross
                .listBoxCables.List(glngListBoxItems, 1) = gstrArrCables(i, CellCable.Diameter)
                .listBoxCables.List(glngListBoxItems, 2) = cablesQuant
                cableFound = True
                
                glngListBoxItems = glngListBoxItems + 1
                Exit For
            End If
        Next i
        
        If Not cableFound Then
            MsgBox "Toks kabelis nerastas"
        End If
    
    End With
    
End Sub

'******************************************************************************
'* Procedure which searches for glands for every cable in the listbox and
'* prepares array with result
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub prepareResultArray()
    Dim listItems As Long
    
    With Menu_Form
        listItems = .listBoxCables.ListCount
        
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
            cableDiam = CDbl(.listBoxCables.List(i, 1))
        
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
                    arrayResult(i + 1, curGlands, CellResult.CableDescription) = .listBoxCables.List(i, 0) + ", " + CStr(cableDiam) + "mm"
                    arrayResult(i + 1, curGlands, CellResult.GlandDescription) = gstrArrGlands(j, CellGland.TypeName) + _
                        " (" + CStr(cableMinD) + "mm-" + CStr(cableMaxD) + "mm)"
                    arrayResult(i + 1, curGlands, CellResult.Manufacturer) = gstrArrGlands(j, CellGland.Manufacturer)
                    arrayResult(i + 1, curGlands, CellResult.Code) = gstrArrGlands(j, CellGland.Code)
                    arrayResult(i + 1, curGlands, CellResult.Quantity) = .listBoxCables.List(i, 2)
                    
                End If
            Next j
        
            ' jei neradom ytraukiam y masyva ir prirasom, kad reikiamas sandariklis nerastas.c
            If curGlands < 1 Then
                arrayResult(i + 1, 1, CellResult.CableDescription) = .listBoxCables.List(i, 0) + ", " + CStr(cableDiam) + "mm"
                arrayResult(i + 1, 1, CellResult.GlandDescription) = "Nerastas"
            End If
            
        Next i
    End With
    
    'surasom viska y nauja faila
    showGlandsResult arrayResult
    
End Sub
