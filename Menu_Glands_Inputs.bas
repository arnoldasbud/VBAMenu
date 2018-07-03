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
Public Sub prepareCableTypesBox(ByVal strConductor As String)
    On Error GoTo HandleErrors

    Application.ScreenUpdating = False
    gblnSkipEvents = True

    With Menu_Form
        'lets clear additional comboboxes when main combobox changes
        .comboBoxCableTypes.Clear
        .comboBoxCableCores.Clear
        .comboBoxCableCross.Clear
    
        .comboBoxCableCores.Enabled = False
        .comboBoxCableCross.Enabled = False
    
        Dim i As Long
        Dim lngSize As Long
        
        lngSize = UBound(gstrArrCables)
        
        'fill combobox with items
        For i = 1 To lngSize
            If gstrArrCables(i, CellCable.Material) = strConductor Then
                .comboBoxCableTypes = gstrArrCables(i, CellCable.Cable)
                
                If Not .comboBoxCableTypes.MatchFound Then
                    .comboBoxCableTypes.AddItem _
                        gstrArrCables(i, CellCable.Cable)
                End If
            End If
        Next i
        
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
    
    With Menu_Form
        .comboBoxCableCores.Clear
        .comboBoxCableCross.Clear
        
        .comboBoxCableCross.Enabled = False
        
        Dim i As Long
        Dim lngSize As Long
        Dim strConductor As String
        
        If .optionCableCu Then
            strConductor = CONDUCTOR_CU
        Else
            strConductor = CONDUCTOR_AL
        End If
        
        lngSize = UBound(gstrArrCables)
        
        For i = 1 To lngSize
            If gstrArrCables(i, CellCable.Material) = strConductor And _
                gstrArrCables(i, CellCable.Cable) = .comboBoxCableTypes Then
                
                .comboBoxCableCores = gstrArrCables(i, CellCable.Cores)
                
                If Not .comboBoxCableCores.MatchFound Then
                    .comboBoxCableCores.AddItem _
                        gstrArrCables(i, CellCable.Cores)
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
    
    With Menu_Form
        .comboBoxCableCross.Clear
        
        Dim i As Long
        Dim lngSize As Long
        Dim strConductor As String
        
        If .optionCableCu Then
            strConductor = CONDUCTOR_CU
        Else
            strConductor = CONDUCTOR_AL
        End If
        
        lngSize = UBound(gstrArrCables)
        
        For i = 1 To lngSize
            If gstrArrCables(i, CellCable.Material) = strConductor _
                And gstrArrCables(i, CellCable.Cable) = .comboBoxCableTypes _
                And gstrArrCables(i, CellCable.Cores) = .comboBoxCableCores _
                Then
                
                .comboBoxCableCross = gstrArrCables(i, CellCable.Cross)
                
                If Not .comboBoxCableCross.MatchFound Then
                    .comboBoxCableCross.AddItem _
                        gstrArrCables(i, CellCable.Cross)
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
    'user input validation
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
        
        Dim strType As String
        Dim strCores As String
        Dim strCross As String
        Dim lngQuant As Long
        
        strType = .comboBoxCableTypes
        strCores = .comboBoxCableCores
        strCross = .comboBoxCableCross
        lngQuant = .textBoxQuantity
        
        If lngQuant < 1 Then
            MsgBox "Neteisingas kiekis"
            Exit Sub
        End If
        
        Dim lngElements As Long
        lngElements = UBound(gstrArrCables)
        
        Dim i As Long
        Dim blnFound As Boolean
        blnFound = False
        For i = 1 To lngElements
            If strType = gstrArrCables(i, CellCable.Cable) _
                And strCores = gstrArrCables(i, CellCable.Cores) _
                And strCross = gstrArrCables(i, CellCable.Cross) Then
            
                .listBoxCables.AddItem
                .listBoxCables.List(glngListBoxItems, 0) = _
                    "Kabelis " & strType & " " & strCores & "x" & strCross
                .listBoxCables.List(glngListBoxItems, 1) = _
                    gstrArrCables(i, CellCable.Diameter)
                .listBoxCables.List(glngListBoxItems, 2) = _
                    lngQuant
                blnFound = True
                
                glngListBoxItems = glngListBoxItems + 1
                Exit For
            End If
        Next i
        
        If Not blnFound Then
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
    Dim lngItems As Long
    
    With Menu_Form
        lngItems = .listBoxCables.ListCount
        
        If lngItems < 1 Then Exit Sub
        
        Dim i As Long
        Dim j As Long
        Dim dblDiam As Double
        Dim dblMinD As Double
        Dim dblMaxD As Double
        Dim lngGlands As Long
        Dim lngMaxGlands As Long
        
        lngMaxGlands = 10
        
        Dim strArrResult() As String
        ReDim strArrResult(1 To lngItems, 1 To lngMaxGlands, 1 To 5)
        'begin search for suitable glands for cable
        For i = 0 To lngItems - 1
        
            lngGlands = 0
            
            dblDiam = CDbl(.listBoxCables.List(i, 1))
        
            For j = 1 To UBound(gstrArrGlands)
                
                If gstrArrGlands(j, CellGland.MinDiameter) = vbNullString _
                    Or gstrArrGlands(j, CellGland.MaxDiameter) = vbNullString _
                    Then
                    
                    Exit For
                
                End If
                
                dblMinD = CDbl(gstrArrGlands(j, CellGland.MinDiameter))
                dblMaxD = CDbl(gstrArrGlands(j, CellGland.MaxDiameter))
                
                If dblDiam < dblMaxD And dblDiam > dblMinD Then
                    lngGlands = lngGlands + 1
                    
                    'expand array if necessary
                    If lngGlands > lngMaxGlands Then
                        lngMaxGlands = lngMaxGlands + 1
                        
                        ReDim Preserve strArrResult(1 To lngItems, _
                            1 To lngMaxGlands, 1 To 5)
                    End If
                    
                    'insert found gland to the array
                    strArrResult(i + 1, lngGlands, CellResult.CableDescription) = _
                        .listBoxCables.List(i, 0) & ", " & dblDiam & "mm"
                    strArrResult(i + 1, lngGlands, CellResult.GlandDescription) = _
                        gstrArrGlands(j, CellGland.GlandName) & " (" & _
                        dblMinD & "mm-" & dblMaxD & "mm)"
                    strArrResult(i + 1, lngGlands, CellResult.Manufacturer) = _
                        gstrArrGlands(j, CellGland.Manufacturer)
                    strArrResult(i + 1, lngGlands, CellResult.Code) = _
                        gstrArrGlands(j, CellGland.Code)
                    strArrResult(i + 1, lngGlands, CellResult.Quantity) = _
                        .listBoxCables.List(i, 2)
                    
                End If
            Next j
        
            If lngGlands < 1 Then
                strArrResult(i + 1, 1, CellResult.CableDescription) = _
                    .listBoxCables.List(i, 0) & ", " & dblDiam & "mm"
                strArrResult(i + 1, 1, CellResult.GlandDescription) = "Nerastas"
            End If
            
        Next i
    End With
    
    showGlandsResult strArrResult
    
End Sub
