Attribute VB_Name = "Menu_Misc"
Option Explicit

'******************************************************************************
'* Menu_Misc
'* Purpose: Contains functions that are used over the project
'*
'* Bugs: -
'*
'* To do: -
'*
'******************************************************************************

'******************************************************************************
'* Function which continues to looking for a empty cell in given column
'*
'* Params:
'*      shtSheet    - sheet to search in
'*      lngCol      - column in which we search for empty cell
'* Return:
'*      Last non empty cell in given column, as row number
'******************************************************************************
Public Function findEmptyCell(ByVal shtSheet As Worksheet, _
        Optional ByVal lngCol As Long = 1) As Long

    Dim lngRow As Long
    lngRow = 1

    Do Until IsEmpty(shtSheet.Cells(lngRow, lngCol))
        lngRow = lngRow + 1
    Loop
    
    findEmptyCell = lngRow - 1

End Function

'******************************************************************************
'* Procedure which exports all current project modules to project folder
'*
'* Params: -
'* Return: -
'******************************************************************************
Public Sub exportModules()

    Const PROJECT_PATH As String = "C:\Users\Arnoldas\Projects\VBA\Menu\"
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_MSForm As Long = 3

    If Dir(PROJECT_PATH, vbDirectory) = "" Then
        MkDir PROJECT_PATH
    End If

    Dim objModules As Object
    Dim objMod As Object
    Dim lngType As Long
    Dim strExtension As String
    Dim blnModFound As Boolean

    Set objModules = ActiveWorkbook.VBProject.VBComponents

    For Each objMod In objModules
        
        lngType = objMod.Type
        
        'We are interesed only in standart modules and forms
        Select Case lngType
            Case vbext_ct_StdModule
                strExtension = ".bas"
                blnModFound = True
            Case vbext_ct_MSForm
                strExtension = ".frm"
                blnModFound = True
            Case Else
                blnModFound = False
        End Select
        
        If blnModFound Then
            objMod.Export (PROJECT_PATH & objMod.Name & strExtension)
            'MsgBox PROJECT_PATH & objMod.Name & strExtension
        End If
        
    Next objMod
    
End Sub

'******************************************************************************
'* Function which converts column number to char
'*
'* Params:
'*      lngColumn - column id as long
'* Return:
'*      Given column letter as string
'******************************************************************************
Public Function getColumnLetter(ByVal lngColumn As Long) As String

    Dim strArr() As String    'to store result from split function
    
    'if we use 'true' as .Address param, then '$' symbol will be not displayed
    'in cell address. We need '$' symbol before row number, so we can split
    'string in to parts. One will contain column letter, other row number.
    strArr() = Split(ActiveSheet.Cells(1, lngColumn).Address(True, False), "$")
    
    'column letter will allways be firs element in array
    getColumnLetter = strArr(0)

End Function
