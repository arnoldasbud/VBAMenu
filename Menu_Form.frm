VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Form 
   Caption         =   "Meniu"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   OleObjectBlob   =   "Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu_Form"
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
    prepareCableCoresBox
End Sub

'**************************************************************************************************
'* Event that fires up when user changes combobox selection
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub comboBoxCableCores_Change()
    prepareCableCrossBox
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'add cable' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCableAdd_Click()
    prepareCableListBox
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'delete cable' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCableRemove_Click()
    listBoxRemoveCalbe
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'search' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonSearchGlands_Click()
    prepareResultArray
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'Range' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCombineRange_Click()
    addSelectedRange
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'Code' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCombineCode_Click()
    addSelectedCodeColumn
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'Value' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCombineValue_Click()
    addSelectedValueColumn
End Sub

'**************************************************************************************************
'* Event that fires up when user clicks 'Combine selected' button
'*
'* Params: -
'* Return: -
'**************************************************************************************************
Private Sub buttonCombineSelectedRange_Click()
    prepareCombineResultArray
End Sub

