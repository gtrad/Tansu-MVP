VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VariablePickerForm
   Caption         =   "Insert Tansu Variable"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "VariablePickerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VariablePickerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Tansu Variable Picker Form
' Displays a list of variables from Tansu for insertion into Word

Option Explicit

Private m_Variables As Collection
Private m_SelectedIndex As Long

Private Sub UserForm_Initialize()
    ' Set up the form
    Me.Caption = "Insert Tansu Variable"

    ' Load variables from API
    LoadVariables

    m_SelectedIndex = -1
End Sub

Private Sub LoadVariables()
    ' Clear existing items
    lstVariables.Clear

    ' Get variables from API
    Set m_Variables = GetVariables()

    ' Populate list
    Dim varData As Variant
    Dim displayText As String

    For Each varData In m_Variables
        displayText = varData(1) & ": " & varData(2)
        If Len(varData(3)) > 0 Then
            displayText = displayText & " " & varData(3)
        End If
        lstVariables.AddItem displayText
    Next varData

    ' Update status
    If m_Variables.Count = 0 Then
        lblStatus.Caption = "No variables found. Is Tansu running?"
    Else
        lblStatus.Caption = m_Variables.Count & " variable(s)"
    End If
End Sub

Private Sub txtSearch_Change()
    ' Filter variables based on search text
    Dim searchText As String
    searchText = LCase(txtSearch.Text)

    lstVariables.Clear

    Dim varData As Variant
    Dim displayText As String
    Dim i As Long

    i = 0
    For Each varData In m_Variables
        displayText = varData(1) & ": " & varData(2)
        If Len(varData(3)) > 0 Then
            displayText = displayText & " " & varData(3)
        End If

        ' Show if matches search or search is empty
        If Len(searchText) = 0 Or InStr(LCase(displayText), searchText) > 0 Then
            lstVariables.AddItem displayText
        End If
    Next varData
End Sub

Private Sub lstVariables_Click()
    m_SelectedIndex = lstVariables.ListIndex
End Sub

Private Sub lstVariables_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click to insert as field
    InsertAsField
End Sub

Private Sub btnInsertField_Click()
    InsertAsField
End Sub

Private Sub btnInsertText_Click()
    InsertAsText
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRefresh_Click()
    LoadVariables
    txtSearch.Text = ""
End Sub

Private Sub InsertAsField()
    If lstVariables.ListIndex < 0 Then
        MsgBox "Please select a variable first.", vbInformation, "Tansu"
        Exit Sub
    End If

    ' Find the selected variable in our collection
    Dim selectedText As String
    selectedText = lstVariables.List(lstVariables.ListIndex)

    Dim varData As Variant
    For Each varData In m_Variables
        Dim displayText As String
        displayText = varData(1) & ": " & varData(2)
        If Len(varData(3)) > 0 Then
            displayText = displayText & " " & varData(3)
        End If

        If displayText = selectedText Then
            ' Insert with or without unit based on checkbox
            Dim valueToInsert As String
            If chkWithUnit.Value And Len(varData(3)) > 0 Then
                valueToInsert = varData(2) & " " & varData(3)
            Else
                valueToInsert = varData(2)
            End If

            InsertVariableField varData(1), valueToInsert
            Unload Me
            Exit Sub
        End If
    Next varData
End Sub

Private Sub InsertAsText()
    If lstVariables.ListIndex < 0 Then
        MsgBox "Please select a variable first.", vbInformation, "Tansu"
        Exit Sub
    End If

    ' Find the selected variable in our collection
    Dim selectedText As String
    selectedText = lstVariables.List(lstVariables.ListIndex)

    Dim varData As Variant
    For Each varData In m_Variables
        Dim displayText As String
        displayText = varData(1) & ": " & varData(2)
        If Len(varData(3)) > 0 Then
            displayText = displayText & " " & varData(3)
        End If

        If displayText = selectedText Then
            ' Insert with or without unit based on checkbox
            Dim valueToInsert As String
            If chkWithUnit.Value And Len(varData(3)) > 0 Then
                valueToInsert = varData(2) & " " & varData(3)
            Else
                valueToInsert = varData(2)
            End If

            InsertVariableText valueToInsert
            Unload Me
            Exit Sub
        End If
    Next varData
End Sub
