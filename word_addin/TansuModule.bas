Attribute VB_Name = "TansuModule"
' Tansu Variable Tracker - Word VBA Module
' This module provides functions to insert variables from Tansu into Word documents.
' Requires the Tansu API server to be running (start tray_app.py or api_server.py)

Option Explicit

Private Const API_URL As String = "http://127.0.0.1:5050"

' Main entry point - shows the variable picker dialog
Public Sub InsertTansuVariable()
    ' Check if API is running
    If Not IsAPIRunning() Then
        MsgBox "Tansu is not running." & vbCrLf & vbCrLf & _
               "Please start the Tansu tray app first.", vbExclamation, "Tansu"
        Exit Sub
    End If

    ' Show the variable picker form
    Dim frm As New VariablePickerForm
    frm.Show
End Sub

' Check if the Tansu API server is running
Public Function IsAPIRunning() As Boolean
    On Error GoTo NotRunning

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", API_URL & "/ping", False
    http.Send

    IsAPIRunning = (http.Status = 200)
    Exit Function

NotRunning:
    IsAPIRunning = False
End Function

' Get all variables from the Tansu API
Public Function GetVariables() As Collection
    On Error GoTo ErrorHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", API_URL & "/variables", False
    http.Send

    If http.Status <> 200 Then
        Set GetVariables = New Collection
        Exit Function
    End If

    ' Parse JSON response
    Dim json As String
    json = http.responseText

    Set GetVariables = ParseVariablesJSON(json)
    Exit Function

ErrorHandler:
    Set GetVariables = New Collection
End Function

' Simple JSON parser for variables array
' Expected format: {"variables": [{"id": 1, "name": "var1", "value": "100", "unit": "USD"}, ...]}
Private Function ParseVariablesJSON(json As String) As Collection
    Dim result As New Collection
    Dim varsStart As Long, varsEnd As Long
    Dim currentPos As Long
    Dim varObj As Object

    ' Find the variables array
    varsStart = InStr(json, """variables""")
    If varsStart = 0 Then
        Set ParseVariablesJSON = result
        Exit Function
    End If

    varsStart = InStr(varsStart, json, "[")
    varsEnd = InStr(varsStart, json, "]")

    If varsStart = 0 Or varsEnd = 0 Then
        Set ParseVariablesJSON = result
        Exit Function
    End If

    ' Extract each variable object
    currentPos = varsStart
    Do
        Dim objStart As Long, objEnd As Long
        objStart = InStr(currentPos, json, "{")

        If objStart = 0 Or objStart > varsEnd Then Exit Do

        objEnd = InStr(objStart, json, "}")
        If objEnd = 0 Then Exit Do

        Dim objStr As String
        objStr = Mid(json, objStart, objEnd - objStart + 1)

        ' Parse individual variable
        Dim varName As String, varValue As String, varUnit As String, varId As String
        varId = ExtractJSONValue(objStr, "id")
        varName = ExtractJSONValue(objStr, "name")
        varValue = ExtractJSONValue(objStr, "value")
        varUnit = ExtractJSONValue(objStr, "unit")

        ' Create a simple array to store variable data
        Dim varData(3) As String
        varData(0) = varId
        varData(1) = varName
        varData(2) = varValue
        varData(3) = varUnit

        result.Add varData

        currentPos = objEnd + 1
    Loop

    Set ParseVariablesJSON = result
End Function

' Extract a value from a simple JSON object
Private Function ExtractJSONValue(json As String, key As String) As String
    Dim keyPos As Long, valueStart As Long, valueEnd As Long
    Dim searchKey As String

    searchKey = """" & key & """"
    keyPos = InStr(json, searchKey)

    If keyPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    ' Find the colon after the key
    valueStart = InStr(keyPos, json, ":")
    If valueStart = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    valueStart = valueStart + 1

    ' Skip whitespace
    Do While Mid(json, valueStart, 1) = " "
        valueStart = valueStart + 1
    Loop

    ' Check if value is a string (starts with quote) or number
    If Mid(json, valueStart, 1) = """" Then
        valueStart = valueStart + 1
        valueEnd = InStr(valueStart, json, """")
        ExtractJSONValue = Mid(json, valueStart, valueEnd - valueStart)
    Else
        ' Number - find end (comma, brace, or bracket)
        valueEnd = valueStart
        Do While valueEnd <= Len(json)
            Dim c As String
            c = Mid(json, valueEnd, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            valueEnd = valueEnd + 1
        Loop
        ExtractJSONValue = Trim(Mid(json, valueStart, valueEnd - valueStart))
    End If
End Function

' Insert a variable as a DOCVARIABLE field at the current cursor position
Public Sub InsertVariableField(varName As String, varValue As String)
    On Error GoTo ErrorHandler

    Dim doc As Document
    Set doc = ActiveDocument

    ' Set the document variable
    On Error Resume Next
    doc.Variables(varName).Delete
    On Error GoTo ErrorHandler
    doc.Variables.Add varName, varValue

    ' Insert a DOCVARIABLE field at the cursor
    Dim rng As Range
    Set rng = Selection.Range

    Dim fld As Field
    Set fld = doc.Fields.Add(Range:=rng, Type:=wdFieldDocVariable, Text:=varName, PreserveFormatting:=True)

    ' Update the field to show the value
    fld.Update

    ' Move cursor after the field
    Selection.MoveRight Unit:=wdCharacter, Count:=1

    Exit Sub

ErrorHandler:
    MsgBox "Error inserting variable: " & Err.Description, vbExclamation, "Tansu"
End Sub

' Insert a variable as plain text (not updatable)
Public Sub InsertVariableText(varValue As String)
    Selection.TypeText Text:=varValue
End Sub
