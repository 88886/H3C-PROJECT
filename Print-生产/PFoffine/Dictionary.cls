VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private keys() As String
Private values() As String
Private count As Integer

Sub Class_Initialize()
    count = 0
    ReDim keys(0) As String
    ReDim values(0) As String
End Sub

Sub Class_Terminate()
    Erase keys
    Erase values
End Sub

Public Sub PutValue(ByVal Key As String, ByVal value As String)
    Dim index As Integer
    index = GetIndex(Key)
    If (index > -1) Then
        values(index) = value
    Else
        count = count + 1
        ReDim Preserve keys(count) As String
        ReDim Preserve values(count) As String
        keys(count) = Key
        values(count) = value
    End If
    
End Sub

Public Function GetValue(ByVal Key As String) As String
    Dim index As Integer
    index = GetIndex(Key)
    If (index = -1) Then
        GetValue = "Null"
        Exit Function
    End If
    GetValue = values(index)
End Function

Public Function GetIndex(ByVal Key As String) As Integer
    For i = LBound(keys) To UBound(keys)
        If (Key = keys(i)) Then
            GetIndex = i
            Exit Function
        End If
    Next
    GetIndex = -1
End Function
