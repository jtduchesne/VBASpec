Attribute VB_Name = "modUtils"
Option Explicit

Sub Assign(Variable, ByRef Into)
    If IsObject(Variable) Then
        Set Into = Variable
    Else
        Into = Variable
    End If
End Sub
Sub AssignProperty(Property As String, Of, ByRef Into)
    Assign GetProperty(Property, Of), Into:=Into
End Sub
Function GetProperty(ByVal Property As String, Of)
    On Error GoTo TryAsMethod
    Assign CallByName(Of, Property, VbGet), Into:=GetProperty
    Exit Function
TryAsMethod:
    Assign CallByName(Of, Property, VbMethod), Into:=GetProperty
    Err.Clear
End Function
Function HasProperty(Object, ByVal Property As String) As Boolean
    If Not IsObject(Object) Then Exit Function
    On Error GoTo TryAsMethod
    Dim Result
    Assign CallByName(Object, Property, VbGet), Into:=Result
    HasProperty = True
    Exit Function
TryAsMethod:
    On Error GoTo -1
    On Error GoTo ExitFunction
    Assign CallByName(Object, Property, VbMethod), Into:=Result
    HasProperty = True
ExitFunction:
    Err.Clear
End Function

Function AnyValueAsText(Value As Variant, Optional bShowAddress As Boolean = False, _
                                          Optional bShowType As Boolean = False) As String
    If VBA.IsMissing(Value) Then
        AnyValueAsText = "(Missing)"
    ElseIf VBA.IsObject(Value) Then
        AnyValueAsText = TypeName(Value)
        If Not Value Is Nothing Then
            If IsEnumerable(Value) Then
                Dim lCount As Long
                lCount = CountEnumerable(Value)
                Select Case lCount
                Case 0:    AnyValueAsText = AnyValueAsText & "(Empty)"
                Case 1:    AnyValueAsText = AnyValueAsText & "(1 item)"
                Case Else: AnyValueAsText = AnyValueAsText & "(" & lCount & " items)"
                End Select
            Else
                If bShowAddress Then AnyValueAsText = AnyValueAsText & "[" & ObjPtr(Value) & "]"
            End If
        End If
    ElseIf VBA.IsEmpty(Value) Or VBA.IsNull(Value) Then
        AnyValueAsText = TypeName(Value)
    ElseIf IsEnumerable(Value) Then
        AnyValueAsText = "Array/" & TypeName(Value)
    Else
        If VarType(Value) = vbString Then
            AnyValueAsText = """" & Value & """"
            If bShowAddress Then AnyValueAsText = AnyValueAsText & "[" & VarPtr(Value) & "]"
        ElseIf VarType(Value) = vbDouble Then
            AnyValueAsText = Format(Value, "#.0#########")
        Else
            AnyValueAsText = CStr(Value)
        End If
        If bShowType Then AnyValueAsText = AnyValueAsText & "(" & TypeName(Value) & ")"
    End If
End Function

Function ElapsedTimeToString(dElapsedTime As Double) As String
    If dElapsedTime >= 60 Then
        ElapsedTimeToString = Int(dElapsedTime / 60) & "m" & Right(Strings.Format(dElapsedTime Mod 60, "00.000"), 6)
    Else
        ElapsedTimeToString = Strings.Format(dElapsedTime, "#0.000")
    End If
End Function

Function IsBoolean(Value As Variant) As Boolean
    IsBoolean = (VarType(Value) = vbBoolean)
End Function
Function IsString(Value As Variant) As Boolean
    IsString = (VarType(Value) = vbString)
End Function

Function IsEnumerable(Value As Variant) As Boolean
    Select Case VarType(Value)
    Case vbArray To vbArray + vbByte
        IsEnumerable = True
    Case vbObject
        If TypeOf Value Is Collection Then
            IsEnumerable = True
        End If
    End Select
End Function
Function CountEnumerable(Value As Variant) As Long
    Select Case VarType(Value)
    Case vbArray To vbArray + vbByte
        CountEnumerable = UBound(Value) - LBound(Value)
    Case vbObject
        If TypeOf Value Is Collection Then
            CountEnumerable = Value.Count
        End If
    End Select
End Function
