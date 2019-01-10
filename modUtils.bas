Attribute VB_Name = "modUtils"
Option Explicit

Function AnyValueAsText(Value As Variant, Optional bShowAddress As Boolean = False, _
                                          Optional bShowType As Boolean = False) As String
    If VBA.IsMissing(Value) Then
        AnyValueAsText = "(Missing)"
    ElseIf VBA.IsObject(Value) Then
        If Value Is Nothing Then
            AnyValueAsText = TypeName(Value)
        Else
            AnyValueAsText = TypeName(Value)
            If bShowAddress Then AnyValueAsText = AnyValueAsText & "[" & ObjPtr(Value) & "]"
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
    If IsEnumerable(Value) Then
        Dim lCount As Long
        lCount = CountEnumerable(Value)
        Select Case lCount
        Case 0:    AnyValueAsText = AnyValueAsText & "(Empty)"
        Case 1:    AnyValueAsText = AnyValueAsText & "(1 item)"
        Case Else: AnyValueAsText = AnyValueAsText & "(" & lCount & " items)"
        End Select
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
