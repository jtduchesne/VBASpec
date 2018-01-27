VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecExpectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const INTERPOLATION_SYMBOL As String = "#{"
Private Const ACTUAL_INTERPOLATION As String = "#{actual}"
Private Const EXPECTED_INTERPOLATION As String = "#{expected}"

Private Const FAILURE_MSG_PREFIX As String = "Expected "
Private Const FAILURE_MSG_CONNECTOR As String = " to be "
Private Const FAILURE_MSG_NEGATED_CONNECTOR As String = " to not be "

Private vActual As Variant
Private vExpected As Variant

Private bPassed As Boolean
Private bFailed As Boolean
Private sFailureMessage As String

Private iSignificantFigures As Integer

'========================================================================================================================================'

Friend Property Get Actual() As Variant
    If IsObject(vActual) Then
        Set Actual = vActual
    Else
        Actual = vActual
    End If
End Property
Friend Property Let Actual(Value As Variant)
    vActual = Value
End Property
Friend Property Set Actual(Value As Variant)
    Set vActual = Value
End Property

Friend Property Get Expected() As Variant
    If IsObject(vExpected) Then
        Set Expected = vExpected
    Else
        Expected = vExpected
    End If
End Property
Friend Property Let Expected(Value As Variant)
    vExpected = Value
End Property
Friend Property Set Expected(Value As Variant)
    Set vExpected = Value
End Property

Friend Property Get Passed() As Boolean
    Passed = bPassed
End Property
Friend Property Get Failed() As Boolean
    Failed = bFailed
End Property

Friend Property Get FailureMessage() As String
    If bFailed Then
        If Len(sFailureMessage) > 0 Then
            FailureMessage = sFailureMessage
        Else
            FailureMessage = Err.Description
        End If
    End If
End Property

'========================================================================================================================================'

Public Sub ToEqual(Expected As Variant)
    VerifyActual "IsEqual", Expected, " to equal #{expected}"
End Sub
Public Sub ToNotEqual(Expected As Variant)
    VerifyActual "IsEqual", Expected, " to not equal #{expected}", Inverse:=True
End Sub
Public Function IsEqual(Actual As Variant, Expected As Variant) As Boolean
    If IsObject(Actual) Or IsObject(Expected) Then
        IsEqual = False
    ElseIf VarType(Actual) = vbDouble And VarType(Expected) = vbDouble Then
        IsEqual = IsCloseTo(Actual, Expected, 15)
    Else
        IsEqual = (Actual = Expected)
    End If
End Function

'========================================================================================================================================'

Public Sub ToBe(Expected As Variant)
    VerifyActual "IsSame", Expected
End Sub
Public Sub ToNotBe(Expected As Variant)
    VerifyActual "IsSame", Expected, Inverse:=True
End Sub
Public Function IsSame(Actual As Variant, Expected As Variant) As Boolean
    If IsObject(Actual) And IsObject(Expected) Then
        IsSame = (Actual Is Expected)
    Else
        IsSame = False
    End If
End Function

'========================================================================================================================================'

Public Sub ToBeA(Expected As Variant)
    VerifyActual "IsA", Expected, " to be a " & CStr(Expected)
End Sub
Public Sub ToBeAn(Expected As Variant)
    VerifyActual "IsA", Expected, " to be an " & CStr(Expected)
End Sub
Public Sub ToNotBeA(Expected As Variant)
    VerifyActual "IsA", Expected, " to not be a " & CStr(Expected), Inverse:=True
End Sub
Public Sub ToNotBeAn(Expected As Variant)
    VerifyActual "IsA", Expected, " to not be an " & CStr(Expected), Inverse:=True
End Sub
Public Function IsA(Actual As Variant, Expected As Variant) As Boolean
    IsA = (TypeName(Actual) = Expected)
End Function

'========================================================================================================================================'

Public Sub ToBeEmpty()
    VerifyActual "IsEmpty", Empty, " to be empty"
End Sub
Public Sub ToNotBeEmpty()
    VerifyActual "IsEmpty", Empty, " to not be empty", Inverse:=True
End Sub
Public Function IsEmpty(Actual As Variant, Optional Expected As Variant) As Boolean
    If IsString(Actual) Then
        IsEmpty = (Len(Trim$(Actual)) = 0)
    Else
        IsEmpty = IsNothing(Actual) Or VBA.IsEmpty(Actual) Or VBA.IsNull(Actual) Or VBA.IsMissing(Actual)
    End If
End Function

'========================================================================================================================================'

Public Sub ToBeNothing()
    VerifyActual "IsNothing", Nothing
End Sub
Public Sub ToNotBeNothing()
    VerifyActual "IsNothing", Nothing, Inverse:=True
End Sub
Public Function IsNothing(Actual As Variant, Optional Expected As Variant) As Boolean
    If IsObject(Actual) Then
        IsNothing = (Actual Is Nothing)
    Else
        IsNothing = False
    End If
End Function

'========================================================================================================================================'

Public Sub ToBeFalsy()
    VerifyActual "IsFalsy", False, " to evaluate to #{expected}"
End Sub
Public Function IsFalsy(Actual As Variant, Optional Expected As Variant) As Boolean
    IsFalsy = IIf(Actual, False, True)
End Function

'========================================================================================================================================'

Public Sub ToBeFalse()
    VerifyActual "IsFalse", False
End Sub
Public Function IsFalse(Actual As Variant, Optional Expected As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsFalse = (Actual = False)
    Else
        IsFalse = False
    End If
End Function

'========================================================================================================================================'

Public Sub ToBeTruthy()
    VerifyActual "IsTruthy", True, " to evaluate to #{expected}"
End Sub
Public Function IsTruthy(Actual As Variant, Optional Expected As Variant) As Boolean
    IsTruthy = IIf(Actual, True, False)
End Function

'========================================================================================================================================'

Public Sub ToBeTrue()
    VerifyActual "IsTrue", True
End Sub
Public Function IsTrue(Actual As Variant, Optional Expected As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsTrue = (Actual = True)
    Else
        IsTrue = False
    End If
End Function

'========================================================================================================================================'

Public Sub ToBeLessThan(Expected As Variant)
    VerifyActual "IsLT", Expected, " to be less than #{expected}"
End Sub
Public Sub ToBeLT(Expected As Variant)
    VerifyActual "IsLT", Expected, " to be < #{expected}"
End Sub
Public Function IsLT(Actual As Variant, Expected As Variant) As Boolean
    IsLT = (Actual < Expected)
End Function

'========================================================================================================================================'

Public Sub ToBeLessThanOrEqualTo(Expected As Variant)
    VerifyActual "IsLTE", Expected, " to be less than or equal to #{expected}"
End Sub
Public Sub ToBeLTE(Expected As Variant)
    VerifyActual "IsLTE", Expected, " to be <= #{expected}"
End Sub
Public Function IsLTE(Actual As Variant, Expected As Variant) As Boolean
    IsLTE = (Actual <= Expected)
End Function

'========================================================================================================================================'

Public Sub ToBeGreaterThan(Expected As Variant)
    VerifyActual "IsGT", Expected, " to be greater than #{expected}"
End Sub
Public Sub ToBeGT(Expected As Variant)
    VerifyActual "IsGT", Expected, " to be > #{expected}"
End Sub
Public Function IsGT(Actual As Variant, Expected As Variant) As Boolean
    IsGT = (Actual > Expected)
End Function

'========================================================================================================================================'

Public Sub ToBeGreaterThanOrEqualTo(Expected As Variant)
    VerifyActual "IsGTE", Expected, " to be greater than or equal to #{expected}"
End Sub
Public Sub ToBeGTE(Expected As Variant)
    VerifyActual "IsGTE", Expected, " to be >= #{expected}"
End Sub
Public Function IsGTE(Actual As Variant, Expected As Variant) As Boolean
    IsGTE = (Actual >= Expected)
End Function

'========================================================================================================================================'

Public Sub ToBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2)
    iSignificantFigures = SignificantFigures
    VerifyActual "IsCloseTo", Expected, " to be close to #{expected} (" & SignificantFigures & " significant figures)"
End Sub
Public Sub ToNotBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2)
    iSignificantFigures = SignificantFigures
    VerifyActual "IsCloseTo", Expected, " to not be close to #{expected} (" & SignificantFigures & " significant figures)", Inverse:=True
End Sub
Public Function IsCloseTo(Actual As Variant, Expected As Variant, Optional SignificantFigures As Integer = -1) As Boolean
    Dim sActual As String
    Dim sExpected As String
    If SignificantFigures < 1 Then SignificantFigures = iSignificantFigures
    sActual = Format$(Actual, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Actual >= 1, "e+0", "e-0"))
    sExpected = Format$(Expected, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Expected >= 1, "e+0", "e-0"))
    
    IsCloseTo = (StrComp(sActual, sExpected, vbBinaryCompare) = 0)
End Function

'========================================================================================================================================'

Public Sub ToInclude(Expected As Variant)
    VerifyActual "Includes", Expected, " to include #{expected}"
End Sub
Public Sub ToNotInclude(Expected As Variant)
    VerifyActual "Includes", Expected, " to not include #{expected}", Inverse:=True
End Sub
Public Function Includes(Actual As Variant, Expected As Variant) As Boolean
    If IsString(Actual) Then
        Includes = InStr(1, Actual, Expected, vbTextCompare) > 0
    ElseIf IsEnumerable(Actual) Then
        Dim i As Integer
        If TypeOf Actual Is Collection Then
            For i = 1 To Actual.Count
                If IsLike(Actual.Item(i), Expected) Then
                    Includes = True
                    Exit Function
                End If
            Next i
        Else
            For i = LBound(Actual) To UBound(Actual)
                If IsLike(Actual(i), Expected) Then
                    Includes = True
                    Exit Function
                End If
            Next i
        End If
    Else
        Includes = False
    End If
End Function

'========================================================================================================================================'

Private Function IsLike(Var1 As Variant, Var2 As Variant) As Boolean
    If IsObject(Var1) Then
        IsLike = IsSame(Var1, Var2)
    Else
        IsLike = IsEqual(Var1, Var2)
    End If
End Function

Private Function IsString(Value As Variant) As Boolean
    IsString = (VarType(Value) = vbString)
End Function

Private Function IsEnumerable(Value As Variant) As Boolean
    Select Case VarType(Value)
    Case vbArray To vbArray + vbByte
        IsEnumerable = True
    Case vbObject
        If TypeOf Value Is Collection Then
            IsEnumerable = True
        End If
    End Select
End Function

'========================================================================================================================================'

Private Sub VerifyActual(ProcName As String, Expected As Variant, Optional sMessage As String = vbNullString, _
                                                                  Optional Inverse As Boolean = False)
    If IsObject(Expected) Then
        Set Me.Expected = Expected
    Else
        Me.Expected = Expected
    End If
    
    If Inverse Then
        bFailed = CallByName(Me, ProcName, VbMethod, Me.Actual, Expected)
        bPassed = Not bFailed
    Else
        bPassed = CallByName(Me, ProcName, VbMethod, Me.Actual, Expected)
        bFailed = Not bPassed
    End If
    
    If bFailed Then
        If Len(sMessage) > 0 Then
            If InStr(1, sMessage, INTERPOLATION_SYMBOL, vbBinaryCompare) > 0 Then
                If InStr(1, sMessage, ACTUAL_INTERPOLATION, vbBinaryCompare) > 1 Then
                    sMessage = Replace$(sMessage, ACTUAL_INTERPOLATION, AnyValueAsText(Me.Actual), Compare:=vbBinaryCompare)
                Else
                    sMessage = AnyValueAsText(Me.Actual) & sMessage
                End If
                If InStr(1, sMessage, EXPECTED_INTERPOLATION, vbBinaryCompare) > 0 Then
                    sMessage = Replace$(sMessage, EXPECTED_INTERPOLATION, AnyValueAsText(Expected), Compare:=vbBinaryCompare)
                End If
                sFailureMessage = FAILURE_MSG_PREFIX & sMessage
            Else
                sFailureMessage = FAILURE_MSG_PREFIX & AnyValueAsText(Me.Actual) & sMessage
            End If
        Else
            sFailureMessage = FAILURE_MSG_PREFIX & AnyValueAsText(Me.Actual) & _
                              IIf(Inverse, FAILURE_MSG_NEGATED_CONNECTOR, FAILURE_MSG_CONNECTOR) & _
                              AnyValueAsText(Expected)
        End If
    End If
End Sub

Private Function AnyValueAsText(Value As Variant, Optional bShowAddress As Boolean = False) As String
    If IsMissing(Value) Then
        AnyValueAsText = "(Missing)"
    ElseIf IsEmpty(Value) Then
        AnyValueAsText = TypeName(Value)
    Else
        Select Case VarType(Value)
        Case vbObject
            AnyValueAsText = TypeName(Value)
            If bShowAddress Then AnyValueAsText = AnyValueAsText & "[" & ObjPtr(Value) & "]"
        Case vbArray To vbArray + vbByte
            AnyValueAsText = "Array/" & TypeName(Value)
        Case vbString
            AnyValueAsText = """" & Value & """"
            If bShowAddress Then AnyValueAsText = AnyValueAsText & "[" & VarPtr(Value) & "]"
        Case Else
            AnyValueAsText = CStr(Value)
        End Select
    End If
End Function
