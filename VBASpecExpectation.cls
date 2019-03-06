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
Private Const RESULT_INTERPOLATION As String = "#{result}"
Private Const TRUE_INTERPOLATION As String = "#{true}"
Private Const FALSE_INTERPOLATION As String = "#{false}"

Private Const FAILURE_MSG As String = "Expected #{expected}, got #{actual}"
Private Const FAILURE_MSG_NEGATED As String = "#{actual} not expected to be #{expected}"
Private Const WARNING_MSG_PREFIX As String = ">>> Warning: "
Private Const WARNING_MSG_SUFFIX As String = " <<<"

Private vActual As Variant
Private vExpected As Variant

Private vCapture As Variant
Private dictOtherCaptures As Scripting.Dictionary

Private bPassed As Boolean
Private bFailed As Boolean
Private sFailureMessage As String
Private sWarningMessage As String

Private vAdditionalArgument As Variant
Private bErrorChecked As Boolean

'========================================================================================================================================'

Friend Property Get Actual()
    Assign vActual, Into:=Actual
End Property
Friend Property Let Actual(Value As Variant)
    vActual = Value
End Property
Friend Property Set Actual(Value As Variant)
    Set vActual = Value
End Property

Friend Property Get Expected()
    Assign vExpected, Into:=Expected
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
        FailureMessage = sFailureMessage
    End If
End Property

Friend Function Fail(sMessage As String) As VBASpecExpectation
    bPassed = False
    bFailed = True
    sFailureMessage = sMessage
    
    Set Fail = Me
End Function

Public Function Warn(sMessage As String) As VBASpecExpectation
    sWarningMessage = WARNING_MSG_PREFIX & sMessage & WARNING_MSG_SUFFIX
    
    Set Warn = Me
End Function

'========================================================================================================================================'

Public Function WithCapture(Capture As Variant, Optional OtherCaptures As Scripting.Dictionary) As VBASpecExpectation
    Assign Capture, Into:=vCapture
    Set dictOtherCaptures = OtherCaptures
    
    Set WithCapture = Me
End Function

Public Function AndAlso() As VBASpecExpectation
    Set AndAlso = Me
End Function
Public Function ButAlso() As VBASpecExpectation
    Set ButAlso = Me
End Function

'========================================================================================================================================'

Public Function ToEqual(Expected As Variant) As VBASpecExpectation
    Set ToEqual = VerifyActual("IsEqual", Expected)
End Function
Public Function ToNotEqual(Expected As Variant) As VBASpecExpectation
    Set ToNotEqual = VerifyActual("IsEqual", Expected, Inverse:=True)
End Function
Public Function IsEqual(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsObject(Actual) Or IsObject(Expected) Then
        IsEqual = False
    ElseIf VarType(Actual) = vbString And VarType(Expected) = vbString Then
        IsEqual = (StrComp(Actual, Expected, vbTextCompare) = 0)
    ElseIf (VarType(Actual) = vbDouble Or VarType(Expected) = vbDouble) Then
        IsEqual = IsCloseTo(Actual, Expected, Result, 15)
    Else
        IsEqual = (Actual = Expected)
    End If
End Function

'========================================================================================================================================'

Public Function ToBe(Expected As Variant) As VBASpecExpectation
    If VarType(Expected) = vbString Then
        If HasProperty(vActual, Expected) Then
            Set ToBe = VerifyActual("ReturnsTrue", CStr(Expected), "Expected #{actual}#" & Expected & " to be #{true}" & _
                                                                   ", but is #{result}")
            Exit Function
        ElseIf HasProperty(vActual, "Is" & Expected) Then
            Set ToBe = VerifyActual("ReturnsTrue", "Is" & Expected, "Expected #{actual}#Is" & Expected & " to be #{true}" & _
                                                                    ", but is #{result}")
            Exit Function
        End If
    End If
    Set ToBe = VerifyActual("IsSame", Expected)
End Function
Public Function ToNotBe(Expected As Variant) As VBASpecExpectation
    If VarType(Expected) = vbString Then
        If HasProperty(vActual, Expected) Then
            Set ToNotBe = VerifyActual("ReturnsTrue", CStr(Expected), "Expected #{actual}#" & Expected & " to be #{false}" & _
                                                                      ", but is #{result}", Inverse:=True)
            Exit Function
        ElseIf HasProperty(vActual, "Is" & Expected) Then
            Set ToNotBe = VerifyActual("ReturnsTrue", "Is" & Expected, "Expected #{actual}#Is" & Expected & " to be #{false}" & _
                                                                       ", but is #{result}", Inverse:=True)
            Exit Function
        End If
    End If
    Set ToNotBe = VerifyActual("IsSame", Expected, Inverse:=True)
End Function
Public Function IsSame(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsObject(Actual) And IsObject(Expected) Then
        IsSame = (Actual Is Expected)
    Else
        IsSame = False
    End If
End Function

'========================================================================================================================================'

Public Function ToHave(Expected As Variant) As VBASpecExpectation
    Set ToHave = VerifyActual("ReturnsTrue", "Has" & Expected, "Expected #{actual}#Has" & Expected & " to be #{true}" & _
                                                               ", but is #{result}")
End Function
Public Function ToNotHave(Expected As Variant) As VBASpecExpectation
    Set ToNotHave = VerifyActual("ReturnsTrue", "Has" & Expected, "Expected #{actual}#Has" & Expected & " to be #{false}" & _
                                                                  ", but is #{result}", Inverse:=True)
End Function
Public Function ReturnsTrue(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    On Error Resume Next
    AssignProperty CStr(Expected), Of:=Actual, Into:=Result
    ReturnsTrue = IsTrue(Result)
    Err.Clear
End Function

'========================================================================================================================================'

Public Function ToBeA(Expected As Variant) As VBASpecExpectation
    Set ToBeA = VerifyActual("IsA", Expected, "Expected #{actual} to be a " & CStr(Expected) & ", but is a #{result}")
End Function
Public Function ToBeAn(Expected As Variant) As VBASpecExpectation
    Set ToBeAn = VerifyActual("IsA", Expected, "Expected #{actual} to be an " & CStr(Expected) & ", but is a #{result}")
End Function
Public Function ToNotBeA(Expected As Variant) As VBASpecExpectation
    Set ToNotBeA = VerifyActual("IsA", Expected, "Expected #{actual} not to be a " & CStr(Expected) & ", but it is", Inverse:=True)
End Function
Public Function ToNotBeAn(Expected As Variant) As VBASpecExpectation
    Set ToNotBeAn = VerifyActual("IsA", Expected, "Expected #{actual} not to be an " & CStr(Expected) & ", but it is", Inverse:=True)
End Function
Public Function IsA(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    Result = TypeName(Actual)
    IsA = (Result = Expected)
End Function

'========================================================================================================================================'

Public Function ToBeEmpty() As VBASpecExpectation
    Set ToBeEmpty = VerifyActual("IsEmpty", Empty, "#{actual} expected to be #{expected}")
End Function
Public Function ToNotBeEmpty() As VBASpecExpectation
   Set ToNotBeEmpty = VerifyActual("IsEmpty", Empty, Inverse:=True)
End Function
Public Function IsEmpty(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsString(Actual) Then
        IsEmpty = (Len(Trim$(Actual)) = 0)
    Else
        IsEmpty = IsNothing(Actual) Or VBA.IsEmpty(Actual) Or VBA.IsNull(Actual) Or VBA.IsMissing(Actual)
    End If
End Function

'========================================================================================================================================'

Public Function ToBeNothing() As VBASpecExpectation
    Set ToBeNothing = VerifyActual("IsNothing", Nothing)
End Function
Public Function ToNotBeNothing() As VBASpecExpectation
    Set ToNotBeNothing = VerifyActual("IsNothing", Nothing, Inverse:=True)
End Function
Public Function IsNothing(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsObject(Actual) Then
        IsNothing = (Actual Is Nothing)
    Else
        IsNothing = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeZero() As VBASpecExpectation
    Set ToBeZero = VerifyActual("IsZero", 0)
End Function
Public Function ToNotBeZero() As VBASpecExpectation
    Set ToNotBeZero = VerifyActual("IsZero", 0, Inverse:=True)
End Function
Public Function IsZero(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsZero = IsEqual(Actual, 0)
End Function

'========================================================================================================================================'

Public Function ToBeFalsy() As VBASpecExpectation
    Set ToBeFalsy = VerifyActual("IsFalsy", False, "Expected #{actual} to evaluate to #{expected}")
End Function
Public Function IsFalsy(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsFalsy = IIf(Actual, False, True)
End Function

'========================================================================================================================================'

Public Function ToBeFalse() As VBASpecExpectation
    Set ToBeFalse = VerifyActual("IsFalse", False)
End Function
Public Function IsFalse(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsFalse = (Actual = False)
    Else
        IsFalse = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeTruthy() As VBASpecExpectation
    Set ToBeTruthy = VerifyActual("IsTruthy", True, "Expected #{actual} to evaluate to #{expected}")
End Function
Public Function IsTruthy(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsTruthy = IIf(Actual, True, False)
End Function

'========================================================================================================================================'

Public Function ToBeTrue() As VBASpecExpectation
    Set ToBeTrue = VerifyActual("IsTrue", True)
End Function
Public Function IsTrue(Actual As Variant, Optional Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsTrue = (Actual = True)
    Else
        IsTrue = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeLessThan(Expected As Variant) As VBASpecExpectation
    Set ToBeLessThan = VerifyActual("IsLT", Expected, "Expected #{actual} to be less than #{expected}")
End Function
Public Function ToBeLT(Expected As Variant) As VBASpecExpectation
    Set ToBeLT = VerifyActual("IsLT", Expected, "Expected #{actual} to be < #{expected}")
End Function
Public Function IsLT(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsLT = (Actual < Expected)
End Function

'========================================================================================================================================'

Public Function ToBeLessThanOrEqualTo(Expected As Variant) As VBASpecExpectation
    Set ToBeLessThanOrEqualTo = VerifyActual("IsLTE", Expected, "Expected #{actual} to be less than or equal to #{expected}")
End Function
Public Function ToBeLTE(Expected As Variant) As VBASpecExpectation
    Set ToBeLTE = VerifyActual("IsLTE", Expected, "Expected #{actual} to be <= #{expected}")
End Function
Public Function IsLTE(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsLTE = (Actual <= Expected)
End Function

'========================================================================================================================================'

Public Function ToBeGreaterThan(Expected As Variant) As VBASpecExpectation
    Set ToBeGreaterThan = VerifyActual("IsGT", Expected, "Expected #{actual} to be greater than #{expected}")
End Function
Public Function ToBeGT(Expected As Variant) As VBASpecExpectation
    Set ToBeGT = VerifyActual("IsGT", Expected, "Expected #{actual} to be > #{expected}")
End Function
Public Function IsGT(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsGT = (Actual > Expected)
End Function

'========================================================================================================================================'

Public Function ToBeGreaterThanOrEqualTo(Expected As Variant) As VBASpecExpectation
    Set ToBeGreaterThanOrEqualTo = VerifyActual("IsGTE", Expected, "Expected #{actual} to be greater than or equal to #{expected}")
End Function
Public Function ToBeGTE(Expected As Variant) As VBASpecExpectation
    Set ToBeGTE = VerifyActual("IsGTE", Expected, "Expected #{actual} to be >= #{expected}")
End Function
Public Function IsGTE(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    IsGTE = (Actual >= Expected)
End Function

'========================================================================================================================================'

Public Function ToBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2) As VBASpecExpectation
    vAdditionalArgument = SignificantFigures
    Set ToBeCloseTo = VerifyActual("IsCloseTo", Expected, "Expected #{actual} to be close to #{expected}#{result}")
End Function
Public Function ToNotBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2) As VBASpecExpectation
    vAdditionalArgument = SignificantFigures
    Set ToNotBeCloseTo = VerifyActual("IsCloseTo", Expected, "Expected #{actual} not to be close to #{expected}#{result}", Inverse:=True)
End Function
Public Function IsCloseTo(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant, _
                                                                  Optional SignificantFigures As Integer = -1) As Boolean
    If IsNumeric(Actual) And IsNumeric(Expected) Then
        Dim sActual As String
        Dim sExpected As String
        If SignificantFigures < 1 Then SignificantFigures = vAdditionalArgument
        sActual = Format$(Actual, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Actual >= 1, "E+0", "E-0"))
        sExpected = Format$(Expected, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Expected >= 1, "E+0", "E-0"))
        
        IsCloseTo = (StrComp(sActual, sExpected, vbBinaryCompare) = 0)
        
        Dim dDiff As Double
        dDiff = Actual - Expected
        If dDiff <> 0 And Abs(dDiff) < 1 Then
            Result = " [Diff "
            If Abs(dDiff) < 0.00001 Then
                Result = Result & Format(dDiff, "+0.0#E+0;-0.0#E+0")
            Else
                Result = Result & Format(dDiff, "+0.0####;-0.0####")
            End If
            Result = Result & "]"
        End If
    Else
        IsCloseTo = False
    End If
End Function

'========================================================================================================================================'

Public Function ToInclude(Expected As Variant) As VBASpecExpectation
    Set ToInclude = VerifyActual("Includes", Expected, "Expected #{actual} to include #{expected}")
End Function
Public Function ToNotInclude(Expected As Variant) As VBASpecExpectation
    Set ToNotInclude = VerifyActual("Includes", Expected, "Expected #{actual} not to include #{expected}", Inverse:=True)
End Function
Public Function Includes(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
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

Public Function ToBeginWith(Expected As Variant) As VBASpecExpectation
    Set ToBeginWith = VerifyActual("BeginsWith", Expected, "Expected #{actual} to begin with #{expected}")
End Function
Public Function ToNotBeginWith(Expected As Variant) As VBASpecExpectation
    Set ToNotBeginWith = VerifyActual("BeginsWith", Expected, "Expected #{actual} not to begin with #{expected}", Inverse:=True)
End Function
Public Function BeginsWith(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsString(Actual) Then
        BeginsWith = (StrComp(Left(Actual, Len(CStr(Expected))), Expected, vbTextCompare) = 0)
    ElseIf IsEnumerable(Actual) Then
        If TypeOf Actual Is Collection Then
            BeginsWith = IsLike(Actual.Item(1), Expected)
        Else
            BeginsWith = IsLike(Actual(LBound(Actual)), Expected)
        End If
    Else
        BeginsWith = False
    End If
End Function

'========================================================================================================================================'

Public Function ToEndWith(Expected As Variant) As VBASpecExpectation
    Set ToEndWith = VerifyActual("EndsWith", Expected, "Expected #{actual} to end with #{expected}")
End Function
Public Function ToNotEndWith(Expected As Variant) As VBASpecExpectation
    Set ToNotEndWith = VerifyActual("EndsWith", Expected, "Expected #{actual} not to end with #{expected}", Inverse:=True)
End Function
Public Function EndsWith(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsString(Actual) Then
        EndsWith = (StrComp(Right(Actual, Len(CStr(Expected))), Expected, vbTextCompare) = 0)
    ElseIf IsEnumerable(Actual) Then
        If TypeOf Actual Is Collection Then
            EndsWith = IsLike(Actual.Item(Actual.Count), Expected)
        Else
            EndsWith = IsLike(Actual(UBound(Actual)), Expected)
        End If
    Else
        EndsWith = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBecome(Expected As Variant) As VBASpecExpectation
    Set ToBecome = ToChange.AndAlso.ToEqual(Expected)
End Function
Public Function ToNotBecome(Expected As Variant) As VBASpecExpectation
    Set ToNotBecome = ToChange.ButAlso.ToNotEqual(Expected)
End Function

Public Function ToChange(Optional by As Variant) As VBASpecExpectation
    If Not IsMissing(by) Then
        If by <> 0 Then
            vAdditionalArgument = by
            Set ToChange = VerifyActual("HasChanged", vCapture, "Expected #{expected} to become " & AnyValueAsText(vCapture + by) & _
                                                                ", but is now #{actual}#{result}", Inverse:=True)
        Else
            Set ToChange = ToNotChange
        End If
    Else
        Set ToChange = VerifyActual("HasChanged", vCapture, "Expected #{expected} to change, but it didn't")
    End If
End Function
Public Function ToNotChange(Optional by As Variant) As VBASpecExpectation
    If Not IsMissing(by) Then
        vAdditionalArgument = by
        Set ToNotChange = VerifyActual("HasChanged", vCapture, "Expected #{expected} not to become " & AnyValueAsText(vCapture + by) & _
                                                               ", but it did")
    Else
        Set ToNotChange = VerifyActual("HasChanged", vCapture, "Expected #{expected}" & _
                                                               " not to change, but is now #{actual}#{result}", Inverse:=True)
    End If
End Function
Public Function HasChanged(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant, _
                                                                   Optional by As Variant) As Boolean
    If IsMissing(by) Then by = vAdditionalArgument
    If IsObject(Actual) Then
        If IsObject(Expected) Then
            If Not dictOtherCaptures Is Nothing Then
                Dim colChanges As New Collection
                Dim CaptureName
                On Error Resume Next
                For Each CaptureName In dictOtherCaptures.Keys
                    If Len(CaptureName) Then
                        If Not IsLike(GetProperty(CaptureName, Of:=Actual), dictOtherCaptures(CaptureName)) Then
                            HasChanged = True
                            colChanges.Add CaptureName
                        End If
                    Else
                        If Not IsSame(Actual, Expected) Then
                            HasChanged = True
                        End If
                    End If
                Next CaptureName
                If colChanges.Count > 0 Then
                    Result = " [" & colChanges(1)
                    If colChanges.Count > 3 Then
                        Result = Result & "," & colChanges(2) & ",+" & colChanges.Count - 2 & "others"
                    Else
                        Dim i As Integer
                        For i = 2 To colChanges.Count
                            Result = Result & "," & colChanges(i)
                        Next i
                    End If
                    Result = Result & "]"
                End If
            Else
                Warn "Compared only the objects reference, not the actual properties of each"
                HasChanged = Not IsSame(Actual, Expected)
            End If
        Else
            HasChanged = True
        End If
    ElseIf (VarType(Actual) = vbBoolean Or VarType(Expected) = vbBoolean) Then
        HasChanged = (Actual <> Expected)
    ElseIf VarType(Actual) = vbString And VarType(Expected) = vbString Then
        HasChanged = (StrComp(Actual, Expected, vbBinaryCompare) <> 0)
    ElseIf (VarType(Actual) = vbDouble Or VarType(Expected) = vbDouble) Then
        HasChanged = Not IsCloseTo(Actual, (Expected + by), Result, 15)
    Else
        HasChanged = (Actual <> (Expected + by))
    End If
End Function

'========================================================================================================================================'

Private Function IsLike(Var1 As Variant, Var2 As Variant) As Boolean
    If IsObject(Var1) Or IsObject(Var2) Then
        IsLike = IsSame(Var1, Var2)
    Else
        IsLike = IsEqual(Var1, Var2)
    End If
End Function

'========================================================================================================================================'

Public Function NoError() As VBASpecExpectation
    Me.Actual = Err.Number
    Set NoError = VerifyActual("IsError", 0, "Expected no errors, but got Error ##{actual}(#{result})")
    Err.Clear
End Function
Public Function Error(Optional Expected As Variant) As VBASpecExpectation
    Me.Actual = Err.Number
    If IsMissing(Expected) Then
        Set Error = VerifyActual("IsError", Expected, "Expected an error, got none")
    Else
        Set Error = VerifyActual("IsError", Expected, "Expected Error ##{expected}, got " & _
                                                      IIf(Err.Number = 0, "none", "##{actual}(#{result})"))
    End If
    Err.Clear
End Function
Public Function IsError(Actual As Variant, Expected As Variant, Optional ByRef Result As Variant) As Boolean
    If IsMissing(Expected) Then
        IsError = (Err.Number <> 0)
    Else
        IsError = (Err.Number = Expected)
    End If
    Result = Err.Description
    bErrorChecked = True
End Function
Public Function WasRaised() As Boolean
    'Mostly for aesthetic reasons...
    'So we can write: .Expect.Error(11).WasRaised
    '             or: .Expect.NoError.WasRaised
    'But also returns True if the expectation is met
    WasRaised = bErrorChecked And Me.Passed
End Function

'========================================================================================================================================'

Private Function VerifyActual(ProcName As String, Expected As Variant, Optional sMessage As String = vbNullString, _
                                                                       Optional Inverse As Boolean = False) As VBASpecExpectation
    If Not bFailed Then
        Assign Expected, Into:=vExpected
        
        Dim Result As Variant
        If Inverse Then
            bFailed = CallByName(Me, ProcName, VbMethod, Me.Actual, Expected, Result)
            bPassed = Not bFailed
        Else
            bPassed = CallByName(Me, ProcName, VbMethod, Me.Actual, Expected, Result)
            bFailed = Not bPassed
        End If
        
        If bFailed Then
            On Error Resume Next
            
            If Len(sMessage) = 0 Then sMessage = IIf(Inverse, FAILURE_MSG_NEGATED, FAILURE_MSG)
            
            Dim bShowAddress As Boolean
            bShowAddress = (AnyValueAsText(Me.Actual) = AnyValueAsText(Expected))
            
            Dim bShowType As Boolean
            If Not Inverse Then
                If (InStr(1, ProcName, "Change", vbBinaryCompare) = 0) Then
                    Select Case VarType(Expected)
                    Case vbBoolean: bShowType = (CBool(Me.Actual) = CBool(Expected))
                    Case vbInteger: bShowType = (CInt(Me.Actual) = CInt(Expected))
                    Case vbDouble:  bShowType = (CDbl(Me.Actual) = CDbl(Expected))
                    Case Else:      bShowType = (AnyValueAsText(Me.Actual) = AnyValueAsText(Expected))
                    End Select
                End If
            End If
            
            If InStr(1, sMessage, ACTUAL_INTERPOLATION, vbBinaryCompare) > 0 Then
                sMessage = Replace$(sMessage, ACTUAL_INTERPOLATION, AnyValueAsText(Me.Actual, bShowAddress, bShowType), Compare:=vbBinaryCompare)
            End If
            If InStr(1, sMessage, EXPECTED_INTERPOLATION, vbBinaryCompare) > 0 Then
                sMessage = Replace$(sMessage, EXPECTED_INTERPOLATION, AnyValueAsText(Expected, bShowAddress, bShowType), Compare:=vbBinaryCompare)
            End If
            If InStr(1, sMessage, RESULT_INTERPOLATION, vbBinaryCompare) > 0 Then
                sMessage = Replace$(sMessage, RESULT_INTERPOLATION, CStr(Result), Compare:=vbBinaryCompare)
            End If
            If InStr(1, sMessage, TRUE_INTERPOLATION, vbBinaryCompare) > 0 Then
                sMessage = Replace$(sMessage, TRUE_INTERPOLATION, AnyValueAsText(True), Compare:=vbBinaryCompare)
            End If
            If InStr(1, sMessage, FALSE_INTERPOLATION, vbBinaryCompare) > 0 Then
                sMessage = Replace$(sMessage, FALSE_INTERPOLATION, AnyValueAsText(False), Compare:=vbBinaryCompare)
            End If
            sFailureMessage = sMessage
            If Len(sWarningMessage) > 0 Then sFailureMessage = sFailureMessage & vbNewLine & sWarningMessage
        End If
    End If
    
    Set VerifyActual = Me
End Function

'========================================================================================================================================'

#If TESTING Then
Friend Sub UnitTest(Suite As VBASpecSuite)
    With Suite.Describe("VBASpecExpectation")
        Dim subject As VBASpecExpectation
        
        Dim ShowExamples As Boolean
        ShowExamples = False
        
        With .Describe("#Actual")
            Set subject = New VBASpecExpectation
            
            With .It("is empty by default")
                .Expect(subject.Actual).ToBeEmpty
            End With
            With .It("accepts anything")
                subject.Actual = True
                .Expect(subject.Actual).ToBeA "Boolean"
                subject.Actual = "String"
                .Expect(subject.Actual).ToBeA "String"
                subject.Actual = 1
                .Expect(subject.Actual).ToBeAn "Integer"
                subject.Actual = 1#
                .Expect(subject.Actual).ToBeA "Double"
                subject.Actual = Null
                .Expect(subject.Actual).ToBeA "Null"
                Set subject.Actual = New VBASpecExpectation
                .Expect(subject.Actual).ToBeA "VBASpecExpectation"
            End With
        End With
        
        With .Describe("#Expected")
            Set subject = New VBASpecExpectation
            
            With .It("is empty by default")
                .Expect(subject.Expected).ToBeEmpty
            End With
            With .It("accepts anything")
                subject.Expected = True
                .Expect(subject.Expected).ToBeA "Boolean"
                subject.Expected = "String"
                .Expect(subject.Expected).ToBeA "String"
                subject.Expected = 1
                .Expect(subject.Expected).ToBeAn "Integer"
                subject.Expected = 1#
                .Expect(subject.Expected).ToBeA "Double"
                subject.Expected = Null
                .Expect(subject.Expected).ToBeA "Null"
                Set subject.Expected = New VBASpecExpectation
                .Expect(subject.Expected).ToBeA "VBASpecExpectation"
            End With
        End With
        
        With .Describe("#Passed")
            Set subject = New VBASpecExpectation
            
            With .It("is false by default")
                .Expect(subject.Passed).ToBeFalse
            End With
            With .It("is true when the expectation is met")
                subject.Actual = 1 + 1
                subject.ToEqual 2
                .Expect(subject.Passed).ToBeTrue
            End With
            With .It("is false when the expectation is not met")
                subject.Actual = 1 + 1
                subject.ToEqual 3
                .Expect(subject.Passed).ToBeFalse
            End With
        End With
        
        With .Describe("#Failed")
            Set subject = New VBASpecExpectation
            
            With .It("is false by default")
                .Expect(subject.Failed).ToBeFalse
            End With
            With .It("is false when the expectation is met")
                subject.Actual = 1 + 1
                subject.ToEqual 2
                .Expect(subject.Failed).ToBeFalse
            End With
            With .It("is true when the expectation is not met")
                subject.Actual = 1 + 1
                subject.ToEqual 3
                .Expect(subject.Failed).ToBeTrue
            End With
        End With
        
        With .Describe("#FailureMessage")
            Set subject = New VBASpecExpectation
            
            With .Context("when the expectation is met")
                subject.Actual = True
                subject.ToBeTrue
                
                With .It("is empty")
                    .Expect(subject.FailureMessage).ToBeEmpty
                End With
            End With
            With .Context("when the expectation is not met")
                subject.Actual = 2358.13
                subject.ToEqual "Fibonacci"
                
                With .It("returns a failure message")
                    .Expect(subject.FailureMessage).ToNotBeEmpty
                End With
                With .It("includes the Actual in the failure message")
                    .Expect(subject.FailureMessage).ToInclude "2358" & Mid$(1 / 2, 2, 1) & "13"
                End With
                With .It("includes the Expected in the failure message")
                    .Expect(subject.FailureMessage).ToInclude "Fibonacci"
                End With
                With .It("surrounds strings with double quotes")
                    .Expect(subject.FailureMessage).ToInclude """Fibonacci"""
                End With
            End With
        End With
        
        With .Describe(".ToEqual/.ToNotEqual")
            With .It("works with Booleans")
                .Expect(True).ToEqual True
                .Expect(True).ToNotEqual False
            End With
            With .It("works with Strings")
                .Expect("Abc").ToEqual "abc"
                .Expect("Abc").ToNotEqual "Abcd"
            End With
            With .It("works with Integers")
                .Expect(2048).ToEqual 2048
                .Expect(2048).ToNotEqual 2049
            End With
            With .It("works with Decimals")
                .Expect(3.14159265358979).ToEqual 3.14159265358979
                .Expect(3.14159265358979).ToNotEqual 3.1415926535898
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(True).ToEqual False
                    .Expect(True).ToNotEqual True
                End With
            End If
        End With
        
        With .Describe(".ToBe/.ToNotBe")
            Dim SameObject As New VBASpecExpectation
            Dim OtherObject As New VBASpecExpectation
            
            Set subject = SameObject
            
            With .It("works if both Actual and Expected are Objects")
                .Expect(subject).ToBe SameObject
                .Expect(subject).ToNotBe OtherObject
            End With
            With .It("gracefully fails if either Actual or Expected is not an Object")
                On Error Resume Next
                .Expect(subject).ToNotBe "aString"
                .Expect("aString").ToNotBe subject
                .Expect("aString").ToNotBe "aString"
                .Expect.NoError.WasRaised
                On Error GoTo 0
            End With
            
            With .Context("followed by a Property Name")
                Set .subject = New Image
                
                With .It("works if that Property does exist")
                    .Expect_It.ToNotBe "AutoSize"
                    .Expect_It.ToBe "Enabled"
                End With
                With .It("gracefully fails if")
                    On Error Resume Next
                    With .Context("that Property doesn't exist")
                        .Expect_It.ToNotBe "NoProp"
                        .Expect.NoError.WasRaised
                    End With
                    With .Context("Actual is not even an object")
                        .Expect("subject").ToNotBe "NoProp"
                        .Expect.NoError.WasRaised
                    End With
                    On Error GoTo 0
                End With
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(subject).ToBe OtherObject
                    .Expect(subject).ToNotBe SameObject
                End With
            End If
        End With
        
        With .Describe(".ToHave/.ToNotHave")
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(1)
            Dim cell As Range
            Set cell = ws.Cells(1, 1)
            
            With .It("works if that Property does exist")
                cell.Formula = "=1"
                .Expect(cell).ToHave "formula"
                cell.Formula = ""
                .Expect(cell).ToNotHave "formula"
            End With
            With .It("gracefully fails if")
                On Error Resume Next
                With .Context("that Property doesn't exist")
                    .Expect_It.ToNotHave "NoProp"
                    .Expect.NoError.WasRaised
                End With
                With .Context("Actual is not even an object")
                    .Expect("subject").ToNotHave "NoProp"
                    .Expect.NoError.WasRaised
                End With
                On Error GoTo 0
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    cell.Formula = ""
                    .Expect(cell).ToHave "formula"
                    cell.Formula = "=1"
                    .Expect(cell).ToNotHave "formula"
                    cell.Formula = ""
                End With
            End If
        End With
        
        With .Describe(".ToBeA/.ToBeAn/.ToNotBeA/.ToNotBeAn")
            With .It("works")
                .Expect(True).ToBeA "Boolean"
                .Expect("String").ToBeA "String"
                .Expect(2048).ToBeAn "Integer"
                .Expect(100#).ToBeA "Double"
                .Expect(Null).ToBeA "Null"
                Dim oObject As New VBASpecExpectation
                .Expect(oObject).ToBeA "VBASpecExpectation"
            End With
            With .It("does not coerce Strings")
                .Expect(True).ToBeA "Boolean"
                .Expect("True").ToNotBeA "Boolean"
                .Expect("True").ToBeA "String"
                .Expect(2048).ToBeAn "Integer"
                .Expect("2048").ToNotBeAn "Integer"
                .Expect("2048").ToBeA "String"
            End With
            With .It("does not coerce Numbers")
                .Expect(100#).ToBeA "Double"
                .Expect(100#).ToNotBeA "Single"
                .Expect(100#).ToNotBeAn "Integer"
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(True).ToBeA "String"
                    .Expect(True).ToNotBeA "Boolean"
                End With
            End If
        End With
        
        With .Describe(".ToBeEmpty/.ToNotBeEmpty")
            With .It("passes when Actual is Missing, Nothing, Empty or Null")
                .Expect(Nothing).ToBeEmpty
                .Expect(Empty).ToBeEmpty
                .Expect(Null).ToBeEmpty
                .Expect().ToBeEmpty
                .Expect(0).ToNotBeEmpty
            End With
            With .It("passes when Actual is an empty String")
                .Expect("").ToBeEmpty
                .Expect(" ").ToBeEmpty
                .Expect("!").ToNotBeEmpty
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect("-").ToBeEmpty
                    .Expect(" ").ToNotBeEmpty
                End With
            End If
        End With
        With .Describe(".ToBeNothing/.ToNotBeNothing")
            With .It("passes only when Actual is Nothing")
                .Expect(Nothing).ToBeNothing
                .Expect(Empty).ToNotBeNothing
                .Expect(Null).ToNotBeNothing
                .Expect().ToNotBeNothing
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(Empty).ToBeNothing
                    .Expect(Nothing).ToNotBeNothing
                End With
            End If
        End With
        With .Describe(".ToBeZero/.ToNotBeZero")
            With .It("passes if Actual equals 0")
                .Expect(0).ToBeZero
                .Expect(0#).ToBeZero
                .Expect("0").ToNotBeZero
                .Expect(0.000000000000001).ToNotBeZero
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeZero
                    .Expect(0).ToNotBeZero
                End With
            End If
        End With
        
        With .Describe(".ToBeFalsy")
            With .It("passes when Actual evaluates to False")
                .Expect(False).ToBeFalsy
                .Expect(0).ToBeFalsy
                .Expect(0#).ToBeFalsy
            End With
            
            If ShowExamples And Not .Failed Then
                .Context("- Example :").Expect(True).ToBeFalsy
            End If
        End With
        With .Describe(".ToBeFalse")
            With .It("passes when Actual is False")
                .Expect(False).ToBeFalse
            End With
            
            If ShowExamples And Not .Failed Then
                .Context("- Example :").Expect(True).ToBeFalse
            End If
        End With
        
        With .Describe(".ToBeTruthy")
            With .It("passes when Actual evaluates to True")
                .Expect(True).ToBeTruthy
                .Expect(1).ToBeTruthy
                .Expect(2).ToBeTruthy
                .Expect(0.1).ToBeTruthy
            End With
            
            If ShowExamples And Not .Failed Then
                .Context("- Example :").Expect(False).ToBeTruthy
            End If
        End With
        With .Describe(".ToBeTrue")
            With .It("passes when Actual is True")
                .Expect(True).ToBeTrue
            End With
            
            If ShowExamples And Not .Failed Then
                .Context("- Example :").Expect(False).ToBeTrue
            End If
        End With
        
        With .Describe(".ToBeLessThan")
            With .It("passes when Actual is less than Expected")
                .Expect(1).ToBeLessThan 2
                .Expect(1.99).ToBeLessThan 2
            End With
        End With
        With .Describe(".ToBeLT")
            With .It("passes when Actual < Expected")
                .Expect(1).ToBeLT 2
                .Expect(1.99).ToBeLT 2
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeLessThan 0
                    .Expect(1).ToBeLT 0
                End With
            End If
        End With
        
        With .Describe(".ToBeLessThanOrEqualTo")
            With .It("passes when Actual is less than or equal to Expected")
                .Expect(1).ToBeLessThanOrEqualTo 2
                .Expect(2).ToBeLessThanOrEqualTo 2
                .Expect(1.99).ToBeLessThanOrEqualTo 2
                .Expect(2#).ToBeLessThanOrEqualTo 2
            End With
        End With
        With .Describe(".ToBeLTE")
            With .It("passes when Actual <= Expected")
                .Expect(1).ToBeLTE 2
                .Expect(2).ToBeLTE 2
                .Expect(1.99).ToBeLTE 2
                .Expect(2#).ToBeLTE 2
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeLessThanOrEqualTo 0
                    .Expect(1).ToBeLTE 0
                End With
            End If
        End With
        
        With .Describe(".ToBeGreaterThan")
            With .It("passes when Actual is greater than Expected")
                .Expect(2).ToBeGreaterThan 1
                .Expect(1.01).ToBeGreaterThan 1
            End With
        End With
        With .Describe(".ToBeGT")
            With .It("passes when Actual > Expected")
                .Expect(2).ToBeGT 1
                .Expect(1.01).ToBeGT 1
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeGreaterThan 2
                    .Expect(1).ToBeGT 2
                End With
            End If
        End With
        
        With .Describe(".ToBeGreaterThanOrEqualTo")
            With .It("passes when Actual is greater than or equal to Expected")
                .Expect(2).ToBeGreaterThanOrEqualTo 1
                .Expect(1).ToBeGreaterThanOrEqualTo 1
                .Expect(1.01).ToBeGreaterThanOrEqualTo 1
                .Expect(1#).ToBeGreaterThanOrEqualTo 1
            End With
        End With
        With .Describe(".ToBeGTE")
            With .It("passes when Actual >= Expected")
                .Expect(2).ToBeGTE 1
                .Expect(1).ToBeGTE 1
                .Expect(1.01).ToBeGTE 1
                .Expect(1#).ToBeGTE 1
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeGreaterThanOrEqualTo 2
                    .Expect(1).ToBeGTE 2
                End With
            End If
        End With
        
        With .Describe(".ToBeCloseTo/.ToNotBeCloseTo")
            With .It("passes when Actual is close to Expected (defaults to 2 sigificant figures)")
                .Expect(1.01).ToBeCloseTo 1
                .Expect(1.12).ToBeCloseTo 1.1
                .Expect(2).ToNotBeCloseTo 1
                .Expect(1.1).ToNotBeCloseTo 1
                .Expect(1.08).ToNotBeCloseTo 1
                .Expect(0.0001).ToNotBeCloseTo 0.00001
            End With
            With .It("can be set up to 15 significant figures")
                .Expect(3.14159265358979).ToBeCloseTo 3, 1
                .Expect(3.14159265358979).ToBeCloseTo 3.1, 2
                .Expect(3.14159265358979).ToBeCloseTo 3.14, 3
                .Expect(3.14159265358979).ToBeCloseTo 3.142, 4
                .Expect(3.14159265358979).ToBeCloseTo 3.1416, 5
                .Expect(3.14159265358979).ToBeCloseTo 3.14159, 6
                .Expect(3.14159265358979).ToBeCloseTo 3.141593, 7
                .Expect(3.14159265358979).ToBeCloseTo 3.1415927, 8
                .Expect(3.14159265358979).ToBeCloseTo 3.14159265, 9
                .Expect(3.14159265358979).ToBeCloseTo 3.141592654, 10
                .Expect(3.14159265358979).ToBeCloseTo 3.1415926536, 11
                .Expect(3.14159265358979).ToBeCloseTo 3.14159265359, 12
                .Expect(3.14159265358979).ToBeCloseTo 3.1415926535898, 14
                .Expect(3.14159265358979).ToBeCloseTo 3.14159265358979, 15
                
                .Expect(3.14159265358979).ToNotBeCloseTo 3, 2
                .Expect(3.14159265358979).ToNotBeCloseTo 3.1, 3
                .Expect(3.14159265358979).ToNotBeCloseTo 3.14, 4
                .Expect(3.14159265358979).ToNotBeCloseTo 3.142, 5
                .Expect(3.14159265358979).ToNotBeCloseTo 3.1416, 6
                .Expect(3.14159265358979).ToNotBeCloseTo 3.14159, 7
                .Expect(3.14159265358979).ToNotBeCloseTo 3.141593, 8
                .Expect(3.14159265358979).ToNotBeCloseTo 3.1415927, 9
                .Expect(3.14159265358979).ToNotBeCloseTo 3.14159265, 10
                .Expect(3.14159265358979).ToNotBeCloseTo 3.141592654, 11
                .Expect(3.14159265358979).ToNotBeCloseTo 3.1415926535898, 15
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect(1).ToBeCloseTo 2
                    .Expect(1).ToBeCloseTo 1.12
                    .Expect(1).ToBeCloseTo 1.000001, 6
                    .Expect(1).ToNotBeCloseTo 1
                    .Expect(1).ToNotBeCloseTo 1.001
                End With
            End If
        End With
        
        Dim sSubject As String
        Dim arrSubject As Variant
        Dim colSubject As Collection
        
        Dim colIncluded As New Collection
        Dim colOther As New Collection
        With .Describe(".ToInclude/.ToNotInclude")
            With .Context("when Actual is a String")
                sSubject = "a String"
                
                With .It("works")
                    .Expect(sSubject).ToInclude "String"
                    .Expect(sSubject).ToNotInclude "Strong"
                End With
                With .It("is not case sensitive")
                    .Expect(sSubject).ToInclude "string"
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect("Abc").ToInclude "d"
                        .Expect("Abc").ToNotInclude "b"
                    End With
                End If
            End With
            
            With .Context("when Actual is an Array")
                arrSubject = Array(1, "Two", colIncluded)
                
                With .It("works with expressions")
                    .Expect(arrSubject).ToInclude 1
                    .Expect(arrSubject).ToInclude "Two"
                    .Expect(arrSubject).ToNotInclude 2
                End With
                With .It("works with objects")
                    .Expect(arrSubject).ToInclude colIncluded
                    .Expect(arrSubject).ToNotInclude colOther
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect(Array(1, 2, 3)).ToInclude "4"
                        .Expect(Array(1, 2, 3)).ToNotInclude "2"
                    End With
                End If
            End With
            
            With .Context("when Actual is a Collection")
                Set colSubject = New Collection
                colSubject.Add 1
                colSubject.Add "Two"
                colSubject.Add colIncluded
                
                With .It("works with expressions")
                    .Expect(colSubject).ToInclude 1
                    .Expect(colSubject).ToInclude "Two"
                    .Expect(colSubject).ToNotInclude 2
                End With
                With .It("works with objects")
                    .Expect(colSubject).ToInclude colIncluded
                    .Expect(colSubject).ToNotInclude colOther
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect(colSubject).ToInclude 2
                        .Expect(colSubject).ToNotInclude 1
                    End With
                End If
            End With
        End With
        
        With .Describe(".ToBeginWith/.ToNotBeginWith")
            With .Context("when Actual is a String")
                sSubject = "String"
                
                With .It("works")
                    .Expect(sSubject).ToBeginWith "Str"
                    .Expect(sSubject).ToNotBeginWith "St_"
                End With
                With .It("is not case sensitive")
                    .Expect(sSubject).ToBeginWith "str"
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect(sSubject).ToBeginWith "St_"
                        .Expect(sSubject).ToNotBeginWith "Str"
                    End With
                End If
            End With
            
            With .Context("when Actual is an Array")
                With .It("works with expressions")
                    arrSubject = Array(1, "Two", colIncluded)
                    
                    .Expect(arrSubject).ToBeginWith 1
                    .Expect(arrSubject).ToNotBeginWith "Two"
                End With
                With .It("works with objects")
                    arrSubject = Array(colIncluded, "Two", 3)
                    
                    .Expect(arrSubject).ToBeginWith colIncluded
                    .Expect(arrSubject).ToNotBeginWith colOther
                End With
            End With
            
            With .Context("when Actual is a Collection")
                With .It("works with expressions")
                    Set colSubject = New Collection
                    colSubject.Add 1
                    colSubject.Add "Two"
                    colSubject.Add colIncluded
                    
                    .Expect(colSubject).ToBeginWith 1
                    .Expect(colSubject).ToNotBeginWith "Two"
                End With
                With .It("works with objects")
                    Set colSubject = New Collection
                    colSubject.Add colIncluded
                    colSubject.Add "Two"
                    colSubject.Add 3
                    
                    .Expect(colSubject).ToBeginWith colIncluded
                    .Expect(colSubject).ToNotBeginWith colOther
                End With
            End With
        End With
        
        With .Describe(".ToEndWith/.ToNotEndWith")
            With .Context("when Actual is a String")
                sSubject = "String"
                
                With .It("works")
                    .Expect(sSubject).ToEndWith "ing"
                    .Expect(sSubject).ToNotEndWith "_ng"
                End With
                With .It("is not case sensitive")
                    .Expect(sSubject).ToEndWith "Ing"
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect(sSubject).ToEndWith "_ng"
                        .Expect(sSubject).ToNotEndWith "ing"
                    End With
                End If
            End With
            
            With .Context("when Actual is an Array")
                With .It("works with expressions")
                    arrSubject = Array(colIncluded, "Two", 3)
                    
                    .Expect(arrSubject).ToEndWith 3
                    .Expect(arrSubject).ToNotEndWith "Two"
                End With
                With .It("works with objects")
                    arrSubject = Array(1, "Two", colIncluded)
                    
                    .Expect(arrSubject).ToEndWith colIncluded
                    .Expect(arrSubject).ToNotEndWith colOther
                End With
            End With
            With .Context("when Actual is a Collection")
                With .It("works with expressions")
                    Set colSubject = New Collection
                    colSubject.Add colIncluded
                    colSubject.Add "Two"
                    colSubject.Add 3
                    
                    .Expect(colSubject).ToEndWith 3
                    .Expect(colSubject).ToNotEndWith "Two"
                End With
                With .It("works with objects")
                    Set colSubject = New Collection
                    colSubject.Add 1
                    colSubject.Add "Two"
                    colSubject.Add colIncluded
                    
                    .Expect(colSubject).ToEndWith colIncluded
                    .Expect(colSubject).ToNotEndWith colOther
                End With
            End With
        End With
        
        With .Describe(".ToChange/.ToNotChange")
            With .It("works with Booleans")
                .subject = False
                .Capture
                .subject = False
                .Expect_It.ToNotChange
                .subject = True
                .Expect_It.ToChange
            End With
            With .It("works with Integers")
                .subject = 1
                .Capture
                .subject = 1
                .Expect_It.ToNotChange
                .Expect_It.ToChange by:=0
                .subject = 2
                .Expect_It.ToChange
                .Expect_It.ToChange by:=1
                .Expect_It.ToNotChange by:=2
            End With
            With .It("works with Doubles")
                .subject = 1.5
                .Capture
                .subject = 2
                .Expect_It.ToChange
                .Expect_It.ToChange by:=0.5
            End With
            With .It("works with Strings")
                .subject = "Origin"
                .Capture
                .subject = "Original"
                .Expect_It.ToChange
                .subject = "Orig"
                .Expect_It.ToChange
            End With
            With .It("works with Objects, but only checks captured properties")
                Set .subject = New Collection
                .Capture
                .subject.Add "something"
                .Expect_It.ToNotChange
                Set .subject = New Collection
                .Expect_It.ToChange
                .Capture
                .subject.Add "something"
                .Expect_It.ToNotChange
                .Capture "Count"
                .subject.Add "something else"
                .Expect_It.ToChange
            End With
            
            With .It("works with mixed types")
                On Error Resume Next
                .subject = "Origin"
                .Capture
                .subject = 1
                .Expect_It.ToChange
                Set .subject = New Collection
                .Expect_It.ToChange
                .Expect.NoError.WasRaised
                On Error GoTo 0
            End With
            
            With .It("works between contexts")
                .subject = 1
                .Capture
                With .Context
                    .subject = 2
                    .Expect_It.ToChange
                End With
                .Expect_It.ToNotChange
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .subject = 1
                    .Capture
                    .Expect_It.ToChange
                    .Expect_It.ToChange by:=1
                    .subject = 2
                    .Expect_It.ToNotChange
                    .Expect_It.ToNotChange by:=1
                End With
            End If
       End With
        
        On Error Resume Next
        Dim errorTrigger
        With .Describe(".NoError")
            Set subject = New VBASpecExpectation
            
            With .It("passes if no error was raised")
                errorTrigger = 1 / 1
                .Expect(subject.NoError.Passed).ToBeTrue
            End With
            With .It("fails if an error was raised")
                errorTrigger = 1 / 0
                .Expect(subject.NoError.Failed).ToBeTrue
            End With
            With .It("clears the Err afterward")
                errorTrigger = 1 / 0
                subject.NoError
                .Expect(Err.Number).ToBeZero
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    errorTrigger = 1 / 0
                    .Expect.NoError
                End With
            End If
            
            With .Describe(".WasRaised")
                Set subject = New VBASpecExpectation
                
                With .Context("when no error was raised")
                    errorTrigger = 1 / 1
                    
                    With .It("returns True")
                        .Expect(subject.NoError.WasRaised).ToBeTrue
                    End With
                End With
                With .Context("when any error was raised")
                    errorTrigger = 1 / 0
                    
                    With .It("returns False")
                        .Expect(subject.NoError.WasRaised).ToBeFalse
                    End With
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect.Error.WasRaised
                        .Expect.Error(11).WasRaised
                        errorTrigger = 1 / 0
                        .Expect.NoError.WasRaised
                    End With
                End If
            End With
        End With
        
        With .Describe(".Error")
            With .Context("when no number given")
                Set subject = New VBASpecExpectation
                
                With .It("passes if any error was raised")
                    errorTrigger = 1 / 0
                    .Expect(subject.Error.Passed).ToBeTrue
                End With
                With .It("fails if no error was raised")
                    errorTrigger = 1 / 1
                    .Expect(subject.Error.Failed).ToBeTrue
                End With
            End With
            With .Context("when a number is given")
                Set subject = New VBASpecExpectation
                
                With .It("passes if the error was raised")
                    errorTrigger = 1 / 0
                    .Expect(subject.Error(11).Passed).ToBeTrue
                End With
                With .It("fails if another error was raised")
                    errorTrigger = 1 / 0
                    .Expect(subject.Error(12).Failed).ToBeTrue
                End With
                With .It("fails if no error was raised")
                    errorTrigger = 1 / 1
                    .Expect(subject.Error(11).Failed).ToBeTrue
                End With
            End With
            
            With .It("clears the Err afterward")
                errorTrigger = 1 / 0
                subject.Error
                .Expect(Err.Number).ToBeZero
            End With
            
            If ShowExamples And Not .Failed Then
                With .Context("- Example :")
                    .Expect.Error
                    .Expect.Error 11
                End With
            End If
            
            With .Describe(".WasRaised")
                With .Context("when no error was raised")
                    Set subject = New VBASpecExpectation
                    
                    errorTrigger = 1 / 1
                    
                    With .It("returns False")
                        .Expect(subject.Error.WasRaised).ToBeFalse
                    End With
                End With
                With .Context("when any error was raised")
                    Set subject = New VBASpecExpectation
                    
                    errorTrigger = 1 / 0
                    
                    With .It("returns True")
                        .Expect(subject.Error.WasRaised).ToBeTrue
                    End With
                End With
                With .Context("when another error was raised")
                    Set subject = New VBASpecExpectation
                    
                    errorTrigger = 1 / 0
                    
                    With .It("returns False")
                        .Expect(subject.Error(1).WasRaised).ToBeFalse
                    End With
                End With
                
                If ShowExamples And Not .Failed Then
                    With .Context("- Example :")
                        .Expect.Error.WasRaised
                        .Expect.Error(11).WasRaised
                    End With
                End If
            End With
        End With
        On Error GoTo 0
    End With
End Sub
#End If
