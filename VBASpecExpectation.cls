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
Private Const ACTUALTYPE_INTERPOLATION As String = "#{actualtype}"
Private Const EXPECTED_INTERPOLATION As String = "#{expected}"

Private Const FAILURE_MSG As String = "Expected #{expected}, got #{actual}."
Private Const FAILURE_MSG_NEGATED As String = "Not expected to be #{actual}."

Private vActual As Variant
Private vExpected As Variant

Private bPassed As Boolean
Private bFailed As Boolean
Private sFailureMessage As String

Private iSignificantFigures As Integer
Private bErrorChecked As Boolean

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

Public Function ToEqual(Expected As Variant) As VBASpecExpectation
    Set ToEqual = VerifyActual("IsEqual", Expected)
End Function
Public Function ToNotEqual(Expected As Variant) As VBASpecExpectation
    Set ToNotEqual = VerifyActual("IsEqual", Expected, Inverse:=True)
End Function
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

Public Function ToBe(Expected As Variant) As VBASpecExpectation
    Set ToBe = VerifyActual("IsSame", Expected)
End Function
Public Function ToNotBe(Expected As Variant) As VBASpecExpectation
    Set ToNotBe = VerifyActual("IsSame", Expected, Inverse:=True)
End Function
Public Function IsSame(Actual As Variant, Expected As Variant) As Boolean
    If IsObject(Actual) And IsObject(Expected) Then
        IsSame = (Actual Is Expected)
    Else
        IsSame = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeA(Expected As Variant) As VBASpecExpectation
    Set ToBeA = VerifyActual("IsA", Expected, "Expected #{actual} to be a " & CStr(Expected) & ", but is a #{actualtype}.")
End Function
Public Function ToBeAn(Expected As Variant) As VBASpecExpectation
    Set ToBeAn = VerifyActual("IsA", Expected, "Expected #{actual} to be an " & CStr(Expected) & ", but is a #{actualtype}.")
End Function
Public Function ToNotBeA(Expected As Variant) As VBASpecExpectation
    Set ToNotBeA = VerifyActual("IsA", Expected, "Expected #{actual} not to be a " & CStr(Expected) & ", but it is.", Inverse:=True)
End Function
Public Function ToNotBeAn(Expected As Variant) As VBASpecExpectation
    Set ToNotBeAn = VerifyActual("IsA", Expected, "Expected #{actual} not to be an " & CStr(Expected) & ", but it is.", Inverse:=True)
End Function
Public Function IsA(Actual As Variant, Expected As Variant) As Boolean
    IsA = (TypeName(Actual) = Expected)
End Function

'========================================================================================================================================'

Public Function ToBeEmpty() As VBASpecExpectation
    Set ToBeEmpty = VerifyActual("IsEmpty", Empty, "#{actualtype} expected to be empty.")
End Function
Public Function ToNotBeEmpty() As VBASpecExpectation
   Set ToNotBeEmpty = VerifyActual("IsEmpty", Empty, Inverse:=True)
End Function
Public Function IsEmpty(Actual As Variant, Optional Expected As Variant) As Boolean
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
Public Function IsNothing(Actual As Variant, Optional Expected As Variant) As Boolean
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
Public Function IsZero(Actual As Variant, Optional Expected As Variant) As Boolean
    IsZero = IsEqual(Actual, 0)
End Function

'========================================================================================================================================'

Public Function ToBeFalsy() As VBASpecExpectation
    Set ToBeFalsy = VerifyActual("IsFalsy", False, "Expected #{actual} to evaluate to #{expected}.")
End Function
Public Function IsFalsy(Actual As Variant, Optional Expected As Variant) As Boolean
    IsFalsy = IIf(Actual, False, True)
End Function

'========================================================================================================================================'

Public Function ToBeFalse() As VBASpecExpectation
    Set ToBeFalse = VerifyActual("IsFalse", False)
End Function
Public Function IsFalse(Actual As Variant, Optional Expected As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsFalse = (Actual = False)
    Else
        IsFalse = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeTruthy() As VBASpecExpectation
    Set ToBeTruthy = VerifyActual("IsTruthy", True, "Expected #{actual} to evaluate to #{expected}.")
End Function
Public Function IsTruthy(Actual As Variant, Optional Expected As Variant) As Boolean
    IsTruthy = IIf(Actual, True, False)
End Function

'========================================================================================================================================'

Public Function ToBeTrue() As VBASpecExpectation
    Set ToBeTrue = VerifyActual("IsTrue", True)
End Function
Public Function IsTrue(Actual As Variant, Optional Expected As Variant) As Boolean
    If VarType(Actual) = vbBoolean Then
        IsTrue = (Actual = True)
    Else
        IsTrue = False
    End If
End Function

'========================================================================================================================================'

Public Function ToBeLessThan(Expected As Variant) As VBASpecExpectation
    Set ToBeLessThan = VerifyActual("IsLT", Expected, "Expected #{actual} to be less than #{expected}.")
End Function
Public Function ToBeLT(Expected As Variant) As VBASpecExpectation
    Set ToBeLT = VerifyActual("IsLT", Expected, "Expected #{actual} to be < #{expected}.")
End Function
Public Function IsLT(Actual As Variant, Expected As Variant) As Boolean
    IsLT = (Actual < Expected)
End Function

'========================================================================================================================================'

Public Function ToBeLessThanOrEqualTo(Expected As Variant) As VBASpecExpectation
    Set ToBeLessThanOrEqualTo = VerifyActual("IsLTE", Expected, "Expected #{actual} to be less than or equal to #{expected}.")
End Function
Public Function ToBeLTE(Expected As Variant) As VBASpecExpectation
    Set ToBeLTE = VerifyActual("IsLTE", Expected, "Expected #{actual} to be <= #{expected}.")
End Function
Public Function IsLTE(Actual As Variant, Expected As Variant) As Boolean
    IsLTE = (Actual <= Expected)
End Function

'========================================================================================================================================'

Public Function ToBeGreaterThan(Expected As Variant) As VBASpecExpectation
    Set ToBeGreaterThan = VerifyActual("IsGT", Expected, "Expected #{actual} to be greater than #{expected}.")
End Function
Public Function ToBeGT(Expected As Variant) As VBASpecExpectation
    Set ToBeGT = VerifyActual("IsGT", Expected, "Expected #{actual} to be > #{expected}.")
End Function
Public Function IsGT(Actual As Variant, Expected As Variant) As Boolean
    IsGT = (Actual > Expected)
End Function

'========================================================================================================================================'

Public Function ToBeGreaterThanOrEqualTo(Expected As Variant) As VBASpecExpectation
    Set ToBeGreaterThanOrEqualTo = VerifyActual("IsGTE", Expected, "Expected #{actual} to be greater than or equal to #{expected}.")
End Function
Public Function ToBeGTE(Expected As Variant) As VBASpecExpectation
    Set ToBeGTE = VerifyActual("IsGTE", Expected, "Expected #{actual} to be >= #{expected}.")
End Function
Public Function IsGTE(Actual As Variant, Expected As Variant) As Boolean
    IsGTE = (Actual >= Expected)
End Function

'========================================================================================================================================'

Public Function ToBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2) As VBASpecExpectation
    iSignificantFigures = SignificantFigures
    Set ToBeCloseTo = VerifyActual("IsCloseTo", Expected, "Expected #{actual} to be close to #{expected} (" & _
                                                          SignificantFigures & " significant figures).")
End Function
Public Function ToNotBeCloseTo(Expected As Variant, Optional SignificantFigures As Integer = 2) As VBASpecExpectation
    iSignificantFigures = SignificantFigures
    Set ToNotBeCloseTo = VerifyActual("IsCloseTo", Expected, "Expected #{actual} not to be close to #{expected} (" & _
                                                             SignificantFigures & " significant figures).", Inverse:=True)
End Function
Public Function IsCloseTo(Actual As Variant, Expected As Variant, Optional SignificantFigures As Integer = -1) As Boolean
    Dim sActual As String
    Dim sExpected As String
    If SignificantFigures < 1 Then SignificantFigures = iSignificantFigures
    sActual = Format$(Actual, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Actual >= 1, "e+0", "e-0"))
    sExpected = Format$(Expected, Left$("0.00000000000000", SignificantFigures + 1) & IIf(Expected >= 1, "e+0", "e-0"))
    
    IsCloseTo = (StrComp(sActual, sExpected, vbBinaryCompare) = 0)
End Function

'========================================================================================================================================'

Public Function ToInclude(Expected As Variant) As VBASpecExpectation
    Set ToInclude = VerifyActual("Includes", Expected, "Expected #{actual} to include #{expected}.")
End Function
Public Function ToNotInclude(Expected As Variant) As VBASpecExpectation
    Set ToNotInclude = VerifyActual("Includes", Expected, "Expected #{actual} not to include #{expected}.", Inverse:=True)
End Function
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

Public Function ToBeginWith(Expected As Variant) As VBASpecExpectation
    Set ToBeginWith = VerifyActual("BeginsWith", Expected, "Expected #{actual} to begin with #{expected}.")
End Function
Public Function ToNotBeginWith(Expected As Variant) As VBASpecExpectation
    Set ToNotBeginWith = VerifyActual("BeginsWith", Expected, "Expected #{actual} not to begin with #{expected}.", Inverse:=True)
End Function
Public Function BeginsWith(Actual As Variant, Expected As Variant) As Boolean
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
    Set ToEndWith = VerifyActual("EndsWith", Expected, "Expected #{actual} to end with #{expected}.")
End Function
Public Function ToNotEndWith(Expected As Variant) As VBASpecExpectation
    Set ToNotEndWith = VerifyActual("EndsWith", Expected, "Expected #{actual} not to end with #{expected}.", Inverse:=True)
End Function
Public Function EndsWith(Actual As Variant, Expected As Variant) As Boolean
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

Private Function IsLike(Var1 As Variant, Var2 As Variant) As Boolean
    If IsObject(Var1) Or IsObject(Var2) Then
        IsLike = IsSame(Var1, Var2)
    Else
        IsLike = IsEqual(Var1, Var2)
    End If
End Function

Private Function IsBoolean(Value As Variant) As Boolean
    IsBoolean = (VarType(Value) = vbBoolean)
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
Private Function CountEnumerable(Value As Variant) As Long
    Select Case VarType(Value)
    Case vbArray To vbArray + vbByte
        CountEnumerable = UBound(Value) - LBound(Value)
    Case vbObject
        If TypeOf Value Is Collection Then
            CountEnumerable = Value.Count
        End If
    End Select
End Function

'========================================================================================================================================'

Public Function NoError() As VBASpecExpectation
    Me.Actual = Err.Number
    Set NoError = VerifyActual("IsError", 0, "no error, got ##{actual}")
    Err.Clear
End Function
Public Function Error(Optional Expected As Variant) As VBASpecExpectation
    Me.Actual = Err.Number
    If IsMissing(Expected) Then
        Set Error = VerifyActual("IsError", Expected, "Expected an error, got none.")
    Else
        Set Error = VerifyActual("IsError", Expected, "Expected error ##{expected}, got ##{actual}.")
    End If
    Err.Clear
End Function
Public Function IsError(Actual As Variant, Expected As Variant) As Boolean
    bErrorChecked = True
    If IsMissing(Expected) Then
        IsError = (Err.Number <> 0)
    Else
        IsError = (Err.Number = Expected)
    End If
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
        If Len(sMessage) = 0 Then sMessage = IIf(Inverse, FAILURE_MSG_NEGATED, FAILURE_MSG)
        
        Dim bShowAddress As Boolean
        bShowAddress = (AnyValueAsText(Me.Actual) = AnyValueAsText(Expected))
        
        Dim bShowType As Boolean
        Select Case VarType(Expected)
        Case vbBoolean: bShowType = (CBool(Me.Actual) = CBool(Expected))
        Case vbInteger: bShowType = (CInt(Me.Actual) = CInt(Expected))
        Case vbDouble:  bShowType = (CDbl(Me.Actual) = CDbl(Expected))
        Case Else:      bShowType = (CStr(Me.Actual) = CStr(Expected))
        End Select
        
        If InStr(1, sMessage, ACTUALTYPE_INTERPOLATION, vbBinaryCompare) > 0 Then
            sMessage = Replace$(sMessage, ACTUALTYPE_INTERPOLATION, AnyValueAsText(Me.Actual, bShowAddress, True), Compare:=vbBinaryCompare)
        End If
        If InStr(1, sMessage, ACTUAL_INTERPOLATION, vbBinaryCompare) > 0 Then
            sMessage = Replace$(sMessage, ACTUAL_INTERPOLATION, AnyValueAsText(Me.Actual, bShowAddress, bShowType), Compare:=vbBinaryCompare)
        End If
        If InStr(1, sMessage, EXPECTED_INTERPOLATION, vbBinaryCompare) > 0 Then
            sMessage = Replace$(sMessage, EXPECTED_INTERPOLATION, AnyValueAsText(Expected, bShowAddress, bShowType), Compare:=vbBinaryCompare)
        End If
        sFailureMessage = sMessage
    End If
    
    Set VerifyActual = Me
End Function

Private Function AnyValueAsText(Value As Variant, Optional bShowAddress As Boolean = False, _
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

'========================================================================================================================================'

#If TESTING Then
Friend Sub UnitTest(Suite As VBASpecSuite)
    With Suite.Describe("VBASpecExpectation")
        Dim subject As VBASpecExpectation
        
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
                .Expect("Abc").ToEqual "Abc"
                .Expect("Abc").ToNotEqual "abc"
            End With
            With .It("works with Integers")
                .Expect(2048).ToEqual 2048
                .Expect(2048).ToNotEqual 2049
            End With
            With .It("works with Decimals")
                .Expect(3.14159265358979).ToEqual 3.14159265358979
                .Expect(3.14159265358979).ToNotEqual 3.1415926535898
            End With
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
                .Expect(subject).ToNotBe "aString"
                .Expect("aString").ToNotBe subject
                .Expect("aString").ToNotBe "aString"
            End With
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
        End With
        With .Describe(".ToBeNothing/.ToNotBeNothing")
            With .It("passes only when Actual is Nothing")
                .Expect(Nothing).ToBeNothing
                .Expect(Empty).ToNotBeNothing
                .Expect(Null).ToNotBeNothing
                .Expect().ToNotBeNothing
            End With
        End With
        With .Describe(".ToBeZero/.ToNotBeZero")
            With .It("passes if Actual equals 0")
                .Expect(0).ToBeZero
                .Expect(0#).ToBeZero
                .Expect("0").ToNotBeZero
                .Expect(0.000000000000001).ToNotBeZero
            End With
        End With
        
        With .Describe(".ToBeFalsy")
            With .It("passes when Actual evaluates to False")
                .Expect(False).ToBeFalsy
                .Expect(0).ToBeFalsy
                .Expect(0#).ToBeFalsy
            End With
        End With
        With .Describe(".ToBeFalse")
            With .It("passes when Actual is False")
                .Expect(False).ToBeFalse
            End With
        End With
        
        With .Describe(".ToBeTruthy")
            With .It("passes when Actual evaluates to True")
                .Expect(True).ToBeTruthy
                .Expect(1).ToBeTruthy
                .Expect(2).ToBeTruthy
                .Expect(0.1).ToBeTruthy
            End With
        End With
        With .Describe(".ToBeTrue")
            With .It("passes when Actual is True")
                .Expect(True).ToBeTrue
            End With
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
            With .Describe(".WasRaised")
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
            End With
        End With
        
        With .Describe(".Error")
            Set subject = New VBASpecExpectation
            
            With .Context("when no number given")
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
            With .Describe(".WasRaised")
                With .Context("when no error was raised")
                    errorTrigger = 1 / 1
                    
                    With .It("returns False")
                        .Expect(subject.Error.WasRaised).ToBeFalse
                    End With
                End With
                With .Context("when any error was raised")
                    errorTrigger = 1 / 0
                    
                    With .It("returns True")
                        .Expect(subject.Error.WasRaised).ToBeTrue
                    End With
                End With
                With .Context("when another error was raised")
                    errorTrigger = 1 / 0
                    
                    With .It("returns False")
                        .Expect(subject.Error(1).WasRaised).ToBeFalse
                    End With
                End With
            End With
        End With
        On Error GoTo 0
    End With
End Sub
#End If
