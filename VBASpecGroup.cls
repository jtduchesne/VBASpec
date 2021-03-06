VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const sSpace As String = " "
Private Const sSharp As String = "#"
Private Const iSharp As Integer = 35
Private Const sPeriod As String = "."
Private Const iPeriod As Integer = 46

Private sDescription As String

Private clsSuite As VBASpecSuite
Private clsParent As VBASpecGroup

Private colExpectations As Collection

Private dictCaptures As Scripting.Dictionary

Private vResult As Variant
Private vSubject As Variant
Private bPending As Boolean

'========================================================================================================================================'

Private Sub Class_Initialize()
    Set colExpectations = New Collection
    bPending = True
End Sub
Private Sub Class_Terminate()
    If Not clsSuite Is Nothing Then
        clsSuite.FinishGroup Me
        Set clsSuite = Nothing
    End If
    Set clsParent = Nothing
    Set colExpectations = Nothing
    Set dictCaptures = Nothing
End Sub

Friend Property Get Description() As String
    Description = sDescription
End Property
Friend Property Let Description(sValue As String)
    sDescription = sValue
End Property
Friend Property Get CompleteDescription() As String
    Dim sParentDescription As String
    If Not clsParent Is Nothing Then
        sParentDescription = clsParent.CompleteDescription
    End If
    If Len(sDescription) > 0 Then
        If Len(sParentDescription) = 0 Then
            CompleteDescription = sDescription
        ElseIf AscW(sDescription) = iSharp Or AscW(sDescription) = iPeriod Then
            CompleteDescription = sParentDescription & sDescription
        Else
            If Right$(sParentDescription, Len(sDescription)) <> sDescription Then
                CompleteDescription = sParentDescription & sSpace & sDescription
            Else
                CompleteDescription = sParentDescription
            End If
        End If
    Else
        CompleteDescription = sParentDescription
    End If
End Property

Friend Property Get Level() As Integer
    If Not clsSuite Is Nothing Then
        Level = clsSuite.Level
    Else
        Level = -1
    End If
End Property

Friend Property Get Parent() As VBASpecGroup
    Set Parent = clsParent
End Property
Friend Property Set Parent(clsValue As VBASpecGroup)
    Set clsParent = clsValue
End Property

Friend Property Get Suite() As VBASpecSuite
    Set Suite = clsSuite
End Property
Friend Property Set Suite(clsValue As VBASpecSuite)
    Set clsSuite = clsValue
End Property

'========================================================================================================================================'

Friend Property Get Pending() As Boolean
    Pending = bPending
End Property

Friend Property Get Failed() As Boolean
    Dim Expectation As VBASpecExpectation
    For Each Expectation In colExpectations
        If Expectation.Failed Then
            Failed = True
            Exit Property
        End If
    Next Expectation
End Property

Friend Property Get Result() As eVBASpecResult
    If Pending Then
        Result = eVBASpecResult.Pending
        Exit Property
    End If
    If IsEmpty(vResult) Then
        vResult = IIf(Me.Failed, eVBASpecResult.Failed, eVBASpecResult.Passed)
    End If
    Result = vResult
End Property

Friend Property Get Expectations(Optional Filter As eVBASpecResult = [_All]) As Collection
    If Pending Then
        Set Expectations = colExpectations
        Exit Property
    End If
    
    Dim Expectation As VBASpecExpectation
    Select Case Filter
    Case eVBASpecResult.Passed
        Set Expectations = New Collection
        
        For Each Expectation In colExpectations
            If Expectation.Passed Then
                Expectations.Add Expectation
            End If
        Next Expectation
    Case eVBASpecResult.Failed
        Set Expectations = New Collection
        
        For Each Expectation In colExpectations
            If Expectation.Failed Then
                Expectations.Add Expectation
            End If
        Next Expectation
    Case Else
        Set Expectations = colExpectations
    End Select
End Property

'========================================================================================================================================'

Public Property Get subject()
    If IsEmpty(vSubject) Then
        If Not clsParent Is Nothing Then
            Assign clsParent.subject, Into:=subject
        End If
    Else
        Assign vSubject, Into:=subject
    End If
End Property
Public Property Let subject(vValue As Variant)
    vSubject = vValue
End Property
Public Property Set subject(vValue As Variant)
    Set vSubject = vValue
End Property

Public Property Get its(ByVal PropertyName As String) As Variant
    AssignProperty PropertyName, Of:=Me.subject, Into:=its
End Property
Public Property Let its(ByVal PropertyName As String, vValue As Variant)
    LetOrSetIts VbLet, PropertyName, vValue
End Property
Public Property Set its(ByVal PropertyName As String, vValue As Variant)
    LetOrSetIts VbSet, PropertyName, vValue
End Property
Private Sub LetOrSetIts(CallType As VbCallType, ByVal PropertyName As String, vValue As Variant)
    On Error Resume Next
    CallByName Me.subject, PropertyName, CallType, CVar(vValue)
    If Err.Number <> 0 Then
        Me.Expect.Fail AnyValueAsText(Me.subject) & "#" & PropertyName & ": " & Err.Description
        Err.Clear
    End If
End Sub

'========================================================================================================================================'

Public Function Capture(ParamArray PropertyNames()) As Variant
    If UBound(PropertyNames) < LBound(PropertyNames) Then
        Assign Me.subject, Into:=Capture
        If IsObject(Capture) Then
            Set Captures(vbNullString) = Capture
        Else
            Captures(vbNullString) = Capture
        End If
    Else
        Dim PropertyName
        For Each PropertyName In PropertyNames
            Assign its(PropertyName), Into:=Capture
            If IsObject(Capture) Then
                Set Captures(CStr(PropertyName)) = Capture
            Else
                Captures(CStr(PropertyName)) = Capture
            End If
        Next PropertyName
    End If
End Function

Friend Property Get Captured(Optional ByVal PropertyName As String) As Variant
    If Captures.Exists(PropertyName) Then
        Assign Captures(PropertyName), Into:=Captured
    Else
        If Not clsParent Is Nothing Then
            Assign clsParent.Captured(PropertyName), Into:=Captured
        End If
    End If
End Property
Friend Property Get Captures() As Scripting.Dictionary
    If dictCaptures Is Nothing Then Set dictCaptures = New Scripting.Dictionary
    Set Captures = dictCaptures
End Property

'========================================================================================================================================'

Public Function Expect_It() As VBASpecExpectation
    If (Captures.Count > 1) And Captures.Exists(vbNullString) Then
        Set Expect_It = Me.Expect(Me.subject).WithCapture(Me.Captured, Captures)
    Else
        Set Expect_It = Me.Expect(Me.subject).WithCapture(Me.Captured)
    End If
End Function

Public Function Expect_Its(ByVal PropertyName As String) As VBASpecExpectation
    If (AscW(PropertyName) = iSharp) Or (AscW(PropertyName) = iPeriod) Then
        PropertyName = Mid$(PropertyName, 2)
    End If
    On Error Resume Next
    Set Expect_Its = Me.Expect(its(PropertyName)).WithCapture(Me.Captured(PropertyName))
    If Err.Number = 438 Then
        Set Expect_Its = Me.Expect.Fail(Err.Description & ": " & AnyValueAsText(Me.subject) & sSharp & PropertyName)
    End If
End Function

'========================================================================================================================================'

Public Function Context(Optional ByVal Description As String) As VBASpecGroup
    Set Context = New VBASpecGroup
    With Context
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateGroup Context
End Function
Public Function Describe(Optional ByVal Description As String) As VBASpecGroup
    Set Describe = New VBASpecGroup
    With Describe
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateGroup Describe
End Function

Public Function It(Optional ByVal Description As String) As VBASpecGroup
    Set It = New VBASpecGroup
    With It
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateGroup It
End Function

Public Function Expect(Optional Actual As Variant) As VBASpecExpectation
    Dim Expectation As New VBASpecExpectation
    
    If IsObject(Actual) Then
        Set Expectation.Actual = Actual
    Else
        Expectation.Actual = Actual
    End If
    colExpectations.Add Expectation
    bPending = False
    
    Set Expect = Expectation
End Function

'========================================================================================================================================'

#If TESTING Then
Friend Sub UnitTest(Suite As VBASpecSuite)
    With Suite.Describe("VBASpecGroup")
        Dim oSuite As New VBASpecSuite
        oSuite.Silent = True
        
        Dim subject As VBASpecGroup
        
        With .Describe("#CompleteDescription")
            With .It("returns a pseudo-sentence made with the parent's description")
                .Expect(oSuite.Describe("An Object").Context("in a given context").CompleteDescription).ToEqual _
                    "An Object in a given context"
                .Expect(oSuite.Describe("An Object").Context("in a context").Context("and another one").CompleteDescription).ToEqual _
                    "An Object in a context and another one"
            End With
            With .It("removes the space before #properties and .functions")
                .Expect(oSuite.Describe("An Object").Describe("#prop").CompleteDescription).ToEqual "An Object#prop"
                .Expect(oSuite.Describe("An Object").Describe(".func()").CompleteDescription).ToEqual "An Object.func()"
            End With
        End With
        
        With .Describe("#Suite")
            With .It("Always returns the original #Suite")
                .Expect(oSuite.Describe("").Suite).ToBe oSuite
                .Expect(oSuite.Describe("").Context.Suite).ToBe oSuite
            End With
        End With
        
        With .Describe("#Pending")
            Set subject = New VBASpecGroup
            
            With .It("defaults to True")
                .Expect(subject.Pending).ToBeTrue
            End With
            
            With .Context("if there is a nested "".Describe""")
                Set subject = New VBASpecGroup
                subject.Describe "Something"
                
                With .It("returns False")
                    .Expect(subject.Pending).ToBeFalse
                End With
            End With
            With .Context("if there is a nested "".Context""")
                Set subject = New VBASpecGroup
                subject.Context "Something"
                
                With .It("returns False")
                    .Expect(subject.Pending).ToBeFalse
                End With
            End With
            With .Context("if there is a nested "".It""")
                Set subject = New VBASpecGroup
                subject.It "does something"
                
                With .It("returns False")
                    .Expect(subject.Pending).ToBeFalse
                End With
            End With
            With .Context("if there is an expectation")
                Set subject = New VBASpecGroup
                subject.Expect("Something").ToEqual "Something"
                
                With .It("returns False")
                    .Expect(subject.Pending).ToBeFalse
                End With
            End With
        End With
        
        With .Describe("#Failed")
            Set subject = New VBASpecGroup
            
            With .It("defaults to False")
                .Expect(subject.Failed).ToBeFalse
            End With
            
            With .Context("if all expectations passes")
                Set subject = New VBASpecGroup
                subject.Expect(True).ToBeTrue
                subject.Expect(False).ToBeFalse
                
                With .It("returns False")
                    .Expect(subject.Failed).ToBeFalse
                End With
            End With
            With .Context("if any expectation fails")
                Set subject = New VBASpecGroup
                subject.Expect(True).ToBeTrue
                subject.Expect(True).ToBeFalse
                
                With .It("returns True")
                    .Expect(subject.Failed).ToBeTrue
                End With
            End With
            
            With .It("does not cache the result")
                Set subject = New VBASpecGroup
                
                subject.Expect(True).ToBeTrue
                .Expect(subject.Failed).ToBeFalse
                subject.Expect(True).ToBeFalse
                .Expect(subject.Failed).ToBeTrue
            End With
        End With
        
        With .Describe("#Result")
            Set subject = New VBASpecGroup
            
            With .Context("if there is no expectations")
                With .It("defaults to -Pending-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Pending
                End With
            End With
            With .Context("if all expectations passes")
                Set subject = New VBASpecGroup
                subject.Expect(True).ToBeTrue
                subject.Expect(False).ToBeFalse
                
                With .It("returns -Passed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Passed
                End With
            End With
            With .Context("if any expectation fails")
                Set subject = New VBASpecGroup
                subject.Expect(True).ToBeTrue
                subject.Expect(True).ToBeFalse
                
                With .It("returns -Failed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Failed
                End With
            End With
        End With
        
        With .Describe("#Expectations([Filter])")
            Set subject = New VBASpecGroup
            Dim oPassed As VBASpecExpectation
            Dim oFailed As VBASpecExpectation
            With subject
                Set oPassed = .Expect(True).ToBeTrue
                Set oFailed = .Expect(True).ToBeFalse
            End With
            
            With .Context("without any filters")
                With .It("returns all expectations")
                    .Expect(subject.Expectations).ToInclude oPassed
                    .Expect(subject.Expectations).ToInclude oFailed
                End With
            End With
            With .Context("if Filter = Passed")
                With .It("returns only passed expectations")
                    .Expect(subject.Expectations(eVBASpecResult.Passed)).ToInclude oPassed
                    .Expect(subject.Expectations(eVBASpecResult.Passed)).ToNotInclude oFailed
                End With
            End With
            With .Context("if Filter = Failed")
                With .It("returns only failed expectations")
                    .Expect(subject.Expectations(eVBASpecResult.Failed)).ToNotInclude oPassed
                    .Expect(subject.Expectations(eVBASpecResult.Failed)).ToInclude oFailed
                End With
            End With
        End With
        
        With .Describe(".Capture(PropertyNames)")
            With .It("takes a snapshot of a Property's value at that point")
                Set .subject = New Collection
                .subject.Add "something"
                .Capture "Count"
                .subject.Add "something"
                
                .Expect(.Captured("Count")).ToEqual .its("Count") - 1
            End With
            
            With .It("stores the value in current Group's context")
                Set .subject = New Collection
                .subject.Add "something"
                .Capture "Count"
                .subject.Add "something else"
                
                With .Context("but")
                    With .It("is accessible in children")
                        .Expect(.Captured("Count")).ToEqual .its("Count") - 1
                    End With
                End With
                With .Context("and")
                    With .It("cannot be changed by children")
                        With .Context("children")
                            .Capture "Count"
                            .Expect(.Captured("Count")).ToEqual .its("Count")
                        End With
                        .Expect(.Captured("Count")).ToEqual .its("Count") - 1
                    End With
                    .subject.Add "some more"
                End With
                
                .Expect(.Captured("Count")).ToEqual .its("Count") - 2
            End With
        End With
        
        With .Describe(".Context(Description)")
            Set subject = New VBASpecGroup
        
            With .It("returns an ExampleGroup")
                .Expect(subject.Context("")).ToBeA "VBASpecGroup"
            End With
            With .It("always returns a new Group")
                .Expect(subject.Context("")).ToNotBe subject
                .Expect(subject.Context("")).ToNotBe subject.Context("")
            End With
            With .It("sets new Group's description")
                .Expect(subject.Context("New Description").Description).ToEqual "New Description"
            End With
            With .It("works without a description")
                .Expect(subject.Context.Description).ToBeEmpty
            End With
        End With
        
        With .Describe(".Describe(Description)")
            Set subject = New VBASpecGroup
        
            With .It("returns a Group")
                .Expect(subject.Describe("")).ToBeA "VBASpecGroup"
            End With
            With .It("always returns a new Group")
                .Expect(subject.Describe("")).ToNotBe subject
                .Expect(subject.Describe("")).ToNotBe subject.Describe("")
            End With
            With .It("sets new Group's description")
                .Expect(subject.Describe("New Description").Description).ToEqual "New Description"
            End With
            With .It("works without a description")
                .Expect(subject.Describe.Description).ToBeEmpty
            End With
        End With
        
        With .Describe(".It(Description)")
            Set subject = New VBASpecGroup
        
            With .It("returns a Group")
                .Expect(subject.It("")).ToBeA "VBASpecGroup"
            End With
            With .It("always returns a new Group")
                .Expect(subject.It("")).ToNotBe subject.It("")
            End With
            With .It("sets new Group's description")
                .Expect(subject.It("New Description").Description).ToEqual "New Description"
            End With
            With .It("works without a description")
                .Expect(subject.It.Description).ToBeEmpty
            End With
        End With
        
        With .Describe(".Expect(Actual)")
            Set subject = New VBASpecGroup
            
            With .It("returns an VBASpecExpectation")
                .Expect(subject.Expect).ToBeA "VBASpecExpectation"
            End With
            With .It("accepts an expression")
                .Expect(subject.Expect(2).Actual).ToEqual 2
                .Expect(subject.Expect(3.5).Actual).ToEqual 3.5
                .Expect(subject.Expect("Eight").Actual).ToEqual "Eight"
            End With
            With .It("accepts an object")
                .Expect(subject.Expect(subject).Actual).ToBe subject
            End With
        End With
    End With
End Sub
#End If
