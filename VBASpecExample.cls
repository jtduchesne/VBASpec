VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const sSpace As String = " "

Private sDescription As String
Private sParentDescription As String

Private colExpectations As Collection

Private clsParent As VBASpecExampleGroup
Private clsSuite As VBASpecSuite

Private bResult As Variant

'========================================================================================================================================'

Private Sub Class_Initialize()
    Set colExpectations = New Collection
End Sub
Private Sub Class_Terminate()
    If Not clsSuite Is Nothing Then
        clsSuite.FinishExample Me
        Set clsSuite = Nothing
    End If
    Set clsParent = Nothing
    Set colExpectations = Nothing
End Sub

Friend Property Get Description() As String
    Description = sDescription
End Property
Friend Property Let Description(sValue As String)
    sDescription = sValue
End Property
Friend Property Get CompleteDescription() As String
    CompleteDescription = sParentDescription & sDescription
End Property

Friend Property Get Level() As Integer
    If Not clsSuite Is Nothing Then
        Level = clsSuite.Level
    Else
        Level = -1
    End If
End Property

Friend Property Get Parent() As VBASpecExampleGroup
    Set Parent = clsParent
End Property
Friend Property Set Parent(clsValue As VBASpecExampleGroup)
    Set clsParent = clsValue
    If Not clsValue Is Nothing Then
        sParentDescription = clsValue.CompleteDescription & sSpace
    End If
End Property

Friend Property Get Suite() As VBASpecSuite
    Set Suite = clsSuite
End Property
Friend Property Set Suite(clsValue As VBASpecSuite)
    Set clsSuite = clsValue
End Property

'========================================================================================================================================'

Friend Property Get Pending() As Boolean
    Pending = (colExpectations.Count = 0)
End Property

Friend Property Get Result() As eVBASpecResult
    If Pending Then
        Result = eVBASpecResult.Pending
        Exit Property
    End If
    If Not bResult Then
        Dim Expectation As VBASpecExpectation
        For Each Expectation In colExpectations
            If Expectation.Failed Then
                bResult = eVBASpecResult.Failed
                GoTo Done
            End If
        Next Expectation
        bResult = eVBASpecResult.Passed
    End If
    
Done:
    Result = bResult
End Property

Friend Property Get Expectations(Optional Filter As eVBASpecResult = [_All]) As Collection
    If Pending Then
        Set Expectations = colExpectations
        Exit Property
    End If
    
    Dim Expectation As VBASpecExpectation
    Select Case Filter
    Case Passed
        Set Expectations = New Collection
        
        For Each Expectation In colExpectations
            If Expectation.Passed Then
                Expectations.Add Expectation
            End If
        Next Expectation
    Case Failed
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

Public Function Expect(Optional Actual As Variant) As VBASpecExpectation
    Dim Expectation As New VBASpecExpectation
    
    If IsObject(Actual) Then
        Set Expectation.Actual = Actual
    Else
        Expectation.Actual = Actual
    End If
    colExpectations.Add Expectation
    
    Set Expect = Expectation
End Function

'========================================================================================================================================'

#If TESTING Then
Friend Sub UnitTest(Suite As VBASpecSuite)
    With Suite.Describe("VBASpecExample")
        Dim oSuite As New VBASpecSuite
        oSuite.Silent = True
        
        Dim subject As VBASpecExample
        
        With .Describe("#CompleteDescription")
            Set subject = oSuite.Describe("An Object").It("returns something")
            
            With .It("returns a pseudo-sentence made with the parent's description")
                .Expect(subject.CompleteDescription).ToEqual "An Object returns something"
            End With
        End With
        
        With .Describe("#Suite")
            With .It("Always return the original #Suite")
                .Expect(oSuite.Describe("").It("").Suite).ToBe oSuite
                .Expect(oSuite.Describe("").Context("").It("").Suite).ToBe oSuite
            End With
        End With
        
        With .Describe("#Pending")
            Set subject = New VBASpecExample
            
            With .Context("if there is no expectations")
                With .It("returns True")
                    .Expect(subject.Pending).ToBeTrue
                End With
            End With
            With .Context("if there is an expectation")
                subject.Expect("Something").ToEqual "Something"
                
                With .It("returns False")
                    .Expect(subject.Pending).ToBeFalse
                End With
            End With
        End With
        
        With .Describe("#Result")
            Set subject = New VBASpecExample
            
            With .Context("if there is no expectations")
                With .It("defaults to -Pending-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Pending
                End With
            End With
            With .Context("if all expectations passes")
                Set subject = New VBASpecExample
                subject.Expect(True).ToBeTrue
                subject.Expect(False).ToBeFalse
                
                With .It("returns -Passed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Passed
                End With
            End With
            With .Context("if any expectation fails")
                Set subject = New VBASpecExample
                subject.Expect(True).ToBeTrue
                subject.Expect(True).ToBeFalse
                
                With .It("returns -Failed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Failed
                End With
            End With
        End With
        
        With .Describe("#Expectations([Filter])")
            Set subject = New VBASpecExample
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
                    .Expect(subject.Expectations(Passed)).ToInclude oPassed
                    .Expect(subject.Expectations(Passed)).ToNotInclude oFailed
                End With
            End With
            With .Context("if Filter = Failed")
                With .It("returns only failed expectations")
                    .Expect(subject.Expectations(Failed)).ToNotInclude oPassed
                    .Expect(subject.Expectations(Failed)).ToInclude oFailed
                End With
            End With
        End With
        
        With .Describe(".Expect(Actual)")
            Set subject = New VBASpecExample
            
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
