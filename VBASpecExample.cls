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
