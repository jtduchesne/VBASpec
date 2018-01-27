VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecExampleGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const sSpace As String = " "

Private sDescription As String
Private sParentDescription As String

Private colExamples As Collection

Private clsSuite As VBASpecSuite
Private clsParent As VBASpecExampleGroup

Private bPending As Boolean

'========================================================================================================================================'

Private Sub Class_Initialize()
    Set colExamples = New Collection
    bPending = True
End Sub
Private Sub Class_Terminate()
    If Not clsSuite Is Nothing Then
        clsSuite.FinishExampleGroup Me
        Set clsSuite = Nothing
    End If
    Set clsParent = Nothing
    Set colExamples = Nothing
End Sub

Friend Property Get Description() As String
    Description = sDescription
End Property
Friend Property Let Description(sValue As String)
    sDescription = sValue
End Property
Friend Property Get CompleteDescription() As String
    If Len(sDescription) > 0 Then
        If Len(sParentDescription) = 0 Or AscW(sDescription) = 35 Or AscW(sDescription) = 46 Then '"#" Or "."
            CompleteDescription = sParentDescription & sDescription
        Else
            CompleteDescription = sParentDescription & sSpace & sDescription
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

Friend Property Get Parent() As VBASpecExampleGroup
    Set Parent = clsParent
End Property
Friend Property Set Parent(clsValue As VBASpecExampleGroup)
    Set clsParent = clsValue
    If Not clsValue Is Nothing Then
        sParentDescription = clsValue.CompleteDescription
    End If
End Property

Friend Property Get Suite() As VBASpecSuite
    Set Suite = clsSuite
End Property
Friend Property Set Suite(clsValue As VBASpecSuite)
    Set clsSuite = clsValue
End Property

Friend Property Get Pending() As Boolean
    Pending = bPending
End Property

'========================================================================================================================================'

Public Function Context(ByVal Description As String) As VBASpecExampleGroup
    Set Context = New VBASpecExampleGroup
    With Context
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateExampleGroup Context
End Function
Public Function Describe(ByVal Description As String) As VBASpecExampleGroup
    Set Describe = New VBASpecExampleGroup
    With Describe
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateExampleGroup Describe
End Function

Public Function It(ByVal Description As String) As VBASpecExample
    Set It = New VBASpecExample
    With It
        .Description = Description
        Set .Parent = Me
        Set .Suite = clsSuite
    End With
    bPending = False
    If Not clsSuite Is Nothing Then clsSuite.CreateExample It
End Function
