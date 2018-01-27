VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Pour activer l'attribut "VB_Creatable" a partir de la fenetre d'Execution:
'Workbooks("VBASpec.xlam").VBProject.VBComponents("VBASpecSuite").Properties("Instancing") = 5
Option Explicit

Public Enum eVBASpecResult
    Failed = 1
    Passed
    Pending
    [_All]
End Enum

Private oOutput As IVBASpecOutput

Private colExamples As Collection

Private iPassedCount As Integer
Private iFailedCount As Integer
Private iPendingCount As Integer

Private iLevel As Integer

'========================================================================================================================================'

Private Sub Class_Initialize()
    Set colExamples = New Collection
    Set oOutput = New VBASpecImmediateOutput
End Sub
Private Sub Class_Terminate()
    Me.FinishSuite
    Set oOutput = Nothing
    Set colExamples = Nothing
End Sub

Friend Property Get Level() As Integer
    Level = iLevel
End Property

Friend Property Get Silent() As Boolean
    Silent = (TypeOf oOutput Is VBASpecNullOutput)
End Property
Friend Property Let Silent(bValue As Boolean)
    If bValue Then Set oOutput = New VBASpecNullOutput
End Property

Public Property Get Output() As IVBASpecOutput
    Set Output = oOutput
End Property
Public Property Set Output(oValue As IVBASpecOutput)
    Set oOutput = oValue
End Property

'========================================================================================================================================'

Public Property Get Count(Optional Filter As eVBASpecResult = [_All]) As Integer
    Select Case Filter
    Case Passed: Count = iPassedCount
    Case Failed: Count = iFailedCount
    Case Pending: Count = iPendingCount
    Case Else: Count = iPassedCount + iFailedCount + iPendingCount
    End Select
End Property

Public Property Get Result() As eVBASpecResult
    If Me.Examples.Count > 0 Then
        Dim Example As VBASpecExample
        For Each Example In colExamples
            If (Example.Result < Result) Then Result = Example.Result
            If Result = Failed Then Exit Property
        Next Example
    Else
        Result = Pending
    End If
End Property

Public Property Get Examples(Optional Filter As eVBASpecResult = [_All]) As Collection
    Select Case Filter
    Case Passed, Failed, Pending
        Set Examples = New Collection
        
        Dim Example As VBASpecExample
        For Each Example In colExamples
            If Example.Result = Filter Then
                Examples.Add Example
            End If
        Next Example
    Case Else
        Set Examples = colExamples
    End Select
End Property

'========================================================================================================================================'

Public Function Describe(ByVal Description As String) As VBASpecExampleGroup
    Set Describe = New VBASpecExampleGroup
    With Describe
        .Description = Description
        Set .Parent = Nothing
        Set .Suite = Me
    End With
    Me.CreateSuite Describe
End Function

'========================================================================================================================================'

Private Sub IncreaseLevel()
    iLevel = iLevel + 1
End Sub
Private Sub DecreaseLevel()
    iLevel = iLevel - 1
End Sub

Friend Sub CreateSuite(FirstExampleGroup As VBASpecExampleGroup)
    oOutput.BeginSuite FirstExampleGroup
    IncreaseLevel
End Sub
Friend Sub CreateExampleGroup(ExampleGroup As VBASpecExampleGroup)
    oOutput.BeginExampleGroup ExampleGroup
    IncreaseLevel
End Sub
Friend Sub CreateExample(Example As VBASpecExample)
    oOutput.BeginExample Example
    IncreaseLevel
End Sub

Friend Sub FinishExample(Example As VBASpecExample)
    DecreaseLevel
    Select Case Example.Result
    Case Passed: iPassedCount = iPassedCount + 1
    Case Failed: iFailedCount = iFailedCount + 1
    Case Else
        iPendingCount = iPendingCount + 1
    End Select
    colExamples.Add Example
    
    oOutput.EndExample Example
End Sub
Friend Sub FinishExampleGroup(ExampleGroup As VBASpecExampleGroup)
    DecreaseLevel
    oOutput.EndExampleGroup ExampleGroup
End Sub
Friend Sub FinishSuite()
    DecreaseLevel
    oOutput.EndSuite Me
End Sub
