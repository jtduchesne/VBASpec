VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecImmediateOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Pour activer l'attribut "VB_Creatable" a partir de la fenetre d'Execution:
'Workbooks("VBASpec.xlam").VBProject.VBComponents("VBASpecImmediateOutput").Properties("Instancing") = 5
Option Explicit

Implements IVBASpecOutput

Private Const sPending As String = "[-PENDING-] "
Private Const sFailed As String = " [-FAILED-] "

Private StartedAt As Date
Private bRunning As Boolean

Private sStatus As String

Private bDone As Boolean

'========================================================================================================================================'

Private Function ResultTypeToString(ResultType As eVBASpecResult) As String
    Select Case ResultType
    Case Passed
        ResultTypeToString = "-"
    Case Failed
        ResultTypeToString = "X"
    Case Pending
        ResultTypeToString = "P"
    End Select
End Function

Private Function ElapsedTimeToString(dElapsedTime As Double) As String
    If dElapsedTime >= 60 Then
        ElapsedTimeToString = Int(dElapsedTime / 60) & "m" & Right(Strings.Format(dElapsedTime Mod 60, "00.000"), 6)
    Else
        ElapsedTimeToString = Strings.Format(dElapsedTime, "#0.000")
    End If
End Function

Private Sub StartStatus(Optional sLabel As String = vbNullString)
    sStatus = sLabel & " ["
End Sub
Private Sub StopStatus()
    UpdateStatus "]"
    PrintStatus
End Sub
Private Sub UpdateStatus(sValue As String)
    sStatus = sStatus & sValue
End Sub

Private Sub PrintStatus()
    Debug.Print sStatus
End Sub
Private Sub PrintLine(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    If iIndent > 0 Then sLine = Space(iIndent * 2) & sLine
    Debug.Print sLine
End Sub

Private Sub PrintSummary(Suite As VBASpecSuite)
    Dim iTotal As Integer
    Dim iPassed As Integer
    Dim iFailed As Integer
    Dim iPending As Integer
    iPassed = Suite.Count(eVBASpecResult.Passed)
    iFailed = Suite.Count(eVBASpecResult.Failed)
    iPending = Suite.Count(eVBASpecResult.Pending)
    iTotal = Suite.Count
    
    Dim sMessage As String
    If iFailed > 0 Then
        sMessage = "FAIL (" & iFailed & " of " & iTotal & " failed"
    Else
        sMessage = "PASS (" & iPassed & " of " & iTotal & " passed"
    End If
    If iPending > 0 Then
        sMessage = sMessage & ", " & iPending & " pending)"
    Else
        sMessage = sMessage & ")"
    End If
    
    PrintLine
    PrintLine "= " & sMessage & " = " & ElapsedTimeToString(Timer - StartedAt) & "s elapsed ="
End Sub

'========================================================================================================================================'

Public Property Get IVBASpecOutput_Running() As Boolean
    IVBASpecOutput_Running = bRunning
End Property
Public Property Let IVBASpecOutput_Running(bValue As Boolean)
    bRunning = bValue
End Property

Public Property Get IVBASpecOutput_Done() As Boolean
    IVBASpecOutput_Done = bDone
End Property
Public Property Let IVBASpecOutput_Done(bValue As Boolean)
    bDone = bValue
End Property

Public Sub IVBASpecOutput_BeginSuite(FirstExampleGroup As VBASpecExampleGroup)
    If Not bRunning Then
        bRunning = True
        StartedAt = Timer
        PrintLine "Started at " & Now
    Else
        StopStatus
    End If
    StartStatus FirstExampleGroup.Description
End Sub

Public Sub IVBASpecOutput_BeginExampleGroup(ExampleGroup As VBASpecExampleGroup)
End Sub
Public Sub IVBASpecOutput_BeginExample(Example As VBASpecExample)
End Sub

Public Sub IVBASpecOutput_EndExample(Example As VBASpecExample)
    UpdateStatus ResultTypeToString(Example.Result)
    If Example.Pending Then
        PrintLine sPending & Example.CompleteDescription
    ElseIf Example.Result = Failed Then
        PrintLine sFailed & Example.CompleteDescription
        Dim Expectation As VBASpecExpectation
        For Each Expectation In Example.Expectations(Failed)
            PrintLine Expectation.FailureMessage, 8
        Next Expectation
    End If
End Sub

Public Sub IVBASpecOutput_EndExampleGroup(ExampleGroup As VBASpecExampleGroup)
    If ExampleGroup.Pending Then
        PrintLine ExampleGroup.CompleteDescription & sPending
    End If
End Sub

Public Sub IVBASpecOutput_EndSuite(Suite As VBASpecSuite)
    StopStatus
    PrintSummary Suite
    Me.IVBASpecOutput_Done = True
End Sub
