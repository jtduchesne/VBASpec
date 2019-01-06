VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecImmediateOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVBASpecOutput

Private Const sPending As String = "[-PENDING-] "
Private Const sFailed As String = " [-FAILED-] "
Private Const sError As String = "  [-ERROR-] "

Private StartedAt As Date
Private bRunning As Boolean

Private sLabel As String
Private sStatus As String
Private iStatusCount As Integer

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

Private Sub StartStatus(Optional Label As String)
    sLabel = Label
    sStatus = Label & " ["
End Sub
Private Sub StopStatus()
    sStatus = sStatus & "] " & iStatusCount
    iStatusCount = 0
    PrintLine sStatus
End Sub
Private Sub UpdateStatus(ResultType As eVBASpecResult)
    sStatus = sStatus & ResultTypeToString(ResultType)
    iStatusCount = iStatusCount + 1
    If (iStatusCount Mod 100) = 0 Then
        PrintLine sStatus
        sStatus = Space(Len(sLabel) + 2)
    End If
End Sub

Private Sub PrintLine(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    If iIndent > 0 Then sLine = Space(iIndent * 2) & sLine
    Debug.Print sLine
End Sub

Private Sub PrintErrorIfAny()
    If Err.Number <> 0 Then
        PrintLine sError & "#" & Err.Number & " " & Err.Description, 2
        Err.Clear
    End If
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

Public Sub IVBASpecOutput_BeginSuite(FirstGroup As VBASpecGroup)
    If Not bRunning Then
        bRunning = True
        StartedAt = Timer
        PrintLine "Started at " & Now
    Else
        StopStatus
    End If
    StartStatus FirstGroup.Description
End Sub

Public Sub IVBASpecOutput_BeginGroup(Group As VBASpecGroup)
End Sub

Public Sub IVBASpecOutput_EndGroup(Group As VBASpecGroup)
    If Group.Expectations.Count > 0 Or Group.Pending Then
        UpdateStatus Group.Result
        If Group.Result = Pending Then
            PrintLine sPending & Group.CompleteDescription
        ElseIf Group.Result = Failed Then
            PrintLine sFailed & Group.CompleteDescription
            Dim Expectation As VBASpecExpectation
            For Each Expectation In Group.Expectations(Failed)
                PrintLine Expectation.FailureMessage, 8
            Next Expectation
            Err.Clear
        End If
    End If
    PrintErrorIfAny
End Sub

Public Sub IVBASpecOutput_EndSuite(Suite As VBASpecSuite)
    StopStatus
    PrintSummary Suite
    Me.IVBASpecOutput_Done = True
End Sub
