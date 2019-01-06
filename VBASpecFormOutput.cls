VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecFormOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVBASpecOutput

Private Const sPending As String = "[-PENDING-] "
Private Const sFailed As String = " [-FAILED-] "
Private Const sError As String = "  [-ERROR-] "

Private frmOutput As VBASpecForm

Private StartedAt As Date
Private bRunning As Boolean

Private sBuffer As String
Private sLabel As String
Private sOutput As String
Private iStatusCount As Integer

Private bDone As Boolean

Private Sub Class_Initialize()
    Set frmOutput = New VBASpecForm
    Set frmOutput.Owner = Me
End Sub

Private Sub Class_Terminate()
    Set frmOutput = Nothing
End Sub

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

Private Sub Update()
    frmOutput.TextBox.Text = sOutput
    frmOutput.Repaint
End Sub

Private Sub AppendLineTo(ByRef sBuffer As String, sLine As String, iIndent As Integer)
    If iIndent > 0 Then
        sBuffer = sBuffer & vbNewLine & Space(iIndent * 2) & sLine
    Else
        sBuffer = sBuffer & vbNewLine & sLine
    End If
End Sub

Private Sub StartStatus(Optional Label As String)
    sLabel = Label
    PrintLine Label & " ["
End Sub
Private Sub StopStatus()
    sOutput = sOutput & "] " & iStatusCount
    iStatusCount = 0
    PrintBuffer
End Sub
Private Sub UpdateStatus(ResultType As eVBASpecResult)
    sOutput = sOutput & ResultTypeToString(ResultType)
    iStatusCount = iStatusCount + 1
    If (iStatusCount Mod 100) = 0 Then
        sOutput = sOutput & vbNewLine & Space(Len(sLabel) + 2)
    End If
End Sub

Private Sub PrintLine(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    AppendLineTo sOutput, sLine, iIndent
    Update
End Sub
Private Sub PrintLater(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    AppendLineTo sBuffer, sLine, iIndent
End Sub
Private Sub PrintBuffer(Optional iIndent As Integer = 0)
    If Len(sBuffer) Then
        If iIndent > 0 Then
            PrintLine Space(iIndent * 2) & Replace$(sBuffer, vbNewLine, vbNewLine + Space(iIndent * 2), , Len(sBuffer) - 1, vbBinaryCompare), iIndent
        Else
            PrintLine sBuffer, iIndent
        End If
        PrintLine
        sBuffer = vbNullString
    End If
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
    frmOutput.Caption = FirstGroup.Description
    If Not bRunning Then
        bRunning = True
        StartedAt = Timer
        
        sOutput = "Started at " & Now & vbNewLine
        
        frmOutput.Show vbModeless
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
            PrintLater sPending & Group.CompleteDescription, 1
        ElseIf Group.Result = Failed Then
            PrintLater sFailed & Group.CompleteDescription, 1
            Dim Expectation As VBASpecExpectation
            For Each Expectation In Group.Expectations(Failed)
                PrintLater Expectation.FailureMessage, 8
            Next Expectation
            Err.Clear
        End If
    End If
    PrintErrorIfAny
End Sub

Public Sub IVBASpecOutput_EndSuite(Suite As VBASpecSuite)
    StopStatus
    PrintSummary Suite
    bRunning = False
    Do Until Me.IVBASpecOutput_Done
        DoEvents
    Loop
End Sub
