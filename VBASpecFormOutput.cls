VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecFormOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Pour activer l'attribut "VB_Creatable" a partir de la fenetre d'Execution:
'Workbooks("VBASpec.xlam").VBProject.VBComponents("VBASpecFormOutput").Properties("Instancing") = 5
Option Explicit

Implements IVBASpecOutput

Private Const sPending As String = "[-PENDING-] "
Private Const sFailed As String = " [-FAILED-] "
Private Const sExamplePrefix As String = "it "

Private frmOutput As VBASpecForm

Private StartedAt As Date
Private bRunning As Boolean

Private sBuffer As String
Private sOutput As String

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
        ResultTypeToString = "- "
    Case Failed
        ResultTypeToString = "X "
    Case Pending
        ResultTypeToString = "**"
    End Select
End Function

Private Function ElapsedTimeToString(dElapsedTime As Double) As String
    If dElapsedTime >= 60 Then
        ElapsedTimeToString = Int(dElapsedTime / 60) & "m" & Right(Strings.Format(dElapsedTime Mod 60, "00.000"), 6)
    Else
        ElapsedTimeToString = Strings.Format(dElapsedTime, "#0.000")
    End If
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

Private Sub PrintLine(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    AppendLineTo sOutput, sLine, iIndent
    Update
End Sub
Private Sub PrintLater(Optional sLine As String = vbNullString, Optional iIndent As Integer = 0)
    AppendLineTo sBuffer, sLine, iIndent
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
    frmOutput.Caption = FirstExampleGroup.Description
    If Not bRunning Then
        bRunning = True
        StartedAt = Timer
        
        sOutput = "Started at " & Now
        
        frmOutput.Show vbModeless
    End If
    PrintLine
    PrintLine FirstExampleGroup.Description
End Sub

Public Sub IVBASpecOutput_BeginExampleGroup(ExampleGroup As VBASpecExampleGroup)
    PrintLine ExampleGroup.Description, ExampleGroup.Level
End Sub

Public Sub IVBASpecOutput_BeginExample(Example As VBASpecExample)
End Sub

Public Sub IVBASpecOutput_EndExample(Example As VBASpecExample)
    PrintLine ResultTypeToString(Example.Result) & sExamplePrefix & Example.Description, Example.Level
    If Example.Pending Then
        PrintLater sPending & Example.CompleteDescription, Example.Level
    ElseIf Example.Result = Failed Then
        PrintLater sFailed & Example.CompleteDescription, Example.Level
        Dim Expectation As VBASpecExpectation
        For Each Expectation In Example.Expectations(Failed)
            PrintLater Expectation.FailureMessage, Example.Level + 8
        Next Expectation
    End If
End Sub

Public Sub IVBASpecOutput_EndExampleGroup(ExampleGroup As VBASpecExampleGroup)
    If Len(sBuffer) Then
        PrintLine sBuffer
        PrintLine
        sBuffer = vbNullString
    End If
End Sub

Public Sub IVBASpecOutput_EndSuite(Suite As VBASpecSuite)
    PrintSummary Suite
    bRunning = False
    Do Until Me.IVBASpecOutput_Done
        DoEvents
    Loop
End Sub

