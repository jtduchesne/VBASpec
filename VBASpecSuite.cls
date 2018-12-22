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

Private iPassedCount As Integer
Private iFailedCount As Integer
Private iPendingCount As Integer

Private iLevel As Integer

'========================================================================================================================================'

Private Sub Class_Initialize()
    Set oOutput = New VBASpecFormOutput
End Sub
Private Sub Class_Terminate()
    Me.FinishSuite
    Set oOutput = Nothing
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

Public Property Get Immediate() As Boolean
    Immediate = (TypeOf oOutput Is VBASpecImmediateOutput)
End Property
Public Property Let Immediate(bValue As Boolean)
    If bValue Then Set oOutput = New VBASpecImmediateOutput
End Property

Public Property Get Verbose() As Boolean
    Verbose = (TypeOf oOutput Is VBASpecVerboseOutput)
End Property
Public Property Let Verbose(bValue As Boolean)
    If bValue Then Set oOutput = New VBASpecVerboseOutput
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
    If iFailedCount > 0 Then
        Result = Failed
    ElseIf iPassedCount > 0 Then
        Result = Passed
    Else
        Result = Pending
    End If
End Property

'========================================================================================================================================'

Public Function Describe(ByVal Description As String) As VBASpecGroup
    Set Describe = New VBASpecGroup
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

Friend Sub CreateSuite(FirstGroup As VBASpecGroup)
    oOutput.BeginSuite FirstGroup
    IncreaseLevel
End Sub
Friend Sub CreateGroup(Group As VBASpecGroup)
    oOutput.BeginGroup Group
    IncreaseLevel
End Sub

Friend Sub FinishGroup(Group As VBASpecGroup)
    DecreaseLevel
    If Group.Expectations.Count > 0 Then
        Select Case Group.Result
        Case Passed: iPassedCount = iPassedCount + 1
        Case Failed: iFailedCount = iFailedCount + 1
        Case Else
            iPendingCount = iPendingCount + 1
        End Select
    End If
    oOutput.EndGroup Group
End Sub
Friend Sub FinishSuite()
    DecreaseLevel
    oOutput.EndSuite Me
End Sub

'========================================================================================================================================'

#If TESTING Then
Private Function SilentVBASpecSuite() As VBASpecSuite
    Set SilentVBASpecSuite = New VBASpecSuite
    SilentVBASpecSuite.Silent = True
End Function
Friend Sub UnitTest(Suite As VBASpecSuite)
    With Suite.Describe("VBASpecSuite")
        Dim subject As VBASpecSuite
        
        With .Describe("#Count([Filter])")
            Set subject = SilentVBASpecSuite
            With subject.Describe("")
                .It("passes").Expect(True).ToBeTrue
                .It("passes").Expect(False).ToBeFalse
                .It("fails").Expect(True).ToBeFalse
            End With
            
            With .Context("without any Filter")
                With .It("returns the total number of examples")
                    .Expect(subject.Count).ToEqual 3
                End With
            End With
            With .Context("with a Filter")
                With .It("returns the filtered number of examples")
                    .Expect(subject.Count(Passed)).ToEqual 2
                    .Expect(subject.Count(Failed)).ToEqual 1
                    .Expect(subject.Count(Pending)).ToEqual 0
                End With
            End With
        End With
        
        With .Describe("#Result")
            With .Context("if there is no expectations")
                Set subject = SilentVBASpecSuite
                
                With .It("defaults to -Pending-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Pending
                End With
            End With
            With .Context("if all examples passes")
                Set subject = SilentVBASpecSuite
                With subject.Describe("")
                    .It.Expect(True).ToBeTrue
                    .It.Expect(False).ToBeFalse
                End With
                
                With .It("returns -Passed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Passed
                End With
            End With
            With .Context("if any examples fails")
                Set subject = SilentVBASpecSuite
                With subject.Describe("")
                    .It.Expect(True).ToBeTrue
                    .It.Expect(True).ToBeFalse
                End With
                
                With .It("returns -Failed-")
                    .Expect(subject.Result).ToEqual eVBASpecResult.Failed
                End With
            End With
        End With
    End With
End Sub
#End If
