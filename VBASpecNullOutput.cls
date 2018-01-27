VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBASpecNullOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVBASpecOutput

Public Property Get IVBASpecOutput_Running() As Boolean
End Property
Public Property Let IVBASpecOutput_Running(bValue As Boolean)
End Property

Public Property Get IVBASpecOutput_Done() As Boolean
End Property
Public Property Let IVBASpecOutput_Done(bValue As Boolean)
End Property

Public Sub IVBASpecOutput_BeginSuite(FirstExampleGroup As VBASpecExampleGroup)
End Sub
Public Sub IVBASpecOutput_BeginExampleGroup(ExampleGroup As VBASpecExampleGroup)
End Sub
Public Sub IVBASpecOutput_BeginExample(Example As VBASpecExample)
End Sub

Public Sub IVBASpecOutput_EndSuite(Suite As VBASpecSuite)
End Sub
Public Sub IVBASpecOutput_EndExampleGroup(ExampleGroup As VBASpecExampleGroup)
End Sub
Public Sub IVBASpecOutput_EndExample(Example As VBASpecExample)
End Sub
