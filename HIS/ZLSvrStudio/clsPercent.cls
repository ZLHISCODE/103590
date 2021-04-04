VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngPercentCount As Long
Private mlngLoop As Long
Private mintSvrPercent As Integer
Private mintPercent As Integer
Private mobjPbr As Object

Public Sub InitPercent(ByRef objPbr As Object, ByVal lngPercentCount As Long)
    mlngPercentCount = lngPercentCount
    mlngLoop = 0
    mintPercent = 0
    mintSvrPercent = 0
    Set mobjPbr = objPbr
    
    mobjPbr.value = 0
    mobjPbr.Max = 100
    DoEvents
    
End Sub


Public Sub LoopPercent()
    mlngLoop = mlngLoop + 1
    
            
    mintPercent = 100 * mlngLoop / mlngPercentCount
    If mintSvrPercent <> mintPercent Then
        mintSvrPercent = mintPercent
        mobjPbr.value = mintSvrPercent
        DoEvents
    End If
            
End Sub
