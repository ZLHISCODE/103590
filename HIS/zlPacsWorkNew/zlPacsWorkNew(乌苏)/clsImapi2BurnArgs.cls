VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImapi2BurnArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ElapsedTime As Long
Public FreeSystemBuffer As Long
Public LastReadLba As Long
Public LastWrittenLba As Long
Public RemainingTime As Long
Public SectorCount As Long
Public StartLba As Long
Public TotalSystemBuffer As Long
Public TotalTime As Long
Public UsedSystemBuffer As Long
