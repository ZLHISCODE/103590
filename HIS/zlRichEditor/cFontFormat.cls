VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFontFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public AllCaps As Boolean
Public BackColor As OLE_COLOR
Public Bold As Boolean
Public ForeColor As OLE_COLOR
Public Hidden As Boolean
Public Italic As Boolean
Public Link As Boolean
Public Name As String
Public Protected As Boolean
Public Size As Single
Public Strikethrough As Boolean
Public Subscript As Boolean
Public Superscript As Boolean
Public Underline As UnderLineEnum
