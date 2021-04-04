VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form FrmMutilSelect 
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "FrmMutilSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   9810
   StartUpPosition =   1  '所有者中心
   Begin MSDataGridLib.DataGrid DbgMutilSelect 
      Align           =   3  'Align Left
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   8599
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMutilSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents gRecCommon As ADODB.Recordset  '公共绑定记录集
Attribute gRecCommon.VB_VarHelpID = -1
Public gStrHideCol As String '列宽值
Public strCaption As String '标题
Public BlnSelect As Boolean '选择标志
Public FrmWidth As Single
Public FrmHeight As Single
Private intCol As Integer '列
Private gStrColWidth As String '列宽
Private HideCol

Private Sub DbgMutilSelect_DblClick()
    BlnSelect = True
    Me.Hide
End Sub

Private Sub DbgMutilSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then DbgMutilSelect_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        BlnSelect = False
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Set DbgMutilSelect.DataSource = gRecCommon
    
    HideCol = Split(gStrHideCol, ",")
    For intCol = 0 To Me.DbgMutilSelect.Columns.Count - 1
        DbgMutilSelect.Columns(intCol).Width = HideCol(intCol)
    Next
    
    Me.Caption = strCaption
    Me.Width = FrmWidth
    Me.Height = FrmHeight
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    With Me.DbgMutilSelect
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gRecCommon = Nothing
End Sub


