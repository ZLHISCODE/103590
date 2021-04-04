VERSION 5.00
Begin VB.Form frmEInvoiceParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   Icon            =   "frmEInvoiceParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4590
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd打印设置 
      Caption         =   "告知单打印设置(&P)"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   2130
      Width           =   2370
   End
   Begin VB.Frame fra 
      Caption         =   "开票点对码方式"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2385
      Begin VB.OptionButton Option对码方式 
         Caption         =   "按客户端和收费员"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option对码方式 
         Caption         =   "按收费员"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   1095
      End
      Begin VB.OptionButton Option对码方式 
         Caption         =   "按客户端"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   5
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   6
      Top             =   810
      Width           =   1100
   End
End
Attribute VB_Name = "frmEInvoiceParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln清除开票点对照 As Boolean
Private mlngSys As Long
Private mlngModule As Long
Private mstrPrivs As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intTmp As Integer, blnSetUp As Boolean
    Dim strSQL As String
    
    blnSetUp = InStr(1, mstrPrivs, ";参数设置;") > 0
    intTmp = IIf(Option对码方式(2).Value, 2, IIf(Option对码方式(1).Value, 1, 0))
    If fra.Tag <> intTmp Then
        zlDatabase.SetPara "开票点对码方式", intTmp, mlngSys, mlngModule, blnSetUp
        If mbln清除开票点对照 Then
            strSQL = "Zl_票据开票点对照_Update(3)"
            Call zlDatabase.ExecuteProcedure(strSQL, "票据开票点对照")
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitPara
End Sub

Private Sub InitPara()
    '初始化参数
    Dim intTmp As Integer, blnSetUp As Boolean
   
    mstrPrivs = ";" & GetPrivFunc(mlngSys, mlngModule) & ";"
    blnSetUp = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    intTmp = zlDatabase.GetPara("开票点对码方式", mlngSys, mlngModule, 1, Option对码方式, blnSetUp)
    fra.Tag = intTmp
    Option对码方式(intTmp).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln清除开票点对照 = False
End Sub

Private Sub Option对码方式_Click(Index As Integer)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    If Val(fra.Tag) = Index Then Exit Sub
    strSQL = "select  1 from 票据开票点对照 Where Rownum<2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then Exit Sub
    If MsgBox("你确认要更改【开票点对码方式】吗，更改了本参数将会清除【票据开票点对照】表中的数据？", _
       vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
        mbln清除开票点对照 = True
    Else
        Option对码方式(Val(fra.Tag)).Value = True
        zlControl.ControlSetFocus Option对码方式(Val(fra.Tag))
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, ByVal lngSys As Long, ByVal lngModule As Long)
    On Error GoTo errHandle
    mlngSys = lngSys: mlngModule = lngModule
    Me.Show 1, frmMain
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


