VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm医保帐户查找 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找医保帐户"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frm医保帐户查找.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4860
      TabIndex        =   17
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4860
      TabIndex        =   18
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4860
      TabIndex        =   16
      Top             =   210
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "查找条件"
      Height          =   2835
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   2340
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65929219
         CurrentDate     =   37405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   757
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1154
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1551
         Width           =   2955
      End
      Begin VB.CommandButton cmd单位 
         Caption         =   "…"
         Height          =   300
         Left            =   3990
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1950
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2970
         TabIndex        =   15
         Top             =   2340
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   65929219
         CurrentDate     =   37405
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1950
         Width           =   2655
      End
      Begin VB.Label lbl就诊时间 
         AutoSize        =   -1  'True
         Caption         =   "就诊时间(&T)"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Left            =   2700
         TabIndex        =   14
         Top             =   2400
         Width           =   180
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号(&Y)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   3
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   255
         TabIndex        =   9
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   435
         TabIndex        =   7
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   615
         TabIndex        =   5
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl说明 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   615
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
   End
End
Attribute VB_Name = "frm医保帐户查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 文本Enum
    Text卡号 = 0
    Text医保号 = 1
    Text姓名 = 2
    Text身份证 = 3
    Text工作单位 = 4
End Enum

Private mstrFind As String
Private mblnOK As Boolean


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    mstrFind = ""
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If Trim(txtEdit(lngIndex).Text) <> "" Then
            Select Case lngIndex
                Case Text卡号
                    mstrFind = mstrFind & " And A.卡号 = '" & Trim(UCase(txtEdit(lngIndex).Text)) & "'"
                Case Text医保号
                    mstrFind = mstrFind & " And A.医保号 = '" & Trim(UCase(txtEdit(lngIndex).Text)) & "'"
                Case Text身份证
                    mstrFind = mstrFind & " And P.身份证号 = '" & Trim(txtEdit(lngIndex).Text) & "'"
                Case Text姓名
                    mstrFind = mstrFind & " And P.姓名 Like '" & Trim(txtEdit(lngIndex).Text) & "%'"
                Case Text工作单位
                    mstrFind = mstrFind & " And P.工作单位 like '" & Trim(txtEdit(lngIndex).Text) & "%'"
            End Select
        End If
    Next
    mstrFind = mstrFind & " And A.就诊时间>=to_date('" & Format(dtpBegin.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') And A.就诊时间<to_date('" & _
                                                        Format(dtpEnd.Value + 1, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetFind(strFind As String) As Boolean
    dtpEnd.Value = CDate(Format(zlDataBase.Currentdate, "yyyy-MM-dd"))
    dtpBegin = DateAdd("m", -1, dtpEnd.Value)
    dtpBegin.MaxDate = dtpEnd.Value
    
    mblnOK = False
    frm医保帐户查找.Show vbModal, frm医保帐户
    If mblnOK = True Then
        strFind = mstrFind
    End If
    GetFind = mblnOK
End Function

Private Sub cmd单位_Click()
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = frmPubSel.ShowSelect(Me, _
            " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
            2, "工作单位", , txtEdit(Text工作单位).Text)
    If Not rsTemp Is Nothing Then
        txtEdit(Text工作单位).Text = rsTemp("名称")
        zlControl.TxtSelAll txtEdit(Text工作单位)
    Else
        txtEdit(Text工作单位).SetFocus
    End If
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text姓名, Text工作单位
            zlCommFun.OpenIme True
        Case Text卡号, Text医保号, Text身份证
            zlCommFun.OpenIme False
    End Select
End Sub
