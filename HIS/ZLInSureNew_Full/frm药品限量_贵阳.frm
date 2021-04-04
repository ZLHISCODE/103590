VERSION 5.00
Begin VB.Form frm药品限量_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品限量"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frm药品限量_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd药品 
      Caption         =   "…"
      Height          =   315
      Left            =   4380
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txt备注 
      Height          =   2085
      Left            =   870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1320
      Width           =   7125
   End
   Begin VB.TextBox txt单位 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   893
      Width           =   1245
   End
   Begin VB.TextBox txt规格 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5310
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   53
      Width           =   2685
   End
   Begin VB.TextBox txt产地 
      Enabled         =   0   'False
      Height          =   315
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   7125
   End
   Begin VB.TextBox txt售价 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   893
      Width           =   1305
   End
   Begin VB.TextBox txt限量 
      Height          =   315
      Left            =   4590
      TabIndex        =   5
      Top             =   893
      Width           =   975
   End
   Begin VB.TextBox txt金额 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   893
      Width           =   1515
   End
   Begin VB.TextBox txt药品名称 
      Height          =   315
      Left            =   870
      TabIndex        =   0
      Top             =   53
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   315
      Left            =   5850
      TabIndex        =   8
      Top             =   3690
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   6900
      TabIndex        =   9
      Top             =   3690
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   30
      TabIndex        =   10
      Top             =   3450
      Width           =   8175
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "本功能涉及到“限量”或“准许超量”时，均指按月限量，如果程序设置成按其他方式限量时，系统会自动换算"
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   8
      Left            =   60
      TabIndex        =   20
      Top             =   3630
      Width           =   5430
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   180
      Index           =   7
      Left            =   480
      TabIndex        =   18
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售价金额"
      Enabled         =   0   'False
      Height          =   180
      Index           =   6
      Left            =   5790
      TabIndex        =   17
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售价"
      Enabled         =   0   'False
      Height          =   180
      Index           =   5
      Left            =   2370
      TabIndex        =   16
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售价单位"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   15
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "限量"
      Height          =   180
      Index           =   3
      Left            =   4260
      TabIndex        =   14
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "产地"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   510
      TabIndex        =   13
      Top             =   540
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "规格"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "药品名称"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frm药品限量_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    If txt药品名称.Tag = "" Then ShowMsgbox "请确定药品信息！": txt药品名称.SetFocus: Exit Sub
    If Val(txt限量.Text) <= 0 Then ShowMsgbox "药品限量必须大于0！": txt限量.SetFocus: Exit Sub
    gstrSQL = "Zl_用药限量目录_贵阳_Insert(" & mint险类 & ",'" & txt药品名称.Tag & "','" & txt限量.Text & "','" & txt备注.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call frm门诊慢性疾病用药限量_贵阳.mnuViewRefresh_Click
    ShowMsgbox "药品限量设置成功！"
    If Me.Tag = "增加" Then
        txt药品名称.Text = "": txt备注.Text = "": txt产地.Text = "": txt金额.Text = ""
        txt规格.Text = ""
        txt规格.Text = ""
        txt单位.Text = "": txt售价.Text = ""
        txt限量.Text = 1: txt药品名称.Tag = "": txt药品名称.SetFocus
    Else
        Unload Me
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd药品_Click()
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select  ID, Null As 上级id, 0 As 末级, 名称 As 分类,  编码, 名称, Null As 规格, Null As 产地," & _
              "      Null As 售价单位, -null As 售价 From 药品用途分类 Union All " & _
              "Select B.药品id As ID, A.用途分类id As 上级id, 1 As 末级, D.名称 As 分类, B.编码, B.名称, B.规格, B.产地, B.售价单位," & _
              "      C.现价 As 售价 From 药品信息 A, 药品目录 B, 收费价目 C, 药品用途分类 D " & _
              "Where A.用途分类id = D.ID And A.药名id = B.药名id And B.药品id = C.收费细目id And " & _
              "     (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 2, "药品目录", , txt药品名称.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txt药品名称.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
            txt规格.Text = Nvl(rsTemp!规格): txt产地.Text = Nvl(rsTemp!产地)
            txt单位.Text = Nvl(rsTemp!售价单位): txt售价.Text = Format(Nvl(rsTemp!售价), "0.00000")
            txt限量.Text = 1: txt药品名称.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt限量.Text = 0: txt药品名称.Tag = ""
        End If
    Else
        txt限量.Text = 0: txt药品名称.Tag = ""
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Sub ShowME(ByVal intinsure As Integer, ByVal strMode As String, ByVal str药品ID As String)
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    Me.Caption = Me.Caption & "-" & strMode
    Me.Tag = strMode
    mint险类 = intinsure
    If strMode = "修改" Then
        gstrSQL = "Select A.药品ID,A.编码, A.名称, A.规格, A.产地, A.售价单位, trim(to_char(B.数量,'900090.00')) As 数量, " & _
              "      trim(to_char(C.现价,'900090.00000'))  As 售价, trim(to_char(Nvl(B.数量, 0) * Nvl(C.现价, 0),'90009990.00')) As 售价金额,B.备注 " & _
              "From 药品目录 A, 用药限量目录_贵阳 B, 收费价目 C " & _
              "Where A.药品id = B.药品id And B.药品id = C.收费细目ID And B.险类=[1]" & _
              " And (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.药品ID='" & str药品ID & "'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
        txt药品名称.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
        txt规格.Text = Nvl(rsTemp!规格): txt产地.Text = Nvl(rsTemp!产地)
        txt单位.Text = Nvl(rsTemp!售价单位): txt售价.Text = Format(Nvl(rsTemp!售价), "0.00000")
        txt限量.Text = Nvl(rsTemp!数量): txt药品名称.Tag = Nvl(rsTemp!药品id)
        txt备注.Text = Nvl(rsTemp!备注): txt药品名称.Enabled = False: cmd药品.Visible = False
    End If
    Me.Show 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt限量_Change()
    txt金额.Text = Format(Val(txt售价.Text) * Val(txt限量.Text), "0.00000")
End Sub

Private Sub txt限量_GotFocus()
     zlControl.TxtSelAll txt限量
End Sub

Private Sub txt限量_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt限量, KeyAscii, m金额式)
End Sub

Private Sub txt药品名称_GotFocus()
    zlControl.TxtSelAll txt药品名称
End Sub

Private Sub txt药品名称_KeyPress(KeyAscii As Integer)
     On Error GoTo errHand
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Distinct B.药品id As ID,Null As 上级ID,D.名称 As 分类, B.编码, B.名称, B.规格, B.产地," & _
              "       B.售价单位, C.现价 As 售价 From 药品信息 A, 药品目录 B, 收费价目 C, 药品用途分类 D, 收费别名 E " & _
              "Where A.用途分类id = D.ID And A.药名id = B.药名id And B.药品id = C.收费细目id And " & _
              "     (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.药品id = E.收费细目id " & _
              "     And (Upper(B.编码) Like '%" & UCase(txt药品名称.Text) & "%' Or  Upper(B.名称) Like '%" & UCase(txt药品名称.Text) & "%' " & _
              "     Or Upper(E.简码) Like '%" & UCase(txt药品名称.Text) & "%')"
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "药品目录", , txt药品名称.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txt药品名称.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
            txt规格.Text = Nvl(rsTemp!规格): txt产地.Text = Nvl(rsTemp!产地)
            txt单位.Text = Nvl(rsTemp!售价单位): txt售价.Text = Format(Nvl(rsTemp!售价), "0.00000")
            txt限量.Text = 1: txt药品名称.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt限量.Text = 0: txt药品名称.Tag = ""
            txt药品名称.SetFocus
        End If
    Else
        txt限量.Text = 0: txt药品名称.Tag = ""
        txt药品名称.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
