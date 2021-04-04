VERSION 5.00
Begin VB.Form frmDrugQueryParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   3270
   ClientLeft      =   3585
   ClientTop       =   4680
   ClientWidth     =   4770
   Icon            =   "frmDrugQueryParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Chk包含停用药品 
      Caption         =   "包含停用药品(&S)"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2910
      Width           =   2730
   End
   Begin VB.TextBox Txt效期报警 
      Height          =   300
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "3"
      Top             =   2160
      Width           =   300
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   3420
      TabIndex        =   12
      Top             =   2730
      Width           =   1100
   End
   Begin VB.CheckBox chk库存数 
      Caption         =   "只显示有库存数量的药品(&L)"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2580
      Width           =   2730
   End
   Begin VB.CommandButton Cmd保存 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   10
      Top             =   300
      Width           =   1100
   End
   Begin VB.CommandButton Cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3420
      TabIndex        =   11
      Top             =   780
      Width           =   1100
   End
   Begin VB.Frame Fra参数设置 
      Caption         =   "显示单位设置"
      Height          =   1905
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3000
      Begin VB.OptionButton Opt单位2 
         Caption         =   "门诊单位(&2)"
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1305
      End
      Begin VB.OptionButton Opt单位4 
         Caption         =   "住院单位(&4)"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1305
      End
      Begin VB.OptionButton Opt单位3 
         Caption         =   "药库单位(&3)"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
      End
      Begin VB.OptionButton Opt单位1 
         Caption         =   "售价单位(&1)"
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Label Lbl月 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1740
      TabIndex        =   7
      Top             =   2220
      Width           =   180
   End
   Begin VB.Label Lbl效期报警 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "效期报警(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   2220
      Width           =   990
   End
End
Attribute VB_Name = "frmDrugQueryParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private StrBillStyle As String
Private IntChoose As Integer
Private BlnBootUp As Boolean '启动成功否
Private mstrPrivs As String
Private mblnSetPara As Boolean      '是否具有参数设置权限
'注意:选择其中一个单位,则入库单据以此单位显示

Public Property Get In_权限() As String
    In_权限 = mstrPrivs
End Property

Public Property Let In_权限(ByVal vNewValue As String)
    mstrPrivs = vNewValue
End Property
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd保存_Click()
    If Val(Txt效期报警.Text) < 0 Then
        MsgBox "效期报警不能小于零！", vbInformation, gstrSysName
        Txt效期报警.SetFocus
        Exit Sub
    End If

    If Opt单位1.Value = True Then IntChoose = 1
    If Opt单位2.Value = True Then IntChoose = 2
    If Opt单位3.Value = True Then IntChoose = 3
    If Opt单位4.Value = True Then IntChoose = 4

    zlDataBase.SetPara "单位", IntChoose, glngSys, 1309
    zlDataBase.SetPara "是否显示无库存药品", chk库存数.Value, glngSys, 1309
    zlDataBase.SetPara "效期报警月数", Val(Txt效期报警.Text), glngSys, 1309
    zlDataBase.SetPara "是否显示停用药品", Chk包含停用药品.Value, glngSys, 1309

    frmDrugQuery.BlnDO = True
    Unload Me
End Sub

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Cmd取消_Click
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    
    Dim bln库存 As Boolean
    Dim intMonths As Integer
    
    mblnSetPara = zlStr.IsHavePrivs(mstrPrivs, "参数设置")

    IntChoose = Val(zlDataBase.GetPara("单位", glngSys, 1309, 3, Array(Fra参数设置), mblnSetPara))
    bln库存 = (zlDataBase.GetPara("是否显示无库存药品", glngSys, 1309, 0, Array(chk库存数), mblnSetPara) = 1)
    intMonths = Val(zlDataBase.GetPara("效期报警月数", glngSys, 1309, 3, Array(Txt效期报警), mblnSetPara))
    Chk包含停用药品.Value = Val(zlDataBase.GetPara("是否显示停用药品", glngSys, 1309, 0, Array(Chk包含停用药品), mblnSetPara))
    
    Select Case IntChoose
        Case 1
            Opt单位1.Value = True
        Case 2
            Opt单位2.Value = True
        Case 3
            Opt单位3.Value = True
        Case 4
            Opt单位4.Value = True
    End Select
    Me.chk库存数.Value = IIf(bln库存, 1, 0)
    Me.Txt效期报警 = intMonths
    
    If glngSys \ 100 = 8 Then
        Opt单位2.Visible = False
        Opt单位4.Visible = False
        Opt单位3.Caption = "采购单位(&3)"
        If Opt单位3.Value = 0 And Opt单位1.Value = 0 Then
            Opt单位1.Value = 1
        End If
    End If
    
    BlnBootUp = True
End Sub

Private Sub Txt效期报警_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub
