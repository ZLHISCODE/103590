VERSION 5.00
Begin VB.Form frm药品限量查找_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frm药品限量查找_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   2
      Top             =   1320
      Width           =   1100
   End
   Begin VB.TextBox txt药品信息 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "药品信息"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "可以输入药品编码、药品名称、简码进行查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   705
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   4260
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frm药品限量查找_贵阳"
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
    Dim rsTemp As New ADODB.Recordset, strWhere As String
    Me.MousePointer = vbHourglass
    strWhere = "  And (Upper(A.编码) Like '%" & UCase(txt药品信息.Text) & "%' Or  Upper(A.名称) Like '%" & UCase(txt药品信息.Text) & "%' " & _
              "     Or Upper(D.简码) Like '%" & UCase(txt药品信息.Text) & "%')"
    gstrSQL = "Select Distinct A.药品ID,A.编码, A.名称, A.规格, A.产地, A.售价单位, trim(to_char(B.数量,'900090.00')) As 限量, " & _
              "      trim(to_char(C.现价,'900090.00000'))  As 售价, trim(to_char(Nvl(B.数量, 0) * Nvl(C.现价, 0),'90009990.00')) As 售价金额,B.备注 " & _
              "From 药品目录 A, 用药限量目录_贵阳 B, 收费价目 C,收费别名 D " & _
              "Where A.药品id = B.药品id And B.药品id = C.收费细目ID And B.险类=[1] And B.药品ID=D.收费细目ID " & _
              " And (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) " & strWhere & " Order By A.名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    Set frm门诊慢性疾病用药限量_贵阳.mshBill.DataSource = rsTemp
    Call CenterTableCaption(frm门诊慢性疾病用药限量_贵阳.mshBill)
    frm门诊慢性疾病用药限量_贵阳.mshBill.ColWidth(0) = 0
    Call frm门诊慢性疾病用药限量_贵阳.SetMenu
    Me.MousePointer = vbDefault
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub ShowME(ByVal intinsure As Integer)
    mint险类 = intinsure
    Me.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
