VERSION 5.00
Begin VB.Form frm��Ӧ�̶�λ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ӧ�̶�λ"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   60
      Left            =   -30
      TabIndex        =   20
      Top             =   600
      Width           =   5580
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   15
      Top             =   3045
      Width           =   1100
   End
   Begin VB.Frame fraTemp3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   2700
      Visible         =   0   'False
      Width           =   4830
      Begin VB.OptionButton optSelect 
         Caption         =   "ģ������(&3)"
         Height          =   180
         Index           =   3
         Left            =   2865
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "��ȷ����(&2)"
         Height          =   180
         Index           =   2
         Left            =   975
         TabIndex        =   10
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���ҷ�ʽ��"
         Height          =   180
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.Frame fraTemp2 
      Height          =   30
      Left            =   -15
      TabIndex        =   17
      Top             =   2910
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.CommandButton cmdAdva 
      Caption         =   "�߼�(&A)"
      Height          =   350
      Left            =   3270
      TabIndex        =   13
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   14
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��λ(&F)"
      Height          =   350
      Left            =   2175
      TabIndex        =   12
      Top             =   3045
      Width           =   1100
   End
   Begin VB.Frame fraTemp1 
      Height          =   30
      Left            =   -30
      TabIndex        =   16
      Top             =   2115
      Width           =   5580
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   2
      Left            =   1335
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "����"
      Top             =   1695
      Width           =   3735
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   1
      Left            =   1335
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "����"
      Top             =   1260
      Width           =   1230
   End
   Begin VB.Frame fraTemp4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   2340
      Visible         =   0   'False
      Width           =   4830
      Begin VB.OptionButton optSelect 
         Caption         =   "����������������(&1)"
         Height          =   180
         Index           =   1
         Left            =   2865
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "������������(&0)"
         Height          =   180
         Index           =   0
         Left            =   975
         TabIndex        =   7
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������ã�"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   0
      Left            =   1335
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "����"
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��������������,����λ��������Ҫ���ҵĹ�Ӧ��."
      Height          =   180
      Left            =   750
      TabIndex        =   21
      Top             =   375
      Width           =   4140
   End
   Begin VB.Image img���� 
      Height          =   480
      Left            =   120
      Picture         =   "frm��Ӧ�̶�λ.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��(&D)"
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   900
      Width           =   990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��λ����(&N)"
      Height          =   180
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��(&J)"
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   2
      Top             =   1320
      Width           =   990
   End
End
Attribute VB_Name = "frm��Ӧ�̶�λ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSql As String
Private mstrOthers(0 To 2) As String '0-����,1-����,2-����

Public Function getSql(ByRef strOthers() As String) As String
    cmdAdva_Click
    Me.Show vbModal
    getSql = mstrSql
    strOthers = mstrOthers
End Function

Private Sub cmdAdva_Click()
    If Left(cmdAdva.Caption, 2) = "�߼�" Then
        Me.Height = 3950
        cmdHelp.Top = Me.fraTemp1.Top + 1000
        cmdAdva.Top = cmdHelp.Top
        cmdOk.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        fraTemp2.Top = cmdHelp.Top - 100
        
        cmdAdva.Caption = "����(&A)"
        fraTemp2.Visible = True
        fraTemp3.Visible = True
        fraTemp4.Visible = True
    Else
        fraTemp2.Visible = False
        fraTemp3.Visible = False
        fraTemp4.Visible = False
        Me.Height = 3050
        cmdHelp.Top = Me.fraTemp1.Top + 100
        cmdAdva.Top = cmdHelp.Top
        cmdOk.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        fraTemp2.Top = cmdHelp.Top - 100
        cmdAdva.Caption = "�߼�(&A)"
    End If
End Sub

Private Sub cmdCancel_Click()
    mstrSql = ""
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strLinkStr As String, strLeftStr As String, strRightStr As String, intTemp As Integer, strTemp As String
    If optSelect(0).Value Then
        strLinkStr = " And "
    Else
        strLinkStr = " Or "
    End If
    If optSelect(2).Value Then
'        strLeftStr = " = '"
'        strRightStr = "'"
        strLeftStr = " = "
        strRightStr = ""
    Else
'        strLeftStr = " Like '" & IIf(gstrMatchMethod = "0", "%", "")
'        strRightStr = "%'"
        strLeftStr = " Like "
        strRightStr = "%"
    End If
    mstrSql = "Select ID,�ϼ�ID,����,���� From ��Ӧ�� Where ("
    strTemp = ""
    For intTemp = 0 To 2
'        If Trim(txtFind(intTemp)) <> "" Then strTemp = strTemp & IIf(strTemp = "", "", strLinkStr) & "Upper(" & txtFind(intTemp).Tag & ") " & strLeftStr & UCase(txtFind(intTemp)) & strRightStr
        If Trim(txtFind(intTemp)) <> "" Then
            strTemp = strTemp & IIf(strTemp = "", "", strLinkStr) & txtFind(intTemp).Tag & strLeftStr & "Upper([" & (intTemp + 8) & "])"
            mstrOthers(intTemp) = IIf(gstrMatchMethod = "0", "%", "") & UCase(txtFind(intTemp)) & strRightStr
        End If
    Next
    If strTemp = "" Then
        MsgBox "��ָ������һ����λ��������Ҫ�˳�������ȡ������ť��", vbInformation, Me.Caption
        txtFind(0).SetFocus
        Exit Sub
    End If
    mstrSql = mstrSql & strTemp & ") and ĩ��=1"
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub optSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtFind_GotFocus(Index As Integer)
    Dim blnOpen As Boolean
    
   Select Case Index
    Case 2
            blnOpen = True
    Case Else
            blnOpen = False
    End Select
    SetTxtGotFocus txtFind(Index), blnOpen
End Sub

Private Sub txtFind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m����ʽ
    Else
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m�ı�ʽ
    End If
End Sub

Private Sub txtFind_LostFocus(Index As Integer)
        ImeLanguage False
End Sub

