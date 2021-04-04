VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���㵥_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frm���㵥_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�ں� 
      Height          =   300
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   1
      Top             =   330
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3420
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   1350
      Width           =   6885
   End
   Begin MSComCtl2.DTPicker dtp�걨���� 
      Height          =   300
      Left            =   3975
      TabIndex        =   7
      Top             =   720
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   72417283
      CurrentDate     =   39349
   End
   Begin VB.TextBox txt����Ա 
      Height          =   300
      Left            =   1110
      TabIndex        =   4
      Top             =   720
      Width           =   1755
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   330
      Width           =   1785
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�걨����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3210
      TabIndex        =   6
      Top             =   780
      Width           =   720
   End
   Begin VB.Label lbl������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3210
      TabIndex        =   5
      Top             =   390
      Width           =   720
   End
   Begin VB.Label lbl����Ա 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����Ա"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   525
      TabIndex        =   3
      Top             =   780
      Width           =   540
   End
   Begin VB.Label lbl�ں� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ں�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   705
      TabIndex        =   0
      Top             =   390
      Width           =   360
   End
End
Attribute VB_Name = "frm���㵥_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFilter As String

Public Function ShowCondition() As String
    mstrFilter = ""
    Me.Show 1
    ShowCondition = mstrFilter
End Function

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim str�ں� As String, strReturn As String
    On Error GoTo errHand
    
    str�ں� = Trim(txt�ں�.Text)
    If str�ں� <> "" Then
        str�ں� = Mid(str�ں�, 1, 4) & "-" & Mid(str�ں�, 5, 2) & "-01"
        If Not IsDate(str�ں�) Then
            MsgBox "�ںŵĸ�ʽ��YYYYMM", vbInformation, gstrSysName
            txt�ں�.SetFocus
            Exit Sub
        End If
    End If
    
    If Me.cbo�������.ListIndex <> 0 Then strReturn = " And �������=" & Me.cbo�������.ListIndex
    strReturn = strReturn & " And ���� >= to_date('" & Format(Me.dtp�걨����.Value, "yyyy-MM-dd") & " 00:00:00','yyyy-MM-dd hh24:mi:ss')"
    If txt�ں�.Text <> "" Then strReturn = strReturn & " And �ں�='" & txt�ں�.Text & "'"
    If txt����Ա.Text <> "" Then strReturn = strReturn & " And ����Ա Like '" & txt����Ա.Text & "%'"
    
    mstrFilter = strReturn
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    With cbo�������
        .Clear
        .AddItem "ȫ��"
        .AddItem "��ҵְ������ҽ�Ʊ���"
        .AddItem "��ҵ����ҽ�Ʊ���"
        .AddItem "������ҵ��λҽ�Ʊ���"
        .AddItem "��������"
        .ListIndex = 0
    End With
    Me.dtp�걨����.Value = Format(DateAdd("m", -1, zldatabase.Currentdate()), "yyyy-MM-dd")
End Sub
