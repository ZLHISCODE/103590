VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form Frm��ҩ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "Frm��ҩ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsfMutiSelect 
      Height          =   2085
      Left            =   1920
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3678
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3600
      TabIndex        =   16
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   17
      Top             =   2970
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1545
      Left            =   870
      TabIndex        =   0
      Top             =   1050
      Width           =   5085
      Begin MSMask.MaskEdBox TxtЧ�� 
         Height          =   300
         Left            =   3390
         TabIndex        =   12
         Top             =   690
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Cmd���� 
         Caption         =   "��"
         Height          =   285
         Left            =   4590
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Width           =   285
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1050
         TabIndex        =   14
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   10
         Top             =   690
         Width           =   1485
      End
      Begin VB.TextBox TxtҩƷ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   8
         Tag             =   "3"
         Top             =   300
         Width           =   3825
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   6
         Tag             =   "1"
         Top             =   690
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   4
         Tag             =   "2"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LblЧ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ч��(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2700
         TabIndex        =   11
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   750
         Width           =   630
      End
      Begin VB.Label LblҩƷ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   750
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   1140
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   210
      Picture         =   "Frm��ҩ����.frx":000C
      Top             =   180
      Width           =   240
   End
   Begin VB.Label LblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ҩƷԭ�����������������ڷ���������ˣ��������ҩƷ��������Ϣ��"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   870
      TabIndex        =   18
      Top             =   240
      Width           =   5040
   End
End
Attribute VB_Name = "Frm��ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrData
Private strPar As String
Private strReturn As String
Private StrFindStyle As String
Private rsTmp As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'    If Trim(Txt����) = "" Then
'        MsgBox "���������ţ�", vbInformation, gstrSysName
'        Txt����.SetFocus
'        Exit Sub
'    End If
    If TxtЧ�� <> "____-__-__" Then
        If Not IsDate(TxtЧ��) Then
            MsgBox "������Ϸ���Ч�ڣ�", vbInformation, gstrSysName
            TxtЧ��.SetFocus
            Exit Sub
        End If
    End If
    If Trim(Txt����) <> "" Then Call Txt����_KeyDown(vbKeyReturn, 0)
    Do While True
        If Not MsfMutiSelect.Visible Then Exit Do
    Loop
    If Txt���� <> Txt����.Tag Then Exit Sub
    strReturn = Txt����.Text & "|" & IIf(TxtЧ�� = "____-__-__", "", TxtЧ��.Text) & "|" & Txt����.Tag
    
    Unload Me
End Sub

Private Sub Cmd����_Click()
    Dim Rec���� As New ADODB.Recordset
    
    On Error GoTo errHandle
    With Rec����
        If .State = 1 Then .Close
        gstrSQL = "Select ����,����,���� From ҩƷ������ Where Order By ���� "
        
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set Rec���� = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
        Call SQLTest
        
        If .EOF Then
            MsgBox "���ʼ��ҩƷ�����̣��ֵ������", vbInformation, gstrSysName
            Me.Txt����.SetFocus
            Txt����.Tag = ""
            Exit Sub
        End If
        
        With MsfMutiSelect
            .Clear
            Set .DataSource = Rec����
            .ColWidth(0) = 800
            .ColWidth(1) = 1500
            .ColWidth(2) = 800
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    strReturn = ""
    arrData = Split(strPar, "|")
    Txt���� = arrData(Val(Txt����.Tag))
    Txt���� = arrData(Val(Txt����.Tag))
    Txt���� = arrData(Val(Txt����.Tag))
    TxtҩƷ = arrData(Val(TxtҩƷ.Tag))
    TxtҩƷ.Tag = arrData(4)
    StrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(���Ч��,0) Ч�� From ҩƷĿ¼ Where ҩƷID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(TxtҩƷ.Tag))

    With rsTmp
        Txt����.Tag = !Ч��
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowME(ByVal frmParent As Object, strShow As String) As String
    'strShow="����|����|����|ҩƷ|ҩƷID"
    'strReturn="����|Ч��|����"
    strPar = strShow
    Me.Show 1, frmParent
    ShowME = strReturn
End Function

Private Sub MsfMutiSelect_DblClick()
    With MsfMutiSelect
        Txt���� = .TextMatrix(.Row, 1)
        Txt����.Tag = Txt����
    End With
    
    MsfMutiSelect.Visible = False
End Sub

Private Sub MsfMutiSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then MsfMutiSelect_DblClick
End Sub

Private Sub MsfMutiSelect_LostFocus()
    MsfMutiSelect.Visible = False
End Sub

Private Sub Txt����_GotFocus()
    Call GetFocus(Txt����)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errHandle
    Dim StrInput As String
    Dim Rec���� As New ADODB.Recordset
    StrInput = UCase(Trim(Txt����))
    If StrInput = "" Then
        Txt����.Tag = ""
        Exit Sub
    End If

    gstrSQL = "Select ����,����,���� From ҩƷ������ Where " & _
             " (Upper(����) Like [1] Or Upper(����) Like [1] Or Upper(����) Like [1]) Order By ���� "
    Set Rec���� = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, StrFindStyle & StrInput & "%")
    
    With Rec����
        If .EOF Then
            If Txt����.Tag <> UCase(Txt����.Text) Then
                MsgBox "û���ҵ�ƥ���ҩƷ�����̣����������룡", vbInformation, gstrSysName
                Txt����.SelStart = 0
                Txt����.SelLength = LenB(StrConv(Txt����, vbFromUnicode))
                Txt����.Tag = ""
            End If
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            With Txt����
                .Text = Rec����!����
                .Tag = .Text
            End With
        Else
            With MsfMutiSelect
                .Clear
                Set .DataSource = Rec����
                .ColWidth(0) = 800
                .ColWidth(1) = 1500
                .ColWidth(2) = 800
                .Visible = True
                .ZOrder 0
                
                .Row = 1
                .ColSel = .Cols - 1
                .SetFocus
            End With
        End If
        CmdOK.SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt����_Change()
    Dim str���� As String
    If Trim(Txt����) = "" Then Exit Sub
    If Len(Trim(Txt����)) <> 8 Then Exit Sub
    If Val(Txt����.Tag) = 0 Then Exit Sub
    str���� = Mid(Txt����, 1, 4) & "-" & Mid(Txt����, 5, 2) & "-" & Mid(Txt����, 7, 2)
    
    If IsDate(str����) Then
        TxtЧ�� = Format(DateAdd("m", Val(Txt����.Tag), str����), "yyyy-MM-dd")
    End If
    TxtЧ��.SetFocus
End Sub

Private Sub Txt����_GotFocus()
    Call GetFocus(Txt����)
End Sub

Public Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TxtЧ��.SetFocus
End Sub

Private Sub TxtЧ��_GotFocus()
    With TxtЧ��
        .SelStart = 0
        .SelLength = Len(TxtЧ��)
    End With
End Sub

Private Sub TxtЧ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt����.SetFocus
End Sub
