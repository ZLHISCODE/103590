VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLaterVisitFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   Icon            =   "frmLaterVisitFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   4515
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   1575
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75628547
         CurrentDate     =   38777
      End
      Begin VB.CheckBox chk 
         Caption         =   "ֻ�Ե�ǰҪ��õ���Ա(&3)"
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   1260
         Width           =   2730
      End
      Begin VB.CheckBox chk 
         Caption         =   "ֻ��ʾ������ڵ���Ա(&2)"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   915
         Width           =   2670
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   300
         Index           =   0
         Left            =   4095
         TabIndex        =   4
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   210
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   1980
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75628547
         CurrentDate     =   38777
      End
      Begin VB.Label lblHint 
         AutoSize        =   -1  'True
         Caption         =   "(��ʾ����Del��������Ͻ���)"
         Height          =   180
         Left            =   1500
         TabIndex        =   5
         Top             =   615
         Width           =   2610
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "������ʱ��(&5)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   1
         Top             =   2055
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��쿪ʼʱ��(&4)"
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   0
         Top             =   1635
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����Ͻ���(&1)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4695
      TabIndex        =   6
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4695
      TabIndex        =   7
      Top             =   510
      Width           =   1100
   End
End
Attribute VB_Name = "frmLaterVisitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mlngLoop As Long
Private mblnOK As Boolean

Private Type Items
    ���� As String
End Type

Private Type CONDITION
    ����id As Long
    ������� As Boolean             'ֻ��ʾ������ڵ���Ա
    ��ʼʱ�� As String              '��ʷ�����Ա����쿪ʼʱ��
    ����ʱ�� As String              '��ʷ�����Ա��������ʱ��
    �����Ա As Boolean             'ֻ��ʾ��ǰҪ��õ���Ա,ǰ�����������Ϊ���
End Type

Private mConditon As CONDITION

Private usrSave As Items

Private mlng����id As Long

Public Function ShowPara(ByVal frmMain As Object, ByRef ����id As Long, ByRef ������� As Boolean, ByRef �����Ա As Boolean, ByRef ��ʼʱ�� As String, ByRef ����ʱ�� As String) As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    mblnOK = False
    
    mConditon.����id = ����id
    mConditon.������� = �������
    mConditon.�����Ա = �����Ա
    mConditon.��ʼʱ�� = ��ʼʱ��
    mConditon.����ʱ�� = ����ʱ��
    
    Set mfrmMain = frmMain
    '��ʼ��
    
    chk(0).Value = IIf(mConditon.�������, 1, 0)
    chk(1).Value = IIf(mConditon.�����Ա, 1, 0)
    
    If mConditon.������� = False Then
        dtp(0).Enabled = True
    Else
        dtp(0).Enabled = False
    End If
    
    If mConditon.������� = False Then
        dtp(1).Enabled = True
    Else
        dtp(1).Enabled = False
    End If
    
    dtp(0).Value = Format(mConditon.��ʼʱ��, dtp(0).CustomFormat)
    dtp(1).Value = Format(mConditon.����ʱ��, dtp(1).CustomFormat)
    
    strSQL = "Select * from �����Ͻ��� Where ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mConditon.����id)
    If rs.BOF = False Then
        
        txt.Text = zlCommFun.NVL(rs("����").Value)
        cmd(0).Tag = mConditon.����id
        usrSave.���� = txt.Text
        
    End If
    
    Me.Show 1, frmMain
    
    ����id = mConditon.����id
    ������� = mConditon.�������
    �����Ա = mConditon.�����Ա
    ��ʼʱ�� = mConditon.��ʼʱ��
    ����ʱ�� = mConditon.����ʱ��
    
    ShowPara = mblnOK
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    Select Case Index
    Case 0
    
        chk(1).Enabled = (chk(Index).Value = 1)
        
        dtp(0).Enabled = (chk(Index).Value = 0)
        dtp(1).Enabled = (chk(Index).Value = 0)
        
        lbl(3).Enabled = dtp(0).Enabled
        lbl(2).Enabled = dtp(1).Enabled
        
        If chk(Index).Value = 0 Then
            chk(1).Value = 0
        End If
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim rsData As New ADODB.Recordset
    
    strSQL = "SELECT -1 AS ID," & _
                        "0 AS �ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "'' AS ����," & _
                        "'���з���' AS ����, " & _
                        "Null+0 AS �Ƿ񼲲�,'' As ��Ͻ��� " & _
                "FROM dual "
                
    strSQL = strSQL & _
            " UNION ALL " & _
            "SELECT ��� AS ID," & _
                        "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "����, " & _
                        "Null+0 AS �Ƿ񼲲�,'' As ��Ͻ��� " & _
                "FROM �����Ͻ��� " & _
                "WHERE NVL(ĩ��,0)=0 " & _
                "START WITH �ϼ���� is NULL CONNECT BY PRIOR ��� = �ϼ���� "
    
    strSQL = strSQL & _
                "UNION ALL " & _
                "SELECT A.��� AS ID, " & _
                        "DECODE(�ϼ����,NULL,-1,�ϼ����) AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.�Ƿ񼲲�,A.��Ͻ��� " & _
                "FROM �����Ͻ��� A " & _
                "WHERE NVL(A.ĩ��,0)=1"
                    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt, "����,900,0,1;����,1800,0,0;��Ͻ���,2700,0,0", Me.Name & "\������ѡ��", "�����±���ѡ��һ����Ͻ��ۡ�", rsData, rs, 8790, 5100) Then
    
        txt.Text = zlCommFun.NVL(rs("����").Value)
        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
        usrSave.���� = txt.Text
                
    End If

    txt.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
        
    mConditon.����id = Val(cmd(0).Tag)
    mConditon.������� = (chk(0).Value = 1)
    mConditon.�����Ա = (chk(1).Value = 1)
    
    mConditon.��ʼʱ�� = Format(dtp(0).Value, dtp(0).CustomFormat)
    mConditon.����ʱ�� = Format(dtp(1).Value, dtp(1).CustomFormat)

    mblnOK = True

    Unload Me
End Sub

Private Sub txt_Change()
    
    txt.Tag = "Changed"
    cmd(0).Tag = ""
    
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmd(0).Tag = ""
        txt.Text = ""
        txt.Tag = ""
        usrSave.���� = ""
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        
        If txt.Tag = "Changed" Then
                            
            strText = UCase(txt.Text) & "%"
            strSQL = _
                    "SELECT A.��� AS ID, " & _
                            "A.����, " & _
                            "A.����, " & _
                            "A.�Ƿ񼲲�,A.��Ͻ��� " & _
                    "FROM �����Ͻ��� A " & _
                    "WHERE NVL(ĩ��,0)=1 "
                    
            strSQL = strSQL & " AND (A.���� Like [1] OR Upper(A.����) Like [2] OR Upper(A.����) Like [2])"
            
            If ParamInfo.��Ŀ����ƥ�䷽ʽ = 0 Then strTmp = "%" & strText
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
            
            If ShowTxtFilter(Me, txt, "����,900,0,1;����,1800,0,0;��Ͻ���,2700,0,0", Me.Name & "\�����۹���", "�������ѡ��һ����Ͻ���", rsData, rs) Then
                
                txt.Text = zlCommFun.NVL(rs("����"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
                usrSave.���� = txt.Text
            Else
                txt.Text = usrSave.����
                Exit Sub
            End If
        End If
        
        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
    If txt.Tag = "Changed" Then
        txt.Text = usrSave.����
        txt.Tag = ""
    End If
    
End Sub
