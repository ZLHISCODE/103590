VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ס"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraGroup 
      Height          =   1335
      Index           =   1
      Left            =   3150
      TabIndex        =   38
      Top             =   1920
      Width           =   3480
      Begin VB.ComboBox cbo���λ�ʿ 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox cbo����ҽʦ 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   945
         Width           =   1830
      End
      Begin VB.ComboBox cbo����ҽʦ 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   565
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(������)ҽʦ"
         Height          =   180
         Left            =   75
         TabIndex        =   24
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���λ�ʿ"
         Height          =   180
         Left            =   795
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   795
         TabIndex        =   23
         Top             =   625
         Width           =   720
      End
   End
   Begin VB.Frame fraGroup 
      Height          =   1335
      Index           =   0
      Left            =   105
      TabIndex        =   37
      Top             =   1920
      Width           =   2970
      Begin VB.ComboBox cbo����ҽʦ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   950
         Width           =   1890
      End
      Begin VB.ComboBox cboҽ��С�� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   1890
      End
      Begin VB.ComboBox cboסԺҽʦ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   565
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��С��"
         Height          =   180
         Left            =   210
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   1010
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺҽʦ"
         Height          =   180
         Left            =   210
         TabIndex        =   20
         Top             =   625
         Width           =   720
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6705
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5610
      Width           =   6705
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   105
         TabIndex        =   18
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4260
         TabIndex        =   16
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5445
         TabIndex        =   17
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "��������"
      Height          =   1830
      Left            =   105
      TabIndex        =   28
      Top             =   3720
      Width           =   6525
      Begin MSComctlLib.ListView lvw 
         Height          =   1425
         Left            =   150
         TabIndex        =   25
         Top             =   255
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2514
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��λ"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1935
      Left            =   105
      TabIndex        =   27
      Top             =   0
      Width           =   6525
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1150
         Width           =   1890
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   765
         Width           =   1890
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   765
         Width           =   1170
      End
      Begin VB.ComboBox cbo����ȼ� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   5505
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "�Ƿ����"
         Height          =   195
         Left            =   4650
         TabIndex        =   15
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3300
         TabIndex        =   4
         Top             =   1203
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4230
         TabIndex        =   35
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2850
         TabIndex        =   34
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   570
         TabIndex        =   33
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4230
         TabIndex        =   32
         Top             =   825
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ȼ�"
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   570
         TabIndex        =   30
         Top             =   825
         Width           =   360
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   570
         TabIndex        =   29
         Top             =   1210
         Width           =   360
      End
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   300
      Left            =   4740
      TabIndex        =   14
      Top             =   3375
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   19
      Format          =   "yyyy-MM-dd hh:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtIn 
      Height          =   300
      Left            =   1065
      TabIndex        =   40
      Top             =   3375
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   16
      Format          =   "yyyy-MM-dd hh:mm"
      Mask            =   "####-##-## ##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ժʱ��"
      Height          =   180
      Left            =   315
      TabIndex        =   39
      Top             =   3435
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��סʱ��"
      Height          =   180
      Left            =   3945
      TabIndex        =   36
      Top             =   3435
      Width           =   720
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long
Public mlng��ҳID As Long
Public mlngUnit As Long
Public mbyt��ס��ʽ As Byte '0-��Ժ��ס��1-ת����ס

Public mstr���� As String '��:ȱʡ��λ�Ĵ���,��ʾ��ͥ����,��:��ס�Ĵ���,���ܶ��Ŵ�,��,�ŷָ�
Public mlng��λ����ID As Long
Public mstrPrivs As String
Private mfrmParent As Object
Private mblnAppoint As Boolean      'T-ԤԼ���Ĳ���;False-��ԤԼ���Ĳ���
Private mstrAppointBed As String    'ԤԼ���İ��Ŵ�λ
Private mstrIDs As String
Private mstrText As String
Private mrsPatiInfo As ADODB.Recordset
Private mint�������� As Integer

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_Click()
    cbo.SetListWidth cbo����.hWnd, cbo����.width * 1.8
    If mblnAppoint Then
        cbo����.Tag = Trim(Split(cbo����.Text, " ")(0))
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����ȼ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ȼ�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ȼ�.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ȼ�.ListIndex = lngIdx
    ElseIf cbo����ȼ�.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_Click()
    Dim rsTmp As ADODB.Recordset
    If mstrText = cbo����.Text Then Exit Sub
    If cbo����.Text = "" Then Exit Sub
    mstrText = cbo����.Text
    '��ʾ�ÿ��ҵĴ�λ
    Call ShowBeds
    If Not Visible Then chk����_Click
    
    On Error GoTo errHandle
    
     'ҽ��С��
    gstrSQL = "Select ID,����,˵��,����ʱ��,����ʱ�� From �ٴ�ҽ��С�� Where ����id=[1] " & _
            " And (����ʱ�� Is NULL Or Trunc(����ʱ��) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val("" & cbo����.ItemData(cbo����.ListIndex)))
    
    cboҽ��С��.Clear
    Do Until rsTmp.EOF
        cboҽ��С��.AddItem rsTmp!ID & "-" & rsTmp!����
        cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    cboҽ��С��.AddItem "": cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = 0: cboҽ��С��.ListIndex = cboҽ��С��.ListCount - 1
    If cboҽ��С��.ListCount = 1 Then cboҽ��С��.Enabled = False
    
    'ȱʡ��λ�ÿ��ҵ�ҽ������ʿ
    
    Call SeekDoctor(cbo���λ�ʿ, NVL(mrsPatiInfo!���λ�ʿ))
    Call SeekDoctor(cbo����ҽʦ, NVL(mrsPatiInfo!����ҽʦ))
    Call SeekDoctor(cboסԺҽʦ, NVL(mrsPatiInfo!סԺҽʦ))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboҽ��С��_Click()
    Dim strSQL As String, strSQLҽ��С�� As String
    Dim rsTmp As ADODB.Recordset
    Dim lngҽʦ As Long
    
    If cboҽ��С��.ListCount = 1 Then Exit Sub
    On Error GoTo errHandle
    '���Ϊ����ָ����ҽ��С�飬��"סԺҽʦ������ҽʦ"���Ӷ�Ӧҽ��С���е�ҽ����ѡ��
    strSQLҽ��С�� = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C, ҽ��С����Ա D" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And a.id = d.��Աid And B.��Ա���� = 'ҽ��' And d.С��id = [1] And" & vbNewLine & _
                        "   (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "   Instr(',' || [2] || ',', ',' || C.����id || ',') > 0 And Instr(',' || [3] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "   And (A.վ��=[4] Or A.վ�� is Null)" & vbNewLine & _
                        " Order By A.����"
    strSQL = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And B.��Ա���� = 'ҽ��' And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "      Instr(',' || [1] || ',', ',' || C.����id || ',') > 0 And Instr(',' || [2] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "      And (A.վ��=[3] Or A.վ�� is Null)" & _
                        " Order By A.����"
    
    If cboҽ��С��.ListIndex <> -1 And cboҽ��С��.ListIndex <> cboҽ��С��.ListCount - 1 Then
        If Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)) > 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
            If Not rsTmp.RecordCount > 0 Then
                '���С��δ����ҽ�����򱣳���ǰ�Ŀ���ѡ��Χ
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
            End If
            If cboסԺҽʦ.ListIndex <> -1 Then
                lngҽʦ = cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)
            Else
                lngҽʦ = 0
            End If
            cboסԺҽʦ.Clear
            Do Until rsTmp.EOF
                cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
            '105133:��סԺҽʦ����ѡҽ��С��ʱ���ı�סԺҽʦ
            If lngҽʦ <> 0 Then Call cbo.SetIndex(cboסԺҽʦ.hWnd, cbo.FindIndex(cboסԺҽʦ, lngҽʦ))
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
            
            If Not rsTmp.RecordCount > 0 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
            End If
            If cbo����ҽʦ.ListIndex <> -1 Then
                lngҽʦ = cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)
            Else
                lngҽʦ = 0
            End If
            cbo����ҽʦ.Clear
            Do Until rsTmp.EOF
                cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
             '105133:������ҽʦ����ѡҽ��С��ʱ���ı�����ҽʦ
            If lngҽʦ <> 0 Then Call cbo.SetIndex(cbo����ҽʦ.hWnd, cbo.FindIndex(cbo����ҽʦ, lngҽʦ))
        End If
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
        cboסԺҽʦ.Clear
        Do Until rsTmp.EOF
            cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
        cbo����ҽʦ.Clear
        Do Until rsTmp.EOF
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
    End If
    
    cboסԺҽʦ.AddItem "����..."
    cbo����ҽʦ.AddItem "����..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboҽ��С��_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cboҽ��С��.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cboҽ��С��.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cboҽ��С��.ListIndex = lngIdx
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����ҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ҽʦ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ҽʦ.ListIndex = lngIdx
    ElseIf cbo����ҽʦ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����ҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ҽʦ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ҽʦ.ListIndex = lngIdx
    ElseIf cbo����ҽʦ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboסԺҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If cboסԺҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cboסԺҽʦ.ListCount - 1
                If cboסԺҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cboסԺҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cboסԺҽʦ.ListCount - 1
            cboסԺҽʦ.ListIndex = cboסԺҽʦ.NewIndex
            cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!�ϼ�ID
        Else
            cboסԺҽʦ.ListIndex = -1
        End If
    Else
        If cboҽ��С��.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,����,˵�� From �ٴ�ҽ��С�� A, ҽ��С����Ա B " & _
                "Where a.id=b.С��id And b.��Աid=[1] And a.����id=[2] And (����ʱ�� Is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cboסԺҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)), Val(cbo����.ItemData(cbo����.ListIndex)))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, NVL(rsTmp!����), True))
                Exit Sub
            End If
        End If
        If cbo����ҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)), Val(cbo����.ItemData(cbo����.ListIndex)))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, NVL(rsTmp!����), True))
            Else
                Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo����ҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    On Error GoTo errHandle
    
    If cbo����ҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo����ҽʦ.ListCount - 1
                If cbo����ҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo����ҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ҽʦ.ListCount - 1
            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        Else
            cbo����ҽʦ.ListIndex = -1
        End If
    Else
        '����ҽʦѡ��ʱҽ��С����סԺҽʦΪ��
        If cboҽ��С��.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,����,˵�� From �ٴ�ҽ��С�� A, ҽ��С����Ա B " & _
                "Where a.id=b.С��id And b.��Աid=[1] And a.����id=[2] And (����ʱ�� Is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cboסԺҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)), Val(cbo����.ItemData(cbo����.ListIndex)))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, NVL(rsTmp!����), True))
                Exit Sub
            End If
        End If
        If cbo����ҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)), Val(cbo����.ItemData(cbo����.ListIndex)))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, NVL(rsTmp!����), True))
            Else
                Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo����ҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo����ҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo����ҽʦ.ListCount - 1
                If cbo����ҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo����ҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ҽʦ.ListCount - 1
            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        Else
            cbo����ҽʦ.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo���λ�ʿ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo���λ�ʿ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("��ʿ", "", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo���λ�ʿ.ListCount - 1
                If cbo���λ�ʿ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo���λ�ʿ.ListIndex = i: Exit Sub
                End If
            Next
            cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo���λ�ʿ.ListCount - 1
            cbo���λ�ʿ.ListIndex = cbo���λ�ʿ.NewIndex
            cbo���λ�ʿ.ItemData(cbo���λ�ʿ.NewIndex) = rsTmp!�ϼ�ID
        Else
            cbo���λ�ʿ.ListIndex = -1
        End If
    End If
End Sub


Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If cbo����.Locked Then Exit Sub
        If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����ҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ҽʦ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ҽʦ.ListIndex = lngIdx
    ElseIf cbo����ҽʦ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo���λ�ʿ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo���λ�ʿ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo���λ�ʿ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo���λ�ʿ.ListIndex = lngIdx
    ElseIf cbo���λ�ʿ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboסԺҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cboסԺҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cboסԺҽʦ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cboסԺҽʦ.ListIndex = lngIdx
    ElseIf cboסԺҽʦ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chk����_Click()
    If chk����.Value = 1 Then
        lbl��λ.Caption = "��Ҫ��λ"
        Call LoadMainBed
        lvw.Visible = True
        Me.Height = Me.Height + fraLvw.Height + 80
        If Visible Then lvw.SetFocus
    Else
        lbl��λ.Caption = "��λ"
        Call ShowBeds
        lvw.Visible = False
        Me.Height = Me.Height - fraLvw.Height - 80
        If Visible Then cmdOK.SetFocus
    End If
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strSQLҽ��С�� As String
    Dim strIDs As String, strID As String, strCode As String
    Dim strTmp As String
    Dim blnNurseGrade As Boolean    '����ȼ�Ĭ��Ϊ�� ?
    Dim strInfo As String, blnHeav As Boolean
    
    On Error GoTo errH
    gblnOK = False
    
    '50194:������,2012-09-21,ת�ƽ�����סʱ��飺
    '����ѽ��в�����ҳ��סԺҽ��ǩ����������ҽ��ǩ����������ҽ��ǩ�������ֹ������ס��������ʾӦ�����ɸ�ҽ��ȡ��ǩ���ٽ��С�
    If mbyt��ס��ʽ <> 0 Then 'ת����ס
        '��ȡ��ҳ�Ѿ�ǩ����߼���
        strInfo = "�ò��˵���ҳ�Ѿ�������ҽ��������ǩ����"
        blnHeav = False
        strSQL = "Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣֵ is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        rsTmp.Filter = "��Ϣ��='סԺҽʦǩ��'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!��Ϣֵ) Then
                strInfo = strInfo & vbCrLf & "סԺҽʦǩ����" & NVL(rsTmp!��Ϣֵ) & "��"
                blnHeav = True
            End If
        End If
        rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!��Ϣֵ) Then
                strInfo = strInfo & vbCrLf & "����ҽʦǩ����" & NVL(rsTmp!��Ϣֵ) & "��"
                blnHeav = True
            End If
        End If
        rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!��Ϣֵ) Then
                strInfo = strInfo & vbCrLf & "����ҽʦǩ����" & NVL(rsTmp!��Ϣֵ) & "��"
                blnHeav = True
            End If
        End If
        strInfo = strInfo & vbCrLf & "����������ҽ��ȡ����ҳǩ���ڽ���ת����ס������"
        strSQL = ""
        
        If blnHeav = True Then
            MsgBox strInfo, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID, mbyt��ס��ʽ)
    '����28432 by lesfeng 2010-03-10
    mint�������� = Val(zlDatabase.GetPara("�����������", glngSys, glngModul, 0))
    '��ʼ������
    With mrsPatiInfo
        cbo����.Enabled = mlng��λ����ID = 0
        If mint�������� = 0 And cbo����.Enabled = True Then
            cbo����.Enabled = False
        End If
        If mbyt��ס��ʽ = 0 Then
            mstrAppointBed = ""
            mblnAppoint = IsAppointPati(Val(!�Һ�ID & ""), mstrAppointBed) 'T-ԤԼ���Ĳ���
        End If
        '��ѡ�����Ŀ����벡�˿��Ҳ�ͬʱ,��������.
        If mlng��λ����ID <> 0 Then
            If mbyt��ס��ʽ = 0 Then      '���벡��
                If mlng��λ����ID <> !��Ժ����id Then
                    '����28432 by lesfeng 2010-03-10
                    If mint�������� = 1 Then
                        cbo����.Enabled = True
                    Else
                        MsgBox "���˵ǼǵĿ��ҡ�" & !��ǰ���� & "����ѡ��Ĵ�λ�������ҡ�" & GetDeptName(mlng��λ����ID) & "����ͬ,������ס�ô�λ,��ѡ��������λ!", vbInformation, gstrSysName
                        Unload Me: Exit Sub
                    End If
                End If
            Else                           'ת�Ʋ���
                If mlng��λ����ID <> !��ס����id Then
                    MsgBox "����ת��Ŀ��ҡ�" & !��ǰ���� & "����ѡ��Ĵ�λ�������ҡ�" & GetDeptName(mlng��λ����ID) & "����ͬ,������ס�ô�λ,��ѡ��������λ!", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
        End If
        
        If mbyt��ס��ʽ = 0 And gbyt���ʱ�� = 0 Then
            txtDate.Text = Format(!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        End If
        
        If mbyt��ס��ʽ = 0 And Val(zlDatabase.GetPara("�����޸���Ժʱ��", glngSys, 1132)) = 1 Then
            lblIn.Visible = True
            txtIn.Visible = True
            txtIn.Text = Format(!��Ժʱ��, "yyyy-MM-dd HH:mm")
        Else
            lblIn.Visible = False
            txtIn.Visible = False
        End If
        
        '������Ϣ
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
                
        
        'ȷ�������ķ������
        strSQL = "Select ������� From ��������˵�� Where ��������='����' And ����ID=[1]" '
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
            
        '�д�λ���ٴ�����
        If rsTmp!������� = 1 Then
            strTmp = "1,3"
        ElseIf rsTmp!������� = 2 Then
            strTmp = "2,3"
        ElseIf rsTmp!������� = 3 Then
            If Val("" & !��������) = 1 Then
                strTmp = "1,3"
            Else
                strTmp = "2,3"
            End If
        End If
        Set rsTmp = GetDeptOrUnit(0, mlngUnit, strTmp)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
                If mlng��λ����ID = 0 Then '�ǲ����ϵ�ָ���Ĵ���
                    If mbyt��ס��ʽ = 0 Then '���벡��ȱʡȡ�Ǽǿ���
                        If rsTmp!ID = !��Ժ����id Then cbo����.ListIndex = cbo����.NewIndex     '����click�¼����ش�λ
                    Else
                        'ת�Ʋ���ȱʡȡת�����
                        If rsTmp!ID = !��ס����id Then cbo����.ListIndex = cbo����.NewIndex
                    End If
                Else
                    '��ס���������Ĳ��˿������ɴ�λ����
                    If rsTmp!ID = mlng��λ����ID Then cbo����.ListIndex = cbo����.NewIndex
                End If
                strIDs = strIDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
        Else
            'û�ж�Ӧ�Ĵ�λ����
            MsgBox "�ڵ�ǰ����û�����ö�Ӧ����,���˲�����ס��" & vbCrLf, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        mstrIDs = strIDs
        'ԤԼ��λ��ռ��
        If mbyt��ס��ʽ = 0 And mblnAppoint And mstrAppointBed <> cbo����.Tag Then
            If MsgBox("�ڵ�ǰ���ô�λ��û���ҵ�����ԤԼ�Ĵ�λ��" & mstrAppointBed & "�����Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Unload Me: Exit Sub
            End If
        End If
        '����
        strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ���� Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 And cbo����.ListIndex = -1 Then cbo����.ListIndex = cbo����.NewIndex
                If rsTmp!���� = "" & !��ǰ���� Then cbo����.ListIndex = cbo����.NewIndex
                rsTmp.MoveNext
            Next
        End If
    
        '����ȼ�
        If mbyt��ס��ʽ = 1 Then cbo����ȼ�.Enabled = InStr(mstrPrivs, ";" & "��������ȼ�" & ";") > 0
        Set rsTmp = GetNurseGrade
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����ȼ�.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����ȼ�.ItemData(cbo����ȼ�.NewIndex) = rsTmp!ID
                If rsTmp!ID = !����ȼ�ID Then cbo����ȼ�.ListIndex = cbo����ȼ�.NewIndex
                rsTmp.MoveNext
            Next
        End If
        
        blnNurseGrade = zlDatabase.GetPara("����ȼ�Ĭ��Ϊ��", glngSys, 1132, 0)
        If blnNurseGrade And mbyt��ס��ʽ = 1 Then cbo����ȼ�.ListIndex = -1
        
        cboҽ��С��.Clear
        If Not cbo����.ListIndex = -1 Then
            'ҽ��С��
            strSQL = "Select ID,����,˵��,����ʱ��,����ʱ�� From �ٴ�ҽ��С�� Where ����id=[1] " & _
                    " And (����ʱ�� Is NULL Or Trunc(����ʱ��) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo����.ItemData(cbo����.ListIndex)))
            
            cboҽ��С��.Clear
            Do Until rsTmp.EOF
                cboҽ��С��.AddItem rsTmp!ID & "-" & rsTmp!����
                cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
        End If
        cboҽ��С��.AddItem "": cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = 0: cboҽ��С��.ListIndex = cboҽ��С��.ListCount - 1
        If cboҽ��С��.ListCount = 1 Then cboҽ��С��.Enabled = False
        'by lesfeng 2010-01-12 �����Ż�
        'סԺҽʦ,����ҽʦ,����ҽʦ
        '���Ϊ����ָ����ҽ��С�飬��"סԺҽʦ������ҽʦ"���Ӷ�Ӧҽ��С���е�ҽ����ѡ��
        strSQLҽ��С�� = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C, ҽ��С����Ա D" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And a.id = d.��Աid And B.��Ա���� = 'ҽ��' And d.С��id = [1] And" & vbNewLine & _
                        "   (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "   Instr(',' || [2] || ',', ',' || C.����id || ',') > 0 And Instr(',' || [3] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "   And (A.վ��=[4] Or A.վ�� is Null)" & vbNewLine & _
                        " Order By A.����"
        strSQL = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                            " From ��Ա�� A, ��Ա����˵�� B, ������Ա C" & vbNewLine & _
                            " Where A.ID = B.��Աid And A.ID = C.��Աid And B.��Ա���� = 'ҽ��' And" & vbNewLine & _
                            "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                            "      Instr(',' || [1] || ',', ',' || C.����id || ',') > 0 And Instr(',' || [2] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                            "      And (A.վ��=[3] Or A.վ�� is Null)" & _
                            " Order By A.����"
        If cboҽ��С��.ListCount = 1 Then
            If cboҽ��С��.ListIndex <> -1 And cboҽ��С��.ListIndex <> cboҽ��С��.ListCount - 1 Then
                If Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)) > 0 Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
                    If Not rsTmp.RecordCount > 0 Then
                        '���С��δ����ҽ�����򱣳���ǰ�Ŀ���ѡ��Χ
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
                    End If
                    cboסԺҽʦ.Clear
                    Do Until rsTmp.EOF
                        cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                        cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
                        rsTmp.MoveNext
                    Loop
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
    
                    If Not rsTmp.RecordCount > 0 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
                    End If
                    cbo����ҽʦ.Clear
                    Do Until rsTmp.EOF
                        cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                        cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
                        rsTmp.MoveNext
                    Loop
                End If
            Else
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
                cboסԺҽʦ.Clear
                Do Until rsTmp.EOF
                    cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                    cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
                    rsTmp.MoveNext
                Loop
    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
                cbo����ҽʦ.Clear
                Do Until rsTmp.EOF
                    cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                    cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
                    rsTmp.MoveNext
                Loop
            End If
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "����ҽʦ,������ҽʦ", gstrNodeNo)
        Do Until rsTmp.EOF
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        'ת����ס
        If mbyt��ס��ʽ = 1 Then
            If Not cbo.Locate(cboסԺҽʦ, "" & !סԺҽʦ) Then
                Call GetPersonnelIDCode("" & !סԺҽʦ, strID, strCode)
                cboסԺҽʦ.AddItem strCode & "-" & !סԺҽʦ
                cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = Val(strID)
                cboסԺҽʦ.ListIndex = cboסԺҽʦ.NewIndex
                strID = "": strCode = ""
            End If
            
            strSQL = " Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where (��Ϣ��='����ҽʦ' Or ��Ϣ��='����ҽʦ') And ����ID=[1] And ��ҳID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            
            rsTmp.Filter = "��Ϣ��='����ҽʦ'"
            If Not rsTmp.EOF Then
                If Not cbo.Locate(cbo����ҽʦ, "" & rsTmp!��Ϣֵ) Then
                    Call GetPersonnelIDCode("" & rsTmp!��Ϣֵ, strID, strCode)
                    cbo����ҽʦ.AddItem strCode & "-" & rsTmp!��Ϣֵ
                    cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = Val(strID)
                    cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
                    strID = "": strCode = ""
                End If
            End If
            
            rsTmp.Filter = "��Ϣ��='����ҽʦ'"
            If Not rsTmp.EOF Then
                If Not cbo.Locate(cbo����ҽʦ, "" & rsTmp!��Ϣֵ) Then
                    Call GetPersonnelIDCode("" & rsTmp!��Ϣֵ, strID, strCode)
                    cbo����ҽʦ.AddItem strCode & "-" & rsTmp!��Ϣֵ
                    cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = Val(strID)
                    cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
                    strID = "": strCode = ""
                End If
            End If
        '��ס
        Else
            Call SeekDoctor(cboסԺҽʦ, "" & !סԺҽʦ)
            Call SeekDoctor(cbo����ҽʦ, "" & !סԺҽʦ)
            '����ҽʦ,һ���޷�ȷ��ȱʡ
        End If
    
        '����ҽʦ(����)
        strSQL = "SELECT DISTINCT a.Id, a.���, a.����, a.���� " & vbNewLine & _
                " FROM ��Ա�� a, ��Ա����˵�� b, ������Ա c, ��������˵�� d " & vbNewLine & _
                " WHERE a.Id = b.��Աid AND a.Id = c.��Աid AND c.����id = d.����id AND b.��Ա���� = 'ҽ��' AND d.������� IN (1, 2, 3) AND " & vbNewLine & _
                "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') OR a.����ʱ�� IS NULL) " & vbNewLine & _
                " ORDER BY ���� "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo����ҽʦ, "" & !����ҽʦ)
        End If
    
        'סԺ��ʿ
        Set rsTmp = GetDoctorOrNurse(1, strIDs & "," & mlngUnit & ",")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo���λ�ʿ.ItemData(i - 1) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo���λ�ʿ, "" & !���λ�ʿ)
        End If
        
        cbo����ҽʦ.AddItem "����..."
        cbo���λ�ʿ.AddItem "����..."
        If InStr(mstrPrivs, "��������ҽʦ") = 0 Then
            cbo����ҽʦ.Enabled = False
        End If
    End With
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngUnit = 0
    mstrText = ""
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadMainBed
End Sub


Private Sub LoadMainBed()
    Dim i As Integer, strBed As String
    
    If cbo����.ListIndex <> -1 Then strBed = cbo����.Text
    cbo����.Clear
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked Then
            cbo����.AddItem lvw.ListItems(i).Text
            If lvw.ListItems(i).Text = strBed Then cbo����.ListIndex = cbo����.NewIndex
            If cbo����.ListIndex = -1 And mstr���� <> "" Then
                If lvw.ListItems(i).Text = mstr���� Then cbo����.ListIndex = cbo����.NewIndex
            End If
        End If
    Next
    If cbo����.ListIndex = -1 And cbo����.ListCount = 1 Then cbo����.ListIndex = 0
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If IsDate(txtDate.Text) And KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ShowBeds()
'���ܣ���ʾ��ǰ������ǰ���ҿ��õĲ���
    Dim i As Integer, objItem As ListItem
    Dim lng����ID As Long
    Dim rsBeds As ADODB.Recordset
    Dim strBed  As String
    
    lvw.ListItems.Clear
    cbo����.Clear: cbo����.Tag = ""
    If InStr(1, mstrPrivs, "��ͥ����") > 0 Then
        cbo����.AddItem "��ͥ����"
        If mstr���� = "��ͥ����" Then cbo����.ListIndex = 0
    End If
    If cbo����.ListIndex <> -1 Then lng����ID = cbo����.ItemData(cbo����.ListIndex)
    Set rsBeds = GetFreeBeds(mlngUnit, lng����ID, mrsPatiInfo!�Ա�, mlng����ID)
    If mstrAppointBed <> "" Then
        strBed = mstrAppointBed
    Else
        strBed = mstr����
    End If
    With rsBeds
        For i = 1 To rsBeds.RecordCount
            Set objItem = lvw.ListItems.Add(, "_" & !����, !���� & IIf(IsNull(!�����), "", " ����:" & !����� & "|") & _
                            IIf(IsNull(!�����) Or ((Not IsNull(!�����)) And Trim(NVL(!�Ա�) = "")), "", "(" & NVL(!�Ա�) & ")"))
            objItem.Tag = !�ȼ�ID
            cbo����.AddItem objItem.Text
            If !���� = strBed Then
                objItem.Checked = True: objItem.Selected = True: objItem.EnsureVisible
                cbo����.ListIndex = cbo����.NewIndex
                cbo����.Tag = !����
            End If
            .MoveNext
        Next
    End With
    
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub SeekDoctor(cbo As ComboBox, Optional strPre As String)
    Dim strIDs As String, i As Integer
    
    If cbo����.ListIndex = -1 Then Exit Sub
    
    If strPre <> "" Then
        For i = 0 To cbo.ListCount - 1
            If zlCommFun.GetNeedName(cbo.List(i)) = strPre Then cbo.ListIndex = i: Exit Sub
        Next
    End If
    
    strIDs = GetDeptDoctors(cbo����.ItemData(cbo����.ListIndex))
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
    
    strIDs = GetDeptDoctors(mlngUnit)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, i As Integer, Curdate As Date
    Dim strPreRoom As String, intRoom As Integer, intCheck As Integer, lngNurseGrade As Long
    Dim strSQL As String, strBed As String, strTmp As String, strNewSql As String
    Dim str���� As String, blnTrans As Boolean, strMainBed As String
    Dim rsTmp As ADODB.Recordset
    Dim str����� As String
    Dim blnTrue As Boolean
    Dim intĸӤת�Ʊ�־ As Integer
    
    If cbo����.ListIndex = -1 Then
        MsgBox "��ȷ������Ҫ��ס�Ŀ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    
    If cbo����.ListIndex = -1 Then
        MsgBox "��ָ�����˵ĵ�ǰ������", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    
    If cbo����ȼ�.ListIndex = -1 And gbln���ȷ������ȼ� Then
        MsgBox "��ָ�����˵ĵ�ǰ����ȼ���", vbInformation, gstrSysName
        cbo����ȼ�.SetFocus: Exit Sub
    End If
    
    '72433:������,2014-08-02
    blnTrue = (Val(zlDatabase.GetPara("��סָ��ҽ��С��", glngSys, glngModul, 0)) = 1) And cboҽ��С��.Enabled
    If cboҽ��С��.ItemData(cboҽ��С��.ListIndex) = 0 And blnTrue = True Then
        MsgBox "��ָ�����˵ĵ�ǰҽ��С�飡", vbInformation, gstrSysName
        If cboҽ��С��.Enabled And cboҽ��С��.Visible Then cboҽ��С��.SetFocus
        Exit Sub
    End If
    blnTrue = False
    
    If cbo����ȼ�.ListIndex <> -1 Then
        lngNurseGrade = cbo����ȼ�.ItemData(cbo����ȼ�.ListIndex)
    End If

    '78877:����ʱ�䲻�ܴ�����Ժʱ��
    If txtIn.Enabled And txtIn.Visible Then
        If CDate(mrsPatiInfo!�������� & "") > CDate(txtIn.Text) Then
            MsgBox "���˵���Ժʱ��[" & Format(txtIn.Text, "YYYY-MM-DD HH:MM:SS") & "]������ڲ��˵ĳ�������[" & mrsPatiInfo!�������� & "]��", vbInformation, gstrSysName
            txtIn.SetFocus
            Exit Sub
        End If
    End If

    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    Curdate = zlDatabase.Currentdate
    If InStr(Trim(cbo����.Text), " ����") <> 0 Then
        str���� = Mid(Trim(cbo����.Text), 1, InStr(Trim(cbo����.Text), " ����") - 1)
        
        If InStr(Trim(cbo����.Text), "|") - InStr(Trim(cbo����.Text), "����:") - 3 > 0 Then
            str����� = Mid(Trim(cbo����.Text), InStr(Trim(cbo����.Text), "����:") + 3, InStr(Trim(cbo����.Text), "|") - InStr(Trim(cbo����.Text), "����:") - 3)
        End If
        strSQL = "Select �Ա� From ������Ϣ A,��λ״����¼ B  Where A.����ID = b.����id And b.����ID Is Not Null And ����ID = [1] And ����� =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit, str�����)
        
        Do While Not rsTmp.EOF
            If Trim(txt�Ա�.Text) <> rsTmp!�Ա� Then
                If (MsgBox("ָ����λ���ڷ��������Ů��ס������Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                    cbo����.SetFocus
                End If
            End If
            rsTmp.MoveNext
        Loop
    ElseIf InStr(cbo����.Text, "��ͥ����") > 0 Then
        str���� = ""
    Else
        str���� = Trim(cbo����.Text)
    End If
    
    If CDate(txtDate.Text) > Curdate Then
        MsgBox "��סʱ������˵�ǰϵͳʱ��,���飡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    If mbyt��ס��ʽ <> 1 And txtIn.Visible Then
        If IsDate(txtIn.Text) Then
            If CDate(txtIn.Text) > Curdate Then
                MsgBox "��Ժʱ������˵�ǰϵͳʱ�䣬���飡", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
            If CDate(txtIn.Text) > CDate(txtDate.Text) Then
                MsgBox "��סʱ�䲻��С����Ժʱ�䣬���飡", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
        Else
            MsgBox "��Ժʱ������������飡", vbInformation, gstrSysName
            txtIn.SetFocus: Exit Sub
        End If
    End If
    
    If mbyt��ס��ʽ = 0 Then
        If Format(txtDate.Text, "yyyyMMddhhmmss") < Format(mrsPatiInfo!��Ժʱ��, "yyyyMMddHHmmss") Then
            MsgBox "��סʱ�䲻��С�ڸò��˵���Ժʱ��[" & Format(mrsPatiInfo!��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "]��", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    '��������Ժʱ����ͬ
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If mbyt��ס��ʽ = 1 Then
        If Format(txtDate.Text, "yyyyMMddhhmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
            MsgBox "��סʱ�������ڸò��˵��ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    Else
        If Format(txtDate.Text, "yyyyMMddhhmmss") < Format(dMax, "yyyyMMddHHmmss") Then
            MsgBox "��סʱ�䲻��С�ڸò��˵��ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    
    If chk����.Value = 1 Then
        strPreRoom = "һ����ͬ"
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                intCheck = intCheck + 1
                strTmp = lvw.ListItems(i).Text
                If InStr(1, strTmp, ":") > 0 Then   'ð�ź��Ƿ����
                    strTmp = Mid(strTmp, InStr(1, strTmp, ":") + 1)
                    If strTmp <> strPreRoom Then
                        intRoom = intRoom + 1
                        strPreRoom = strTmp
                    End If
                End If
            End If
        Next
        If intCheck < 2 Then
            MsgBox "�������˱�������������ϵĴ�λ��", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
        If intRoom > 1 Then
            MsgBox "��������������Ĵ�λ������һ�������ڣ�", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
    End If
    
    
    If cbo����.ItemData(cbo����.ListIndex) <> mrsPatiInfo!��ס����id Then
        '����28432 by lesfeng 2010-03-10
        If mint�������� = 0 And mbyt��ס��ʽ = 0 Then
            MsgBox "��ǰѡ��Ŀ��ҡ�" & zlCommFun.GetNeedName(cbo����.Text) & "�����ǲ���ԭ�ȵǼǵĿ��ҡ�" & mrsPatiInfo!��ǰ���� & "�������ܲ�����", vbInformation, gstrSysName
            If cbo����.Enabled Then cbo����.SetFocus
            Exit Sub
        Else
            If MsgBox("��ǰѡ��Ŀ��ҡ�" & zlCommFun.GetNeedName(cbo����.Text) & "�����ǲ���ԭ�ȵǼǵĿ��ҡ�" & mrsPatiInfo!��ǰ���� & "��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    
    If chk����.Value = 0 Then
        strBed = str����
        strMainBed = str����
    Else
        strMainBed = str����
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                strBed = strBed & "," & Mid(lvw.ListItems(i).Key, 2)
            End If
        Next
        strBed = Mid(strBed, 2)
    End If
    On Error GoTo errH
    intĸӤת�Ʊ�־ = 1
    '�Ƚ���ĸӤ���룬�ȴ�λ��Ϣ������ɺ�������
    If mbyt��ס��ʽ <> 0 And 1 = 0 Then
        strSQL = "Select Count(1) as Ӥ���� From ������������¼ Where ����id = [1] And ��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp!Ӥ���� > 0 Then
            '��Ӥ����ת��ʱ��ʾ�Ƿ�Ҫת��
            strSQL = "Select Ӥ������id, Ӥ������id From ������ҳ Where ����id = [1] And ��ҳid = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If rsTmp!Ӥ������ID & "" = "" Then
                'Ϊnull��ʾĸӤ��δ����
                If MsgBox("��ǰ��������������¼���������Ƿ�һ����ס��", vbQuestion + vbDefaultButton1 + vbYesNo) = vbYes Then
                    intĸӤת�Ʊ�־ = 1
                Else
                    intĸӤת�Ʊ�־ = 0
                End If
            End If
        End If
    End If
    
    strSQL = "zl_���˱䶯��¼_InDept(" & mlng����ID & "," & mlng��ҳID & ",'" & strBed & "'," & _
            mlngUnit & "," & cbo����.ItemData(cbo����.ListIndex) & "," & cboҽ��С��.ItemData(cboҽ��С��.ListIndex) & "," & _
            lngNurseGrade & ",'" & zlCommFun.GetNeedName(cbo����.Text) & "'," & _
            "'" & zlCommFun.GetNeedName(cbo���λ�ʿ.Text) & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "'," & _
            "'" & zlCommFun.GetNeedName(cboסԺҽʦ.Text) & "'," & chk���.Value & "," & _
            "To_Date('" & IIf(txtIn.Text = "____-__-__ __:__", mrsPatiInfo!��Ժʱ��, txtIn.Text) & "','YYYY-MM-DD HH24:MI:SS')," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(mbyt��ס��ʽ = 0, 1, 0) & ",'" & _
            zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "','" & strMainBed & "','" & intĸӤת�Ʊ�־ & "')"
    
    
    
    
    blnTrue = False
    strNewSql = " Select Count(*) ��¼" & vbNewLine & _
        "  From סԺ���ü�¼" & vbNewLine & _
        "  Where ����id = [1] And ��ҳid = [2] And ��¼���� = 3 And ��¼״̬ = 1 And ���ӱ�־ = 8"
    Set rsTmp = zlDatabase.OpenSQLRecord(strNewSql, "��鲡���Ƿ���������һ�ε�һ�η���", mlng����ID, mlng��ҳID)
    blnTrue = (Val(NVL(rsTmp!��¼, 0)) > 0)
    
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '36454,������,2012-09-06,û�м���������һ����Ŀ��סʱ���м���
    If mbyt��ס��ʽ <> 1 And blnTrue = False Then
         '�����Ժ�Ǽ�ʱû��ȷ������,��ʱmbyt��ס��ʽ��ȷ������,��Ҫ������Ժһ�η���
         '�����л��Զ��ж��Ƿ��Ѽ����(���ӱ�־=8,��¼����=3)
        strSQL = "ZL_סԺһ�η���_Insert(" & mlng����ID & "," & mlng��ҳID & ")"

        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
     
    If Val("" & mrsPatiInfo!����) <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng����ID, mlng��ҳID, Val("" & mrsPatiInfo!����), "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '����96847
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    mstr���� = strBed
    gblnOK = True
          
    On Error Resume Next
    '��Ƴɹ����Ǵ�����Ϣ
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txt����.Text, xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", NVL(mrsPatiInfo!סԺ��), xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        If mbyt��ס��ʽ = 0 Then '�������
            strSQL = " Select A.ID,B.���� ��λ�ȼ�,C.���� ��������  From  ���˱䶯��¼ A,�շ���ĿĿ¼ B,���ű� C" & _
                " Where NVl(A.���Ӵ�λ,0)=0 And A.��λ�ȼ�id=B.id(+) And A.����Id=C.id(+) And A.����ID=[1] And A.��ҳID=[2] And A.��ʼԭ��=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, 2)
            
            'סԺ��Ϣ
            mclsXML.AppendNode "in_hospital"
            'in_date     ��Ժʱ��    1   s
            mclsXML.appendData "in_date", Format(IIf(txtIn.Text = "____-__-__ __:__", mrsPatiInfo!��Ժʱ��, txtIn.Text), "yyyy-MM-dd HH:mm:ss"), xsString '��Ժ����
            'in_area_id      ��Ժ����id  0..1    N
            mclsXML.appendData "in_area_id", mlngUnit, xsNumber '��Ժ����ID
            'in_area_title       ��Ժ����    0..1    S
            mclsXML.appendData "in_area_title", NVL(rsTmp!��������), xsString  '��Ժ����
            'in_dept_id      ��Ժ����id  1   N
            mclsXML.appendData "in_dept_id", cbo����.ItemData(cbo����.ListIndex), xsNumber '��Ժ����id
            'in_dept_title       ��Ժ����    1   S
            mclsXML.appendData "in_dept_title", zlCommFun.GetNeedName(cbo����.Text), xsString  '��Ժ����
            mclsXML.appendData "in_again", Val(NVL(mrsPatiInfo!����Ժ, 0)), xsNumber
            mclsXML.AppendNode "in_hospital", True
            '��ס���
            mclsXML.AppendNode "dept_arrange"
            'change_id       �䶯id  1   N
            mclsXML.appendData "change_id", rsTmp!ID, xsNumber '�䶯ID
            'in_room     ��ס����    0..1    S
            mclsXML.appendData "in_room", str�����, xsString
            'in_bed      ��ס����    1   S
            mclsXML.appendData "in_bed", strMainBed, xsString
            'in_tendgrade        ����ȼ�    0..1    S
            If cbo����ȼ�.ListIndex <> -1 Then
                mclsXML.appendData "in_tendgrade", zlCommFun.GetNeedName(cbo����ȼ�.Text), xsString
            Else
                mclsXML.appendData "in_tendgrade", "", xsString
            End If
            'in_bedgrade     ��λ�ȼ�    0..1    S
            mclsXML.appendData "in_bedgrade", NVL(rsTmp!��λ�ȼ�), xsString
            'in_doctor       סԺҽʦ    0..1    S
            mclsXML.appendData "in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'duty_nurse      ���λ�ʿ    0..1    S
            mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
            mclsXML.AppendNode "dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_002", mclsXML.XmlText
        Else 'ת�����
            strSQL = " Select A.ID,B.���� ��λ�ȼ�,C.���� ��������  From  ���˱䶯��¼ A,�շ���ĿĿ¼ B,���ű� C" & _
                " Where NVl(A.���Ӵ�λ,0)=0 And A.��λ�ȼ�id=B.id(+) And A.����Id=C.id(+) And A.����ID=[1] And A.��ҳID=[2] And A.��ʼԭ��=[3] And ��ʼʱ��+0=[4]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, 3, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
            
            'סԺ��Ϣ
            mclsXML.AppendNode "in_hospital"
            'in_date     ��Ժʱ��    1   s
            mclsXML.appendData "in_date", Format(mrsPatiInfo!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), xsString
            'out_area_id     ת������id  0..1    N
            mclsXML.appendData "out_area_id", Val(NVL(mrsPatiInfo!��ǰ����ID)), xsNumber
            'out_area_title      ת������    0..1    S
            mclsXML.appendData "out_area_title", NVL(mrsPatiInfo!��ǰ����), xsString
            'out_dept_id     ת������id  1   N
            mclsXML.appendData "out_dept_id", Val(NVL(mrsPatiInfo!��Ժ����id, 0)), xsNumber
            'out_dept_title      ת������    1   S
            mclsXML.appendData "out_dept_title", NVL(mrsPatiInfo!��ǰ����), xsString
            'in_area_id      ת�벡��id  0..1    N
            mclsXML.appendData "in_area_id", mlngUnit, xsNumber
            'in_area_title       ת�벡��    0..1    S
            mclsXML.appendData "in_area_title", NVL(rsTmp!��������), xsString
            'in_dept_id      ת�����id  1   N
            mclsXML.appendData "in_dept_id", cbo����.ItemData(cbo����.ListIndex), xsNumber
            'in_dept_title       ת�����    1   S
            mclsXML.appendData "in_dept_title", zlCommFun.GetNeedName(cbo����.Text), xsString
            mclsXML.AppendNode "in_hospital", True
            'ת�����
            mclsXML.AppendNode "change_dept_arrange"
            'change_id       �䶯id  1   N
            mclsXML.appendData "change_id", rsTmp!ID, xsNumber '�䶯ID
            'in_room     ��ס����    0..1    S
            mclsXML.appendData "in_room", str�����, xsString
            'in_bed      ��ס����    1   S
            mclsXML.appendData "in_bed", strMainBed, xsString
            'in_tendgrade        ����ȼ�    0..1    S
            If cbo����ȼ�.ListIndex <> -1 Then
                mclsXML.appendData "in_tendgrade", zlCommFun.GetNeedName(cbo����ȼ�.Text), xsString
            Else
                mclsXML.appendData "in_tendgrade", "", xsString
            End If
            'in_bedgrade     ��λ�ȼ�    0..1    S
            mclsXML.appendData "in_bedgrade", NVL(rsTmp!��λ�ȼ�), xsString
            'in_doctor       סԺҽʦ    0..1    S
            mclsXML.appendData "in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'duty_nurse      ���λ�ʿ    0..1    S
            mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_012", mclsXML.XmlText
        End If
    End If
    If Err <> 0 Then Err.Clear
    
    '������ҽӿ�
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInBranchAfter(mlng����ID, mlng��ҳID)
        Call zlPlugInErrH(Err, "InPatiCheckInBranchAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    '49854:������,2013-10-31,���������ӡ
    'ֻ������ס�Ĳ��˲Ŵ�ӡ���
    If InStr(mstrPrivs, ";�����ӡ;") And mbyt��ס��ʽ <> 1 Then
        blnTrue = True
        If gbytCourseWristletPrint = 0 Then
            blnTrue = False
        Else
            If gbytCourseWristletPrint = 2 Then
                If MsgBox("�Ƿ��ӡ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnTrue = False
                End If
            End If
        End If
        
        If blnTrue = True Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, 2)
        End If
    End If
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'����28432 by lesfeng 2010-03-10
Private Function GetDeptName(ByVal lngID As Long) As String

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ���ű� Where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        GetDeptName = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByRef str���� As String, ByVal lng��λ����ID As Long, ByVal byt��ס��ʽ As Byte, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr���� = str����
    mlng��λ����ID = lng��λ����ID
    mbyt��ס��ʽ = byt��ס��ʽ
    mstrPrivs = strPrivs
    mstrAppointBed = ""
    mblnAppoint = False
    
    Me.Show 1, frmParent
    str���� = mstr����
    ShowMe = gblnOK
End Function

