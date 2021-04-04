VERSION 5.00
Begin VB.Form frmSendInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���汨�ͱ�ע"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   Icon            =   "frmSendInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   14
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   13
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "�걨��Ϣ"
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6300
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   650
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1000
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1725
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���͵�λ��"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ�䣺"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "������Ա��"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��    ��:"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.TextBox txtContent 
      Height          =   3420
      Left            =   3360
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   3060
   End
   Begin VB.TextBox txtSuggestion 
      Height          =   3420
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   3060
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�������˵��"
      Height          =   180
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   720
   End
End
Attribute VB_Name = "frmSendInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '1-����;2-�޸�
Private mlng�걨ID As Long
Private mdatTime As Date
Private mblnChange As Boolean

Private mstrOld���� As String
Private mstrOld����˵�� As String

Public Enum TextInfoEnum
    txt_���� = 0
    txt_������Ա = 1
    txt_����ʱ�� = 2
    txt_���͵�λ = 3
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal intType As Integer, ByVal lngID As Long, Optional ByVal datTime As Date) As Boolean
    mintType = intType
    mlng�걨ID = lngID
    mdatTime = datTime
    
    Me.Show 1, frmParent
End Function

Private Sub SetFormState()
    Dim blnLock As Boolean
    
    blnLock = IIf(mintType = 3, True, False)
    txtSuggestion.Locked = blnLock
    txtContent.Locked = blnLock
End Sub

Private Sub loadData()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
     
    On Error GoTo errH
    If mintType = 2 Or mintType = 3 Then
        strSQL = "Select a.������Ϣ, a.�Ǽ���, a.�������˵��, b.����, b.������, b.����ʱ��, b.���͵�λ" & vbNewLine & _
                "From �����걨���� A, �����걨��¼ B" & vbNewLine & _
                "Where a.�걨id = b.�ļ�id And a.�걨id = [1] And a.�Ǽ�ʱ�� = [2]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�걨ID, mdatTime)
        
        If rsTmp.RecordCount > 0 Then
            mstrOld���� = NVL(rsTmp!������Ϣ)
            mstrOld����˵�� = NVL(rsTmp!�������˵��)
            txtSuggestion.Text = mstrOld����
            txtContent.Text = mstrOld����˵��
            
            txtInfo(txt_����) = NVL(rsTmp!����)
            txtInfo(txt_������Ա) = NVL(rsTmp!������)
            txtInfo(txt_����ʱ��) = NVL(rsTmp!����ʱ��)
            txtInfo(txt_���͵�λ) = NVL(rsTmp!���͵�λ)
        End If
    ElseIf mintType = 1 Then
        strSQL = "Select  ����, ������, ����ʱ��, ���͵�λ From �����걨��¼ A Where �ļ�id = [1] "
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�걨ID)

        If rsTmp.RecordCount > 0 Then
            txtInfo(txt_����) = NVL(rsTmp!����)
            txtInfo(txt_������Ա) = NVL(rsTmp!������)
            txtInfo(txt_����ʱ��) = NVL(rsTmp!����ʱ��)
            txtInfo(txt_���͵�λ) = NVL(rsTmp!���͵�λ)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CheckData(False)
    If mintType = 1 Then
        If mblnChange Then
            If MsgBox("�Ѿ���д�����ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    ElseIf mintType = 2 Then
        If mblnChange Then
            If MsgBox("�����Ѿ������ı䣬�Ƿ�����޸ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Function CheckData(ByVal blnOK As Boolean) As Boolean
    Dim strSuggestion As String, strContent As String
    
    strSuggestion = Trim(txtSuggestion.Text)
    strContent = Trim(txtContent.Text)
    
    If strSuggestion <> mstrOld���� Or strContent <> mstrOld����˵�� Then
        mblnChange = True
    Else
        mblnChange = False
    End If
    
    If blnOK Then
        If strSuggestion = "" Then
            MsgBox "������Ϣ���ܹ�Ϊ�գ�������д������Ϣ��", vbInformation, gstrSysName
            txtSuggestion.SetFocus
            Exit Function
        ElseIf strContent = "" Then
            MsgBox "�������˵�����ܹ�Ϊ�գ�������д�������˵����", vbInformation, gstrSysName
            txtSuggestion.SetFocus
            Exit Function
        ElseIf LenB(StrConv(strSuggestion, vbFromUnicode)) > txtSuggestion.MaxLength Then
            MsgBox "������Ϣ���������" & txtSuggestion.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
            txtSuggestion.SetFocus: Exit Function
        ElseIf LenB(StrConv(strContent, vbFromUnicode)) > txtContent.MaxLength Then
            MsgBox "�������˵�����������" & txtContent.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
            txtContent.SetFocus: Exit Function
        End If
    End If
    CheckData = True
End Function

Private Sub cmdOK_Click()
    Dim strSuggestion As String, strContent As String, str�Ǽ�ʱ�� As String
    Dim str�Ǽ��� As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not CheckData(True) Then Exit Sub
    
    strSuggestion = "'" & Trim(txtSuggestion.Text) & "'"
    strContent = "'" & Trim(txtContent.Text) & "'"
    str�Ǽ��� = "'" & UserInfo.���� & "'"
    
    If mintType = 1 Then
        str�Ǽ�ʱ�� = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "zl_�����걨����_insert(" & mlng�걨ID & "," & strSuggestion & "," & str�Ǽ��� & "," & str�Ǽ�ʱ�� & "," & strContent & ")"
    ElseIf mintType = 2 Then
        str�Ǽ�ʱ�� = "to_date('" & Format(mdatTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "zl_�����걨����_Update(" & mlng�걨ID & "," & strSuggestion & "," & str�Ǽ��� & "," & str�Ǽ�ʱ�� & "," & strContent & ")"
    End If

    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    mstrOld���� = Trim(txtSuggestion.Text)
    mstrOld����˵�� = Trim(txtContent.Text)
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnChange = False
    mstrOld���� = ""
    mstrOld����˵�� = ""
    Call loadData
End Sub

Private Sub txtContent_GotFocus()
    Me.txtContent.SelStart = 0: Me.txtContent.SelLength = 500
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_LostFocus()
    Me.txtContent.Text = Replace(Me.txtContent, Chr(vbKeyReturn), "")
End Sub

Private Sub txtSuggestion_GotFocus()
    Me.txtSuggestion.SelStart = 0: Me.txtSuggestion.SelLength = 500
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtSuggestion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSuggestion_LostFocus()
    Me.txtSuggestion.Text = Replace(Me.txtSuggestion, Chr(vbKeyReturn), "")
End Sub
