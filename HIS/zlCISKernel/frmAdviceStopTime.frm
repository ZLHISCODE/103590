VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceStopTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ֹͣҽ��"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4170
   Icon            =   "frmAdviceStopTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraTZYY 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   1305
      TabIndex        =   6
      Top             =   990
      Width           =   2205
      Begin VB.TextBox txtTZYY 
         Height          =   300
         Left            =   0
         MaxLength       =   200
         TabIndex        =   1
         Top             =   240
         Width           =   1800
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   265
         Left            =   1800
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblTZYY 
         AutoSize        =   -1  'True
         Caption         =   "ִ����ֹԭ��"
         Height          =   180
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -255
      TabIndex        =   4
      Top             =   1800
      Width           =   4845
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2625
      TabIndex        =   3
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1485
      TabIndex        =   2
      Top             =   1920
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   645
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   232980483
      UpDown          =   -1  'True
      CurrentDate     =   39668.3388888889
   End
   Begin VB.Image imgCharge 
      Height          =   240
      Left            =   240
      Picture         =   "frmAdviceStopTime.frx":058A
      Top             =   1170
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgAudit 
      Height          =   720
      Left            =   360
      Picture         =   "frmAdviceStopTime.frx":6DDC
      Stretch         =   -1  'True
      Top             =   255
      Width           =   720
   End
   Begin VB.Label lblBT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ִ����ֹʱ��"
      Height          =   180
      Left            =   1320
      TabIndex        =   5
      Top             =   375
      Width           =   1080
   End
   Begin VB.Image ImgStop 
      Height          =   720
      Left            =   360
      Picture         =   "frmAdviceStopTime.frx":7166
      Top             =   255
      Width           =   720
   End
End
Attribute VB_Name = "frmAdviceStopTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsAdvice As ADODB.Recordset
Private mlngҽ��ID As Long
Private mblnOK As Boolean

Private mstrTime As String
Private mintMode As Integer '0-ҽ��ֹͣ��1��ҽ����˵Ĵ��壬2����ʿ��Һ��ҩ��¼����
Private mdatRegister As Date
Private mstrԭ�� As String '�� mintMode=0 ҽ��ֹͣʱҪ��¼��ֹͣԭ��
Private mlng����ID As Long

Public Function ShowMe(frmParent As Object, ByVal lngҽ��ID As Long, ByVal lng����ID As Long, Optional ByVal intMode As Integer = 0, Optional ByVal datRegister As Date = 0, Optional ByRef strԭ�� As String) As String
     '******************************************************************************************************************
    '������intMode,Ϊ1�Ļ���ʾ�ǵ���ҽ����˵Ĵ���
    '      datRegister,ҽ��ִ�еĵǼ�ʱ��
    '˵��������ѡ���ʱ����ַ���
    '******************************************************************************************************************
    mlngҽ��ID = lngҽ��ID
    mlng����ID = lng����ID
    mintMode = intMode
    mdatRegister = datRegister
    Me.Show 1, frmParent
    If mblnOK Then
        strԭ�� = mstrԭ��
        ShowMe = mstrTime
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
   '���Ϸ���
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    
    If mintMode = 0 Then
        '������ڿ�ʼִ��ʱ��
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") <= Format(mrsAdvice!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm") Then
            MsgBox "�����ִ����ֹʱ��������ҽ���Ŀ�ʼִ��ʱ�� " & Format(mrsAdvice!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm") & "��", vbInformation, gstrSysName
            dtpTime.SetFocus: Exit Sub
        End If
        '�Ǽ�ִ��ʱ��>�ϴ�ִ��ʱ��
        mstrTime = GetAdviceStopTime(mlngҽ��ID)
        If mstrTime <> "" Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mstrTime, "yyyy-MM-dd HH:mm") Then
                MsgBox "����ֹͣ��ִ��ʱ�� " & mstrTime & " ֮ǰ�������ֹͣʱ�䣬���ȷʵҪֹͣ��ִ��ʱ��֮ǰ������ȡ��ִ�еǼǡ�", vbInformation, gstrSysName
                dtpTime.SetFocus: Exit Sub
            End If
        End If
        '��ӦС���ϴ�ִ��ʱ��
        If Not IsNull(mrsAdvice!�ϴ�ִ��ʱ��) Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm") Then
                If MsgBox("�����ִ����ֹʱ��С��ҽ�����ϴ�ִ��ʱ�� " & Format(mrsAdvice!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    dtpTime.SetFocus: Exit Sub
                End If
            End If
        End If
        
        'δ��д��ֹԭ��
        If gblnҽ����ֹԭ�� Then
            If Trim(txtTZYY.Text) = "" And InStr(gstr�ɲ���ͣ��ԭ�����, "," & mlng����ID & ",") = 0 Then
                MsgBox "��¼����ֹԭ��", vbInformation, gstrSysName
                txtTZYY.SetFocus: Exit Sub
            If zlCommFun.ActualLen(txtTZYY.Text) > txtTZYY.MaxLength Then
                    MsgBox "��ֹԭ������̫����������� " & txtTZYY.MaxLength / 2 & " �����ֻ� " & txtTZYY.MaxLength & " ���ַ���", vbInformation, gstrSysName
                    txtTZYY.SetFocus: Exit Sub
                End If
            End If
            mstrԭ�� = Trim(txtTZYY.Text)
        End If
    ElseIf mintMode = 1 Then
        '�������ִ��ʱ��
        If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mdatRegister, "yyyy-MM-dd HH:mm") Then
            MsgBox "����ĺ˶�ʱ�䲻�ܹ�С��ҽ��ִ�еĵǼ�ʱ�� " & Format(mdatRegister, "yyyy-MM-dd HH:mm") & "��", vbExclamation, gstrSysName
            dtpTime.SetFocus: Exit Sub
        End If
    ElseIf mintMode = 2 Then
        If Trim(txtTZYY.Text) = "" Then
            MsgBox "��¼������ԭ��", vbInformation, gstrSysName
            txtTZYY.SetFocus: Exit Sub
        If zlCommFun.ActualLen(txtTZYY.Text) > txtTZYY.MaxLength Then
                MsgBox "����ԭ������̫����������� " & txtTZYY.MaxLength / 2 & " �����ֻ� " & txtTZYY.MaxLength & " ���ַ���", vbInformation, gstrSysName
                txtTZYY.SetFocus: Exit Sub
            End If
        End If
        mstrԭ�� = Trim(txtTZYY.Text)
    End If
    mstrTime = Format(dtpTime.value, "yyyy-MM-dd HH:mm")
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click()
'���ܣ�����ѡ����
    Call GetItemԭ��(1)
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    If dtpTime.Enabled Then dtpTime.SetFocus
    Me.Refresh
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
    zlCommFun.PressKey vbKeyRight
End Sub

Private Sub Form_Load()
    Dim datCurr As Date
    Dim strSQL As String
    
    mblnOK = False
    datCurr = zlDatabase.Currentdate
    
    On Error GoTo errH
    If mintMode = 0 Then
        ImgAudit.Visible = False
        ImgStop.Visible = True
        fraTZYY.Visible = gblnҽ����ֹԭ��
        
        Set Me.Icon = ImgStop.Picture
        
        lblBT.Caption = "ִ����ֹʱ��"
        Me.Caption = "ֹͣҽ��"
        
        strSQL = "Select ��ʼִ��ʱ��,ִ����ֹʱ��,�ϴ�ִ��ʱ��,����ʱ�� From ����ҽ����¼ Where ID=[1]"
        Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
        
        If gbln����ҽ��������Ч Then
            dtpTime.value = CDate(Format(datCurr + 1, "yyyy-MM-dd 00:00"))
        Else
            dtpTime.value = CDate(Format(datCurr, "yyyy-MM-dd HH:mm"))
        End If
        
        If Not IsNull(mrsAdvice!�ϴ�ִ��ʱ��) Then
            If Format(dtpTime.value, "yyyy-MM-dd HH:mm") < Format(mrsAdvice!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm") Then
                dtpTime.value = Format(mrsAdvice!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm")
            End If
        End If
    ElseIf mintMode = 1 Then
        ImgAudit.Visible = True
        ImgStop.Visible = False
        fraTZYY.Visible = False
        Set Me.Icon = ImgAudit.Picture
        
        lblBT.Caption = "�˶�ʱ��"
        Me.Caption = "ҽ���˶�"
        dtpTime.value = CDate(Format(datCurr, "yyyy-MM-dd HH:mm"))
    ElseIf mintMode = 2 Then
        Set Me.Icon = imgCharge.Picture
        Me.Caption = "��Һ��ҩ��¼����"
        lblTZYY.Caption = "����ԭ��"
        Set ImgAudit.Picture = imgCharge.Picture
        ImgStop.Visible = False
        dtpTime.Enabled = False
        strSQL = "select ����˵�� from ����ҽ��״̬ where ҽ��id=[1] and ��������=8 and ����˵�� is not null"
        Set mrsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
        If Not mrsAdvice.EOF Then
            txtTZYY.Text = mrsAdvice!����˵�� & ""
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If mintMode = 1 Or mintMode = 0 And Not gblnҽ����ֹԭ�� Then
        fraLine.Top = 1800 - fraTZYY.Height
        cmdOK.Top = 1920 - fraTZYY.Height
        cmdCancel.Top = 1920 - fraTZYY.Height
        Me.Height = 2775 - fraTZYY.Height
    ElseIf mintMode = 2 Then
        fraTZYY.Top = lblBT.Top
        
        fraLine.Top = 1800 - fraTZYY.Height
        cmdOK.Top = 1920 - fraTZYY.Height
        cmdCancel.Top = 1920 - fraTZYY.Height
        Me.Height = 2775 - fraTZYY.Height
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
        Set mrsAdvice = Nothing
    End If
End Sub

Private Sub GetItemԭ��(ByVal intType As Integer)
'���ܣ�ѡ��ͣ��ԭ��
'������intType =0 KeyPress���ã�=1 ������ť����
    Dim strSQL As String, rsTmp As Recordset
    Dim blnCancel As Boolean, vRect As RECT
    Dim strMatch As String
    Dim strInput As String
    
    On Error GoTo errH
    
    If intType = 0 Then
       strInput = txtTZYY.Text
       If IsNumeric(strInput) Then '10,11.����ȫ������ʱֻƥ�����
           If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " where  A.���� Like [1]"
       ElseIf zlCommFun.IsCharAlpha(strInput) Then '01,11.����ȫ����ĸʱֻƥ�����
           If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " where  a.���� Like [1]"
       ElseIf zlCommFun.IsCharChinese(strInput) Then
           strMatch = " where  a.���� Like [1]"
       End If
    End If
    
    strSQL = "select a.���� as id, a.����,a.����,a.���� from ͣ��ԭ�� a  " & strMatch & " order by a.����"
    vRect = zlControl.GetControlRect(txtTZYY.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, Me.Caption, False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtTZYY.Height, blnCancel, False, True, UCase(txtTZYY.Text) & "%")

    If Not rsTmp Is Nothing Then
''        If Not blnCancel Then
''            MsgBox "δ�ҵ�ƥ�����Ŀ��", vbInformation, gstrSysName
''        End If
''        Call zlControl.TxtSelAll(txtTZYY)
''        txtTZYY.SetFocus: Exit Sub
'    Else
        txtTZYY.Text = rsTmp!���� & ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtTZYY_GotFocus()
    Call zlControl.TxtSelAll(txtTZYY)
End Sub

Private Sub txtTZYY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GetItemԭ��(0)
    Else
        If KeyAscii = 39 Then KeyAscii = 0 '������
    End If
End Sub
