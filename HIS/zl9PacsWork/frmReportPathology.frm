VERSION 5.00
Begin VB.Form frmReportPathology 
   BorderStyle     =   0  'None
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame frmCell 
      Caption         =   "ϸ����Ŀ��"
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   7695
      Begin VB.TextBox txt����Ƥϸ������ 
         Height          =   300
         Left            =   3480
         TabIndex        =   8
         Top             =   320
         Width           =   1300
      End
      Begin VB.TextBox txt��״��Ƥϸ������ 
         Height          =   300
         Left            =   3480
         TabIndex        =   9
         Top             =   697
         Width           =   1300
      End
      Begin VB.TextBox txt��֢ϸ�� 
         Height          =   300
         Left            =   6000
         TabIndex        =   11
         Top             =   690
         Width           =   1300
      End
      Begin VB.TextBox txt��״ϸ�� 
         Height          =   300
         Left            =   6000
         TabIndex        =   10
         Top             =   315
         Width           =   1300
      End
      Begin VB.ComboBox cbo����ϸ�� 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0000
         Left            =   960
         List            =   "frmReportPathology.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   697
         Width           =   800
      End
      Begin VB.ComboBox cbo����ϸ�� 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0016
         Left            =   960
         List            =   "frmReportPathology.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   320
         Width           =   800
      End
      Begin VB.Label Label12 
         Caption         =   "����Ƥϸ�����䣺"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   345
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "��״��Ƥϸ�����䣺"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "����ϸ����"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "����ϸ����"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "��֢ϸ����"
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "��״ϸ����"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.Frame frmMicrobe 
      Caption         =   "΢������Ϣ��"
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cboHPV��Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":002C
         Left            =   6720
         List            =   "frmReportPathology.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo�������Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0042
         Left            =   3960
         List            =   "frmReportPathology.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo��˾���Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0058
         Left            =   1320
         List            =   "frmReportPathology.frx":0062
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   800
      End
      Begin VB.ComboBox cbo���߾���Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":006E
         Left            =   6720
         List            =   "frmReportPathology.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   800
      End
      Begin VB.ComboBox cbo�������Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":0084
         Left            =   3960
         List            =   "frmReportPathology.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   800
      End
      Begin VB.ComboBox cbo�γ��Ⱦ 
         Height          =   300
         ItemData        =   "frmReportPathology.frx":009A
         Left            =   1320
         List            =   "frmReportPathology.frx":00A4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   800
      End
      Begin VB.Label Label6 
         Caption         =   "HPV��Ⱦ��"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "�������Ⱦ��"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "��˾���Ⱦ��"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   743
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "���߾���Ⱦ��"
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "�������Ⱦ��"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "�γ��Ⱦ��"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   383
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmReportPathology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Private mlngAdviceID As Long    'ҽ��ID
Private mlngReportID As Long    '�����ļ���
Private mintEditType As Integer '����״̬ 0 ������1��д��2 �޶�
Private mblnCheckModity As Boolean      '�Ƿ����������޸ļ�¼
Private mblnEditable As Boolean         '�Ƿ���Ա༭����
Private mblnMoved As Boolean            '�Ƿ��Ѿ�ת��

'����ר�Ʊ���Ҫ��
Private Const Report_Element_�γ��Ⱦ = "�γ��Ⱦ"
Private Const Report_Element_��˾���Ⱦ = "��˾���Ⱦ"
Private Const Report_Element_�������Ⱦ = "�������Ⱦ"
Private Const Report_Element_�������Ⱦ = "�������Ⱦ"
Private Const Report_Element_���߾���Ⱦ = "���߾���Ⱦ"
Private Const Report_Element_HPV��Ⱦ = "HPV��Ⱦ"
Private Const Report_Element_����ϸ�� = "����ϸ��"
Private Const Report_Element_����ϸ�� = "����ϸ��"
Private Const Report_Element_����Ƥϸ������ = "����Ƥϸ������"
Private Const Report_Element_��״��Ƥϸ������ = "��״��Ƥϸ������"
Private Const Report_Element_��״ϸ�� = "��״ϸ��"
Private Const Report_Element_��֢ϸ�� = "��֢ϸ��"

Public pModified As Boolean     '��¼�Ƿ����޸�

Public Sub zlRefresh(frmParentReport As frmReport, ByVal lngAdviceID As Long, lngReportID As Long, _
    blnSingleWindow As Boolean, blnEditable As Boolean, ByVal blnMoved As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    mblnSingleWindow = blnSingleWindow
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    
    '��ʼ���ؼ�
    Call InitControls
    
    mblnCheckModity = False         '�ر������޸ļ�¼
    pModified = False
    
    strSql = "Select �����ı�,Ҫ������ From ���Ӳ������� Where �ļ�ID=[1] And ��������=4 And ��ֹ��=0 And �滻��=0"
    If mblnMoved = True Then
        strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngReportID)
    
    While rsTemp.EOF = False
        Select Case Nvl(rsTemp!Ҫ������)
            Case Report_Element_�γ��Ⱦ
                cbo�γ��Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_��˾���Ⱦ
                cbo��˾���Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_�������Ⱦ
                cbo�������Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_�������Ⱦ
                cbo�������Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_���߾���Ⱦ
                cbo���߾���Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_HPV��Ⱦ
                cboHPV��Ⱦ.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_����ϸ��
                cbo����ϸ��.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_����ϸ��
                cbo����ϸ��.Text = Nvl(rsTemp!�����ı�, "��")
            Case Report_Element_����Ƥϸ������
                txt����Ƥϸ������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��״��Ƥϸ������
                txt��״��Ƥϸ������.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��״ϸ��
                txt��״ϸ��.Text = Nvl(rsTemp!�����ı�)
            Case Report_Element_��֢ϸ��
                txt��֢ϸ��.Text = Nvl(rsTemp!�����ı�)
        End Select
        rsTemp.MoveNext
    Wend
    
    '���ý���ؼ��Ƿ���Ա༭
    frmMicrobe.Enabled = mblnEditable
    frmCell.Enabled = mblnEditable
    
    mblnCheckModity = True         '���������޸ļ�¼
End Sub

Public Function getElementString() As String
    Dim strElements As String
    
    strElements = SPLITER_REPORT & Report_Element_�γ��Ⱦ & SPLITER_ELEMENT & cbo�γ��Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_��˾���Ⱦ & SPLITER_ELEMENT & cbo��˾���Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_�������Ⱦ & SPLITER_ELEMENT & cbo�������Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_�������Ⱦ & SPLITER_ELEMENT & cbo�������Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_���߾���Ⱦ & SPLITER_ELEMENT & cbo���߾���Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_HPV��Ⱦ & SPLITER_ELEMENT & cboHPV��Ⱦ.Text & SPLITER_REPORT & _
                    Report_Element_����ϸ�� & SPLITER_ELEMENT & cbo����ϸ��.Text & SPLITER_REPORT & _
                    Report_Element_����ϸ�� & SPLITER_ELEMENT & cbo����ϸ��.Text & SPLITER_REPORT & _
                    Report_Element_����Ƥϸ������ & SPLITER_ELEMENT & txt����Ƥϸ������.Text & SPLITER_REPORT & _
                    Report_Element_��״��Ƥϸ������ & SPLITER_ELEMENT & txt��״��Ƥϸ������.Text & SPLITER_REPORT & _
                    Report_Element_��״ϸ�� & SPLITER_ELEMENT & txt��״ϸ��.Text & SPLITER_REPORT & _
                    Report_Element_��֢ϸ�� & SPLITER_ELEMENT & txt��֢ϸ��.Text
    getElementString = strElements
End Function

Private Sub InitControls()
    cbo�γ��Ⱦ.ListIndex = 1
    cbo��˾���Ⱦ.ListIndex = 1
    cbo�������Ⱦ.ListIndex = 1
    cbo�������Ⱦ.ListIndex = 1
    cbo���߾���Ⱦ.ListIndex = 1
    cboHPV��Ⱦ.ListIndex = 1
    cbo����ϸ��.ListIndex = 1
    cbo����ϸ��.ListIndex = 1
    
    txt����Ƥϸ������.Text = ""
    txt��״��Ƥϸ������.Text = ""
    txt��״ϸ��.Text = ""
    txt��֢ϸ��.Text = ""
End Sub

Private Sub cboHPV��Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cboHPV��Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�γ��Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�γ��Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo���߾���Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo���߾���Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo����ϸ��_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo����ϸ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo����ϸ��_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo����ϸ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�������Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�������Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�������Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo�������Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo��˾���Ⱦ_DropDown()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub cbo��˾���Ⱦ_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    
    '�ڷſؼ�λ��
    If Me.Width > 7500 Then    '�ڳ�2��
        '΢������Ϣ
        Label3.Left = cbo�������Ⱦ.Left + cbo�������Ⱦ.Width + 500
        Label3.Top = Label2.Top
        cbo���߾���Ⱦ.Left = Label3.Left + Label3.Width + 50
        cbo���߾���Ⱦ.Top = cbo�������Ⱦ.Top
        
        Label6.Left = Label3.Left
        Label6.Top = Label5.Top
        cboHPV��Ⱦ.Left = cbo���߾���Ⱦ.Left
        cboHPV��Ⱦ.Top = cbo�������Ⱦ.Top
        
        'ϸ����Ŀ
        Label7.Left = txt����Ƥϸ������.Left + txt����Ƥϸ������.Width + 300
        Label7.Top = Label12.Top
        txt��״ϸ��.Left = Label7.Left + Label7.Width + 50
        txt��״ϸ��.Top = txt����Ƥϸ������.Top
        
        Label8.Left = txt��״��Ƥϸ������.Left + txt��״��Ƥϸ������.Width + 300
        Label8.Top = Label11.Top
        txt��֢ϸ��.Left = txt��״ϸ��.Left
        txt��֢ϸ��.Top = txt��״��Ƥϸ������.Top
    Else        '�ڳ�3��
        '΢������Ϣ
        Label3.Left = Label4.Left
        Label3.Top = Label4.Top + Label4.Height + 200
        cbo���߾���Ⱦ.Left = cbo��˾���Ⱦ.Left
        cbo���߾���Ⱦ.Top = Label3.Top - 50
        
        Label6.Left = Label5.Left
        Label6.Top = Label5.Top + Label5.Height + 200
        cboHPV��Ⱦ.Left = cbo�������Ⱦ.Left
        cboHPV��Ⱦ.Top = Label6.Top - 50
        
        'ϸ����Ŀ
        Label7.Left = Label10.Left
        Label7.Top = Label10.Top + Label10.Height + 200
        txt��״ϸ��.Left = Label7.Left + Label7.Width '+ 10
        txt��״ϸ��.Top = Label7.Top - 50
        
        Label8.Left = txt��״ϸ��.Left + txt��״ϸ��.Width + 200
        Label8.Top = Label11.Top + Label11.Height + 200
        txt��֢ϸ��.Left = Label8.Left + Label8.Width ' + 10
        txt��֢ϸ��.Top = Label8.Top - 50
    End If
    
    '�ڷ����
    frmMicrobe.Left = 0
    frmMicrobe.Top = 0
    lngTemp = Me.Width - 100
    frmMicrobe.Width = IIf(lngTemp > 0, lngTemp, 0)
    frmMicrobe.Height = Me.cboHPV��Ⱦ.Top + Me.cboHPV��Ⱦ.Height + 100
    
    frmCell.Left = 0
    frmCell.Top = frmMicrobe.Height + 50
    frmCell.Width = frmMicrobe.Width
    frmCell.Height = Me.txt��֢ϸ��.Top + Me.txt��֢ϸ��.Height + 100
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    
    SaveSetting "ZLSOFT", strRegPath, "CY22", Me.Height
End Sub

Private Sub txt��״��Ƥϸ������_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt��״��Ƥϸ������_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��״��Ƥϸ������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��״ϸ��_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt��״ϸ��_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��״ϸ��_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt����Ƥϸ������_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt����Ƥϸ������_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����Ƥϸ������_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txt��֢ϸ��_Change()
    If mblnCheckModity = True Then
        pModified = True
    End If
End Sub

Private Sub txt��֢ϸ��_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��֢ϸ��_LostFocus()
    Call zlCommFun.OpenIme
End Sub
