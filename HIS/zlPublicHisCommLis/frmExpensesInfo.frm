VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmExpensesInfo 
   Caption         =   "������Ϣ"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpensesInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   13695
   StartUpPosition =   1  '����������
   Begin VB.Frame fraPati 
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cboShow 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmExpensesInfo.frx":6852
         Left            =   840
         List            =   "frmExpensesInfo.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   4815
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   7
         Left            =   7200
         TabIndex        =   18
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   4
         Left            =   9960
         TabIndex        =   17
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   5
         Left            =   7260
         TabIndex        =   16
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   3
         Left            =   4830
         TabIndex        =   15
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblCaption 
         Caption         =   "�����ˣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6660
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "������ң�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9000
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCaption 
         Caption         =   "���䣺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4230
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "���Σ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "�걾���ͣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   6300
         TabIndex        =   8
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   6
         Left            =   6840
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   2
         Left            =   2940
         TabIndex        =   6
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.PictureBox picRefresh 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10896
      Picture         =   "frmExpensesInfo.frx":6856
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "ˢ��(F5)"
      Top             =   312
      Width           =   480
   End
   Begin VB.PictureBox PicWindows 
      BorderStyle     =   0  'None
      Height          =   276
      Left            =   11496
      ScaleHeight     =   270
      ScaleWidth      =   510
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   516
   End
   Begin XtremeSuiteControls.TabControl TabCtlWindow 
      Height          =   5895
      Left            =   180
      TabIndex        =   0
      Top             =   1380
      Width           =   10545
      _Version        =   589884
      _ExtentX        =   18606
      _ExtentY        =   10393
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmExpensesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pҽ�����ѹ��� As Integer = 1257                       '���˷���ģ����Ȩ
Private Const p����ҽ���´� As Integer = 1252                       '����ҽ���´�
Private Const pסԺҽ���´� As Integer = 1253                       'סԺҽ���´�
Private Const p���ﲡ������ As Integer = 1250                       '���ﲡ��
Private Const pסԺ�������� As Integer = 1251
Private Const p�°没������ As Integer = 2250                       '�°没��
Private Const p�°没������_���� As Integer = 2251                  '�°没�������
Private Const p�°没������_סԺ As Integer = 2252                  '�°没����סԺ��

Private mcolSubForm As Collection                                   'ж���Ӵ���

Private mclsExpenses As Object                                          '���ö���
Private mclsOutAdvices As Object                                        '����ҽ������
Private mclsInAdvices As Object                                         'סԺҽ������
Private mclsOutEPRs As Object                                           '���ﲡ��
Private mclsInEPRs As Object                                            'סԺ����
Private mobjKernel As Object                                            'ҽ������
Private mclsEMR As Object                                               '�°���Ӳ���
Private mobjRichEPR As Object                                           '�������Ĳ���

Private mlngSapmeID As Long                                             '�걾ID
Private mrsInfo As New ADODB.Recordset                                  '������Ļ�����Ϣ
Private mblnLoadfrm As Boolean                                          '�Ƿ�������



Private Sub cboShow_Click()
        Call RefreshTab(TabCtlWindow.Selected.Index)
End Sub

Private Sub Form_Activate()
    gobjHisComLib.InitCommon gcnHisOracle
    gobjHisComLib.RegCheck
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 116 Then       'F5
        picRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngSysNo As Long
    Dim intIndex As Integer
    Dim strPrivs As String
    Dim strSQL As String, strTmp As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo Form_Load_Error

    mblnLoadfrm = False

    lngSysNo = 100

    '��ʼ�����Ĳ���
    Set mobjKernel = CreateObject("zlCISKernel.clsCISKernel")
    Set mobjRichEPR = CreateObject("zlRichEPR.cRichEPR")

    Call mobjKernel.InitCISKernel(gcnHisOracle, Me, lngSysNo, "")
    Call mobjRichEPR.InitRichEPR(gcnHisOracle, Me, lngSysNo, False)



    With Me.TabCtlWindow
        Set .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, pҽ�����ѹ���)  'û��ҽ�����ѹ���Ȩ��ʱ����ʾ
        .InsertItem(0, "���ò�ѯ", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���ò�ѯ", "")
        '        .Item(0).Visible = IIf(strPrivs <> "", True, False)
        .Item(0).Visible = False

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p����ҽ���´�)
        .InsertItem(1, "����ҽ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "����ҽ��", "")
        .Item(1).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, pסԺҽ���´�)
        .InsertItem(2, "סԺҽ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "סԺҽ��", "")
        .Item(2).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p���ﲡ������)
        .InsertItem(3, "���ﲡ��", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���ﲡ��", "")
        .Item(3).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, pסԺ��������)
        .InsertItem(4, "סԺ����", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "סԺ����", "")
        .Item(4).Visible = IIf(strPrivs <> "", True, False)

        If mrsInfo("������Դ") & "" = 2 Then
            strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p�°没������_סԺ)
        Else
            strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p�°没������_����)
        End If

        '�����°���Ӳ�������
        On Error Resume Next
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            'û���ӷ�����������
            .InsertItem(5, "���Ӳ���", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���Ӳ���", "")
            .Item(5).Visible = False
        Else
            Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")

            Err.Clear: On Error GoTo Form_Load_Error
            If mclsEMR Is Nothing Then
                strPrivs = ""
            Else

            End If
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
                If Not mclsEMR Is Nothing Then
                    mcolSubForm.Add mclsEMR.zlGetForm, "_���Ӳ���"
                End If
            End If
            .InsertItem(5, "���Ӳ���", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "���Ӳ���", "")
            .Item(5).Visible = IIf(strPrivs <> "", True, False)
        End If
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        lblInformation(1).Caption = mrsInfo("����")
        lblInformation(2).Caption = mrsInfo("�Ա�")
        lblInformation(3).Caption = mrsInfo("����")
        lblInformation(4).Caption = mrsInfo("�������")
        lblInformation(5).Caption = mrsInfo("������")
        lblInformation(7).Caption = mrsInfo("�걾����")
        If mrsInfo("������Դ") & "" = 2 Then
            strSQL = "Select rownum as ���,����id,��ҳID,NVL(��������,0) ��������,��ǰ����id,סԺ��,To_Char(��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ���� From ������ҳ Where ��ҳID<>0 And ����ID=[1] Order by ��ҳID Desc"
            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������Ϣ", Val(mrsInfo("����ID") & ""))
        Else
            strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ����,a.����ʱ�� From ���˹Һż�¼ A,���ű� B" & _
                   " Where A.ִ�в���ID=B.ID And A.����ID=[1] And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������Ϣ", Val(mrsInfo("����ID") & ""), Now)


        End If
        If rsTmp.RecordCount = 0 Then
            Exit Sub
        End If
        cboShow.Clear
        Do While Not rsTmp.EOF
            If mrsInfo("������Դ") & "" = 2 Then
                strTmp = "�� " & rsTmp!��ҳID & " ��"    '&  Decode(rsTmp!��������, 1, "(��������)", 2, "(סԺ����)", "")
                cboShow.AddItem strTmp
                cboShow.ItemData(cboShow.NewIndex) = rsTmp!��ҳID

            Else
                strTmp = Format(rsTmp!ʱ��, "YYMMdd") & "/" & rsTmp!���� & "/" & rsTmp!ִ����
                cboShow.AddItem strTmp
                cboShow.ItemData(cboShow.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Loop

        cboShow.ListIndex = 0

        ' Call cboShow_Click


        'ֻ��ʾ�����סԺ

        With Me.TabCtlWindow
            If mrsInfo("������Դ") & "" = 2 Then
                TabCtlWindow.Item(1).Visible = False
                .Item(2).Visible = True
                .Item(3).Visible = False
                .Item(4).Visible = True
            Else
                .Item(1).Visible = True
                .Item(2).Visible = False
                .Item(3).Visible = True
                .Item(4).Visible = False
            End If
        End With

        mblnLoadfrm = True

        'Ĭ�ϼ��ص�һ��û�����ص�ҳ��
        For intIndex = 0 To 5
            If .Item(intIndex).Visible = True Then
                .Item(intIndex).Selected = True
                Call RefreshTab(intIndex)
                Exit For
            End If
        Next

    End With



    Exit Sub
Form_Load_Error:
    Call WriteErrLog("zl9LisInsideComm", "frmExpensesInfo", "ִ��(Form_Load)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With fraPati
        .Top = 50
        .Left = 50
        .Width = Me.ScaleWidth - 100
    End With
    
    
    With Me.TabCtlWindow
        .Top = fraPati.Top + fraPati.Height
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = Me.ScaleHeight - fraPati.Height - 100
    End With
    
'    With Me.picRefresh
'        .Top = 100
'        .Left = Me.ScaleWidth - .Width - 100
'
'
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call mobjKernel.InitCISKernel(gcnLisOracle, Me, 100, "")
    If Not gobjEmr Is Nothing Then
        Call gobjEmr.CloseForms
    End If
    
    Set mcolSubForm = Nothing
    Set mclsExpenses = Nothing
    Set mclsInAdvices = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsOutEPRs = Nothing
    Set mclsInEPRs = Nothing
    Set mclsEMR = Nothing
    Set mobjKernel = Nothing
    Set mobjRichEPR = Nothing
    Set mrsInfo = Nothing
    mblnLoadfrm = False
    TabCtlWindow.RemoveAll
End Sub

Private Sub picRefresh_Click()
    Call RefreshTab(Me.TabCtlWindow.Selected.Index)
End Sub

Private Sub picRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picRefresh.BorderStyle = 1
End Sub

Private Sub picRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picRefresh.BorderStyle = 0
End Sub

Private Sub RefreshTab(intIndex As Integer)
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngMainID As Long, lngDeptID As Long
          Dim strData As String
          Dim lngSysNo As Long

1         On Error GoTo RefreshTab_Error

2         If mblnLoadfrm = False Then Exit Sub

3         mblnLoadfrm = False

4         strData = cboShow.ItemData(cboShow.ListIndex)

          'û�м�¼ʱ�˳�
5         If mrsInfo.RecordCount <= 0 Then Exit Sub

6         lngSysNo = 100

          'ֻ��ʾ�����סԺ
7         With Me.TabCtlWindow
8             If mrsInfo("������Դ") & "" = 2 Then
9                 TabCtlWindow.Item(1).Visible = False
10                .Item(2).Visible = True
11                .Item(3).Visible = False
12                .Item(4).Visible = True
13            Else
14                .Item(1).Visible = True
15                .Item(2).Visible = False
16                .Item(3).Visible = True
17                .Item(4).Visible = False
18            End If
19        End With

20        Select Case intIndex

          Case 0                                                                  '����
21            If mcolSubForm Is Nothing Then
22                Set mcolSubForm = New Collection
23            End If
24            If mclsExpenses Is Nothing Then
25                Set mclsExpenses = CreateObject("zlCISKernel.clsDockExpense")
26                mcolSubForm.Add mclsExpenses.zlGetForm, "_����"             '�õ��Ӵ���
27            End If
28            With Me.TabCtlWindow
29                If .Item(intIndex).Handle = PicWindows.hWnd Then
30                    .RemoveItem (intIndex)
31                    .InsertItem(intIndex, "���ò�ѯ", mcolSubForm("_����").hWnd, 0).Tag = "���ò�ѯ"
32                End If
33            End With

34            strSQL = "select a.id as ҽ��ID, b.���ͺ�,b.ִ�в���ID from ����ҽ����¼ a,����ҽ������ b " & vbCrLf & _
                     " Where a.ID = b.ҽ��id And a.���id = [1] "
35            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", Val(mrsInfo("����ID") & ""))
36            If rsTmp.EOF = False Then
37                mclsExpenses.zlRefresh rsTmp("ִ�в���ID"), rsTmp("ҽ��ID"), rsTmp("���ͺ�")
38            End If
39        Case 1
40            If mcolSubForm Is Nothing Then
41                Set mcolSubForm = New Collection
42            End If
43            If mclsOutAdvices Is Nothing Then
44                Set mclsOutAdvices = CreateObject("zlCISKernel.clsDockOutAdvices")
45                mcolSubForm.Add mclsOutAdvices.zlGetForm, "_����ҽ��"
46            End If
              '��һ�δ�ʱ�ټ���
47            With Me.TabCtlWindow
48                If .Item(intIndex).Handle = PicWindows.hWnd Then

49                    .RemoveItem (intIndex)
50                    .InsertItem(intIndex, "����ҽ��", mcolSubForm("_����ҽ��").hWnd, 1).Tag = "����ҽ��"
51                    .Item(intIndex).Selected = True
52                End If
53                strSQL = "select d.no �Һŵ�  from ���˹Һż�¼ d " & vbCrLf & _
                         " Where d.id = [1] "
54                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", strData)

55                mclsOutAdvices.zlRefresh Val(mrsInfo("����ID") & ""), rsTmp("�Һŵ�") & "", False
56                TabCtlWindow.Item(intIndex).Selected = True
57            End With
58        Case 2
59            If mcolSubForm Is Nothing Then
60                Set mcolSubForm = New Collection
61            End If
62            If mclsInAdvices Is Nothing Then
63                Set mclsInAdvices = CreateObject("zlCISKernel.clsDockInAdvices")
64                mcolSubForm.Add mclsInAdvices.zlGetForm, "_סԺҽ��"
65            End If
              '��һ�δ�ʱ�ټ���
66            With Me.TabCtlWindow
67                If .Item(intIndex).Handle = PicWindows.hWnd Then

68                    .RemoveItem (intIndex)
69                    .InsertItem(intIndex, "סԺҽ��", mcolSubForm("_סԺҽ��").hWnd, 1).Tag = "סԺҽ��"
70                    .Item(intIndex).Selected = True
71                End If
72            End With
73            strSQL = "Select a.��Ժ����id ���˿���ID  ,a.��ǰ����id ����ID From ������ҳ a where a.����id =[1] and a.��ҳid =[2]"
74            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", Val(mrsInfo("����ID") & ""), strData)
75            If rsTmp.EOF = False Then
76                mclsInAdvices.zlRefresh Val(mrsInfo("����ID") & ""), strData, Val(rsTmp("����ID") & ""), _
                                          Val(rsTmp("���˿���ID") & ""), 0
77            End If
78            TabCtlWindow.Item(intIndex).Selected = True
79        Case 3
80            If mcolSubForm Is Nothing Then
81                Set mcolSubForm = New Collection
82            End If
83            If mclsOutEPRs Is Nothing Then
84                Set mclsOutEPRs = CreateObject("zlRichEPR.cDockOutEPRs")
85                mcolSubForm.Add mclsOutEPRs.zlGetForm, "_���ﲡ��"
86            End If
              '��һ�δ�ʱ�ټ���
87            With Me.TabCtlWindow
88                If .Item(intIndex).Handle = PicWindows.hWnd Then

89                    .RemoveItem (intIndex)
90                    .InsertItem(intIndex, "���ﲡ��", mcolSubForm("_���ﲡ��").hWnd, 1).Tag = "���ﲡ��"
91                    .Item(intIndex).Selected = True
92                End If
93            End With
94            strSQL = "select a.id as ҽ��ID, b.���ͺ�,b.ִ�в���ID,c.����ID,a.���˿���ID,d.id �Һ�ID from ����ҽ����¼ a,����ҽ������ b,�������Ҷ�Ӧ c,���˹Һż�¼ d " & vbCrLf & _
                     " Where a.ID = b.ҽ��id and a.���˿���ID = ����ID(+) and a.�Һŵ� = d.no And d.id = [1] "
95            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", strData)
96            If rsTmp.EOF = False Then
97                mclsOutEPRs.zlRefresh Val(mrsInfo("����ID") & ""), rsTmp("�Һ�ID"), Val(rsTmp("���˿���ID") & ""), False
98            Else
99                mclsOutEPRs.zlRefresh 0, 0, 0, False
100           End If
101           TabCtlWindow.Item(intIndex).Selected = True
102       Case 4
103           If mcolSubForm Is Nothing Then
104               Set mcolSubForm = New Collection
105           End If
106           If mclsInEPRs Is Nothing Then
107               Set mclsInEPRs = CreateObject("zlRichEPR.cDockInEPRs")
108               mcolSubForm.Add mclsInEPRs.zlGetForm, "_סԺ����"
109           End If
              '��һ�δ�ʱ�ټ���
110           With Me.TabCtlWindow
111               If .Item(intIndex).Handle = PicWindows.hWnd Then
112                   .RemoveItem (intIndex)
113                   .InsertItem(intIndex, "סԺ����", mcolSubForm("_סԺ����").hWnd, 1).Tag = "סԺ����"
114                   .Item(intIndex).Selected = True
115               End If
116           End With
117           strSQL = "Select a.��Ժ����id ���˿���ID ,a.��ǰ����id  From ������ҳ a where a.����id =[1] and a.��ҳid =[2]"
118           Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", Val(mrsInfo("����ID") & ""), strData)
119           If rsTmp.EOF = False Then
120               mclsInEPRs.zlRefresh Val(mrsInfo("����ID") & ""), strData, Val(rsTmp("���˿���ID") & ""), False
121           Else
122               mclsInEPRs.zlRefresh 0, 0, 0, False
123           End If
124           TabCtlWindow.Item(intIndex).Selected = True
125       Case 5
126           If mcolSubForm Is Nothing Then
127               Set mcolSubForm = New Collection

128           End If
129           If mclsEMR Is Nothing Then
130               Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")
131               mcolSubForm.Add mclsEMR.zlGetForm, "_���Ӳ���"
132           End If
133           If Not mclsEMR Is Nothing Then
134               If Not mclsEMR.Init(gobjEmr, gcnHisOracle, lngSysNo) Then
135                   Set mclsEMR = Nothing
136               End If
137           End If
              '��һ�δ�ʱ�ټ���
138           With Me.TabCtlWindow
139               If .Item(intIndex).Handle = PicWindows.hWnd Then
140                   .RemoveItem (intIndex)
141                   .InsertItem(intIndex, "���Ӳ���", mcolSubForm("_���Ӳ���").hWnd, 1).Tag = "���Ӳ���"
142                   .Item(intIndex).Selected = True
143               End If
144           End With
145           If mrsInfo("������Դ") & "" = 2 Then
146               strSQL = "Select a.��Ժ����id ���˿���ID ,a.��ǰ����id  From ������ҳ a where a.����id =[2] and a.��ҳid =[1]"
147           Else
148               strSQL = "select a.id as ҽ��ID, b.���ͺ�,b.ִ�в���ID,c.����ID,a.���˿���ID,d.id �Һ�ID from ����ҽ����¼ a,����ҽ������ b,�������Ҷ�Ӧ c,���˹Һż�¼ d " & vbCrLf & _
                         " Where a.ID = b.ҽ��id and a.���˿���ID = ����ID(+) and a.�Һŵ� = d.no And d.id = [1] "
149           End If
150           Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�鿴������Ϣ", strData, Val(mrsInfo("����ID") & ""))

151           If rsTmp.RecordCount > 0 Then
152               mclsEMR.zlRefresh Val(mrsInfo("����ID") & ""), strData, Val(rsTmp("���˿���ID") & ""), 0, IIf(mrsInfo("������Դ") & "" = 2, 2, 1)
153           End If
154           TabCtlWindow.Item(intIndex).Selected = True
155       End Select

156       mblnLoadfrm = True

157       Exit Sub
RefreshTab_Error:
158       mblnLoadfrm = True

160       Call WriteErrLog("zl9LisInsideComm", "frmExpensesInfo", "ִ��(RefreshTab)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
161       Err.Clear
End Sub
'Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptID As Long, Optional ByVal bnEdit As Boolean, _
 '                            Optional ByVal blnMoved As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal lngAdviceID As Long) As Long
'    '����:����ˢ��ָ�����˵Ĳ������ݣ�����������ṩ�༭����
'    '����:  lngPatiId-����id;
'    '       lngBillId-�Һ�id;
'    '       lngDeptId-��ǰ�������ţ�ע�ⲻ�ǲ��˱��ξ�����ң�
'    '       blnEdit-�Ƿ�����༭��ͨ����ǰ�������Ų��ǲ��˱��ξ�����ң���Ӧ�ò�����༭��
'    '       blnMoved-�����Ƿ�ת��
'    '       lngAdviceID ҽ��ID��Ŀǰֻ������ģ����ô���
'    zlRefresh = frmOutEPRs.zlRefresh(lngPatiID, lngBillId, lngDeptID, bnEdit, blnForce, blnMoved, lngAdviceID)
'End Function
Public Sub ShowMe(lngSapmeID, objEMR As Object, parfrom As Object)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    gobjHisComLib.InitCommon gcnHisOracle
    gobjHisComLib.RegCheck
    Set gobjEmr = objEMR

    mlngSapmeID = lngSapmeID

    strSQL = "select id,HIS����ID ����ID,������Դ,����ID,ҽ��ID,�����,סԺ��,��ҳID,�Һŵ�,���˿��ұ���,��������,������ұ���,����, decode(�Ա�,1,'��',2,'Ů','δ֪') �Ա�,����,������,�������,�걾���� from ����������� where �걾id = [1] Order By ҽ��id "

    Set mrsInfo = ComOpenSQL(Sel_Lis_DB, strSQL, "", mlngSapmeID)

    If mrsInfo.RecordCount <= 0 Then
        Unload Me
        MsgBox "û���ҵ���ǰ�걾��������Ϣ,����!", vbInformation, "�鿴����"
        Exit Sub
    Else
        If mrsInfo("������Դ") & "" = "" Or mrsInfo("ҽ��ID") & "" = "" Then
            MsgBox "�ֹ����벡�ˣ����ܲ鿴������Ϣ", vbInformation, "�鿴����"
            Exit Sub
        End If
    End If
    If mrsInfo("������Դ") & "" = 2 Then
        strSQL = "Select rownum as ���,����id,��ҳID,NVL(��������,0) ��������,��ǰ����id,סԺ��,To_Char(��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժ���� From ������ҳ Where ��ҳID<>0 And ����ID=[1] Order by ��ҳID Desc"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������Ϣ", Val(mrsInfo("����ID") & ""))
    Else
        strSQL = "Select A.ID,A.NO,A.����ʱ�� as ʱ��,B.���� as ����,a.ִ����,a.����ʱ�� From ���˹Һż�¼ A,���ű� B" & _
               " Where A.ִ�в���ID=B.ID And A.����ID=[1] And A.����ʱ��<=[2] And A.��¼����=1 And A.��¼״̬=1 Order by A.����ʱ�� Desc,a.����ʱ�� Desc"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������Ϣ", Val(mrsInfo("����ID") & ""), Now)


    End If
    If rsTmp.RecordCount = 0 Then
        Unload Me
        MsgBox "�ò���û���ҵ�������Ϣ!", vbInformation, "�鿴����"
        Exit Sub
    End If
    Me.Show


End Sub

Private Sub TabCtlWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
        Call RefreshTab(Item.Index)
End Sub
