VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdvicePrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ҽ������ӡ"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmAdvicePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Tag             =   "�����ӡ"
      Top             =   720
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Frame fraClear 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3465
         TabIndex        =   24
         Top             =   4320
         Width           =   2385
         Begin VB.TextBox txtClearPage 
            Height          =   270
            Left            =   945
            MaxLength       =   3
            TabIndex        =   26
            Top             =   45
            Width           =   510
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "���(&D)"
            Height          =   350
            Left            =   1485
            TabIndex        =   25
            Top             =   0
            Width           =   800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�����ʼҳ"
            Height          =   180
            Left            =   0
            TabIndex        =   27
            Top             =   80
            Width           =   900
         End
      End
      Begin VB.CheckBox chkSeqPage 
         Caption         =   "�ش򡰴�����ҳ"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   4020
         TabIndex        =   22
         Top             =   585
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   2
         Left            =   2655
         MouseIcon       =   "frmAdvicePrint.frx":058A
         Picture         =   "frmAdvicePrint.frx":0B14
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   1
         Left            =   1395
         MouseIcon       =   "frmAdvicePrint.frx":11FE
         Picture         =   "frmAdvicePrint.frx":1788
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":1E72
         Height          =   360
         Index           =   0
         Left            =   195
         MouseIcon       =   "frmAdvicePrint.frx":255C
         Picture         =   "frmAdvicePrint.frx":2AE6
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblPrint 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdvicePrint.frx":31D0
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   12
         Top             =   90
         Width           =   3600
      End
      Begin VB.Label lblStopPrint 
         AutoSize        =   -1  'True
         Caption         =   "���ѣ�"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   100
         TabIndex        =   11
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ѵ�ӡ       ������        δ��ӡ"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   585
         Width           =   2970
      End
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����"
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&F)"
      Height          =   315
      Left            =   5355
      TabIndex        =   28
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdClsLastPrint 
      Caption         =   "����ϴδ�ӡ(&C)"
      Height          =   350
      Left            =   1800
      TabIndex        =   23
      Top             =   5730
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   4950
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar vsc 
      Height          =   1815
      Left            =   6090
      SmallChange     =   50
      TabIndex        =   19
      Top             =   2130
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox chkH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   5895
      ScaleHeight     =   810
      ScaleWidth      =   1320
      TabIndex        =   13
      Top             =   -75
      Visible         =   0   'False
      Width           =   1320
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1245
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   15
         Top             =   525
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.PictureBox picPaperB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1365
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   495
         Picture         =   "frmAdvicePrint.frx":321A
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmAdvicePrint.frx":3C1C
         Top             =   465
         Width           =   240
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   16
         Top             =   855
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.OptionButton optReport 
      Caption         =   "����ҽ����"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton optReport 
      Caption         =   "��ʱҽ����"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboBaby 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3510
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   75
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   5940
      TabIndex        =   8
      Top             =   5730
      Width           =   800
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   945
      TabIndex        =   7
      Top             =   5730
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   100
      TabIndex        =   6
      Top             =   5730
      Width           =   800
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   5200
      Left            =   105
      TabIndex        =   3
      Top             =   400
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�����ӡ"
            Key             =   "�����ӡ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ���״�"
            Key             =   "ҽ���״�"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Tag             =   "ҽ���״�"
      Top             =   960
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":41A6
         Height          =   360
         Index           =   3
         Left            =   210
         MouseIcon       =   "frmAdvicePrint.frx":4890
         Picture         =   "frmAdvicePrint.frx":4E1A
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "�״�"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lblInSidePrint 
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ���״�ָ���Ѵ�ӡ��ҽ�����϶�У��/ֹͣ/ȷ��ͣ�����״��뵥�������ͼƬѡ��Ҫ�״��ҽ����ҳ�š�"
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   10
         Top             =   135
         Width           =   4320
      End
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ��������13ҳ��"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3630
      TabIndex        =   21
      Top             =   5820
      Width           =   1800
   End
   Begin VB.Label lblBaby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��"
      Height          =   180
      Left            =   3090
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAdvicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ӿڲ�����
Private mfrmParent As Object
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstrDefKey As String 'ȱʡ��λ���Ĵ�ӡ����

'ģ�����
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mrsPrint As ADODB.Recordset
Private mintPrintCount As Integer
Attribute mintPrintCount.VB_VarHelpID = -1
Private mblnTrans As Boolean '�����¼�֮��������Ƕ�׿���
Private mlngPrintType As Long   'ҽ����ӡģʽ   1-�¿����ӡ��0-У�Ժ��ӡ

Private mlngRows���� As Long    '��1ҳֽ�ϴ�ӡ��������
Private mlngRows���� As Long    '
Private mintMid As Integer      '�ش�ҳ��Ӧ��ҳ���ٽ�ҳ��mintMid  ��δ��ҳ/����ҳ��
Private mbln����ҳ As Boolean   '�Ƿ��д�����ҳ
Private mstrTurnPages As String '���еĻ�ҳ��ӡ��ҳ�ţ���ʽ "2,3,6,8,9"
Private mlngPrintedMaxPage As Long ' �Ѿ���ӡ����ҽ�������ҳ��
Private mlngPage����ǰ As Long '���һ������ǰ��ӡ�������ҳ��

Private mintPageCount As Integer        '��ҳ��  �����ӡ��ҳ����ֻҪ���봰����������ǹ̶��ġ�
Private mintStopPageCount As Integer    '��ҳ��  ͣ����ӡ��ҳ����ֻ��С��ִ���״���
Private mdat����ʱ�� As Date
Private mint��ҩ���г��� As Integer
Private mint��ҩ�������� As Integer

Private Enum mCtlID
    optҽ��_���� = 0
    optҽ��_���� = 1
    
    fra����_���� = 0
    fra����_�״� = 1
    
    optλ��_���� = 0
    optλ��_���� = 1
    optλ��_���� = 2
    
    lblͼ��˵��_���� = 0
    lblͼ��˵��_�״� = 1
    
    img�Ѵ� = 0
    img���� = 1
    imgδ�� = 2
    img�״� = 3
    
    pic����_���� = 1
    pic����_ֽ�� = 2
    
    pic�״�_���� = 3
    pic�״�_ֽ�� = 4
    
End Enum

Public Sub ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal strDefKey As String)
    Set mfrmParent = frmParent
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrDefKey = strDefKey
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdPrintSet_Click()
'�����ӡ����
    Dim strReport As String
    strReport = IIF(optReport(optҽ��_����).value, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, strReport, Me)
End Sub

Private Sub cmdRefresh_Click()
'ˢ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    Dim i As Long
    Dim lng��Ч As Long
    Dim lngӤ�� As Long
    Dim arrSQL As Variant
    
    On Error GoTo errH
 
    Set mrsPrint = Nothing
    mbln����ҳ = False
    mintPageCount = 0
    mintStopPageCount = 0
 
    lngӤ�� = cboBaby.ListCount - 1
    lng��Ч = IIF(optReport(optҽ��_����).value, 0, 1)
    
    arrSQL = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & lngӤ�� & "," & lng��Ч & ",null,null,3)"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_����ҽ����ӡ_Insert(" & mlng����ID & "," & mlng��ҳID & "," & lngӤ�� & "," & lng��Ч & "," & IIF(lng��Ч = 0, mlngRows����, mlngRows����) & ")"
    
    '�ύ����
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    '�ж��Ƿ�Ҫ���ֽ��
    strSQL = "select max(a.ҳ��) as ҳ�� from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    lngTmp = Val(rsTmp!ҳ�� & "") * 2
    If Val(picPaper(0).Tag) < lngTmp Then
        For i = Val(picPaper(0).Tag) + 1 To lngTmp
            Call LoadPaper(0, i)
            Call LoadPaper(1, i)
        Next
    End If
    Call tbsMain_Click
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim arrBaby As Variant, strBaby As String
    Dim blnPriv As Boolean, i As Long
 
    mblnTrans = False
    '���ñ���Ȩ��
    '����ҽ����
    blnPriv = False
    If InStr(UserInfo.����, "ҽ��") > 0 Then
        If InStr(GetInsidePrivs(pסԺҽ���´�), "����ҽ����") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.����, "��ʿ") > 0 Then
        If InStr(GetInsidePrivs(pסԺҽ������), "����ҽ����") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(optҽ��_����).value = True
        optReport(optҽ��_����).Enabled = False
    End If
    
    '��ʱҽ����
    blnPriv = False
    If InStr(UserInfo.����, "ҽ��") > 0 Then
        If InStr(GetInsidePrivs(pסԺҽ���´�), "��ʱҽ����") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.����, "��ʿ") > 0 Then
        If InStr(GetInsidePrivs(pסԺҽ������), "��ʱҽ����") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(optҽ��_����).value = True
        optReport(optҽ��_����).Enabled = False
    End If
    
    '�����������������Ӧ������һ����Ȩ��
    If Not optReport(optҽ��_����).Enabled And Not optReport(optҽ��_����).Enabled Then
        Unload Me: Exit Sub
    End If
    
    '��ʼ��Ӥ��ѡ��
    cboBaby.AddItem "����ҽ��"
    Call Cbo.SetIndex(cboBaby.hwnd, 0)
    
    strBaby = GetBabyRegList(mlng����ID, mlng��ҳID)
    
    If strBaby <> "" Then
        arrBaby = Split(strBaby, "<Split>")
        For i = 0 To UBound(arrBaby)
            cboBaby.AddItem "Ӥ�� " & i + 1 & IIF(arrBaby(i) <> "", "��" & arrBaby(i), "")
        Next
    Else
        lblBaby.Visible = False
        cboBaby.Visible = False
    End If
    Call Cbo.SetListWidth(cboBaby.hwnd, cboBaby.Width * 1.55)
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    mdat����ʱ�� = GetRsRedoDate(mlng����ID, mlng��ҳID)
    
    'ҽ����ӡģʽ
    mlngPrintType = Val(zldatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�))
    
    mlngRows���� = GetReportRows(glngSys, "ZL1_INSIDE_1254_2")
    mlngRows���� = GetReportRows(glngSys, "ZL1_INSIDE_1254_1")
    
    mint��ҩ���г��� = Val(zldatabase.GetPara("��������ҩҽ��������ʾ����", glngSys, pסԺҽ������))
    mint��ҩ�������� = Val(zldatabase.GetPara("��������ҩҽ��������ʾ����", glngSys, pסԺҽ������))
    
    Call Insert��ӡ��¼
    
    Call LoadAllPaper
    
    'ˢ�½�������
    If mstrDefKey <> "" And tbsMain.SelectedItem.Key <> mstrDefKey Then
        tbsMain.Tag = "NoneClick"
        For i = 1 To tbsMain.Tabs.Count
            If tbsMain.Tabs(i).Key = mstrDefKey Then
                tbsMain.Tabs(i).Selected = True
                Exit For
            End If
        Next
        tbsMain.Tag = ""
    End If
     
    Call tbsMain_Click
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    Me.Height = 6600
    
    With tbsMain
        .Left = 100
        .Top = 400
        .Width = Me.ScaleWidth - 200
        .Height = 5200
    End With
    
    cmdPreview.Top = Me.ScaleHeight - 470
    cmdPrint.Top = cmdPreview.Top
    cmdCancel.Top = cmdPreview.Top
    cmdCancel.Left = tbsMain.Width + tbsMain.Left - cmdCancel.Width
    cmdClsLastPrint.Top = cmdPreview.Top
    lblTotal.Top = cmdPreview.Top + 100
    fraClear.Visible = cmdClsLastPrint.Visible
    For i = 0 To 2
        fraPrint(i).Top = 750
        fraPrint(i).Left = 150
        fraPrint(i).Width = tbsMain.Width - 400
        fraPrint(i).Height = tbsMain.Height - 430
    Next
    
    
    For i = 0 To 3
        imgIcon(i).Top = 530
    Next
    
    imgIcon(img�״�).Left = imgIcon(img�Ѵ�).Left
    
    lblPrintIcoInfo(lblͼ��˵��_����).Top = 670
    lblPrintIcoInfo(lblͼ��˵��_�״�).Top = 670
    lblStopPrint.Top = 4500
    lblStopPrint.Left = 100
    fraClear.Top = lblStopPrint.Top - 80
    fraClear.Left = tbsMain.Width - fraClear.Width - 350
 
    lblPrint.Left = lblInSidePrint.Left
    lblInSidePrint.Top = lblPrint.Top
    cmdRefresh.Top = cboBaby.Top
    cmdRefresh.Left = cmdCancel.Left + cmdCancel.Width - cmdRefresh.Width
    
    cmdPrintSet.Top = cmdRefresh.Top
    cmdPrintSet.Left = cmdRefresh.Left - cmdPrintSet.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False '�Է���һ
    Set mrsPrint = Nothing
    mbln����ҳ = False
    Call UnLoadPaper
    mintPageCount = 0
    mintStopPageCount = 0
End Sub

Private Sub cboBaby_Click()
    Call RefreshFace
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Call AdvicePrint(1)
    Call RefreshFace
End Sub

Private Sub cmdPrint_Click()
    Call AdvicePrint(2)
    Call RefreshFace
End Sub

Private Sub optReport_Click(Index As Integer)
    If Not Visible Then Exit Sub
    Call tbsMain_Click
End Sub

Private Sub tbsMain_Click()
    Dim i As Long
    
    If tbsMain.Tag = "NoneClick" Then Exit Sub
    
    For i = 0 To fraPrint.UBound
        fraPrint(i).Visible = fraPrint(i).Tag = tbsMain.SelectedItem.Key
        If fraPrint(i).Tag = tbsMain.SelectedItem.Key Then
            fraPrint(i).ZOrder
            If i = fra����_���� Then picContainer(pic����_����).ZOrder
            If i = fra����_�״� Then picContainer(pic�״�_����).ZOrder
        End If
    Next
    Call RefreshFace
    picContainer(pic�״�_ֽ��).Top = 0
    picContainer(pic����_ֽ��).Top = 0
    vsc.value = 0
    cmdRefresh.Enabled = tbsMain.SelectedItem.Key = "�����ӡ"
End Sub

Private Sub RefreshFace()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim intIco As Integer '0 - �Ѵ�ӡ 1 - ������ 2 �� ����ӡ
    Dim intIndex As Integer
    Dim lngTmp As Long
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "�����ӡ" Then
        strSQL = "select m.ҳ��,sum(m.��ӡ) as ��ӡ,sum(m.δ��ӡ) as δ��ӡ,count(1) as ����" & vbNewLine & _
            "from (select a.ҳ��,decode(a.��ӡʱ��,null,0,1) as ��ӡ,decode(a.��ӡʱ��,null,1,0) as δ��ӡ" & vbNewLine & _
            "from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2] and nvl(a.Ӥ��,0)=[3] and a.��Ч=[4] and �к�>0) m" & vbNewLine & _
            "group by m.ҳ�� order by m.ҳ��"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex, IIF(optReport(optҽ��_����).value, 0, 1))
        
        mintMid = 0
        mlngPrintedMaxPage = 0
        mstrTurnPages = ""
        mbln����ҳ = False
        mintPageCount = rsTmp.RecordCount
        chkSeqPage.Visible = False
        For i = 1 To rsTmp.RecordCount

            If Val(rsTmp!��ӡ & "") > 0 And Val(rsTmp!δ��ӡ & "") = 0 Then
                intIco = 0
                mlngPrintedMaxPage = Val(rsTmp!ҳ�� & "")
            ElseIf Val(rsTmp!��ӡ & "") > 0 And Val(rsTmp!δ��ӡ & "") > 0 Then
                If 0 = mintMid Then
                    mintMid = Val(rsTmp!ҳ�� & "")
                    mbln����ҳ = True
                    chkSeqPage.Visible = True
                End If
                intIco = 1
                mlngPrintedMaxPage = Val(rsTmp!ҳ�� & "")
            ElseIf Val(rsTmp!��ӡ & "") = 0 And Val(rsTmp!δ��ӡ & "") > 0 Then
                If 0 = mintMid Then mintMid = Val(rsTmp!ҳ�� & "")
                intIco = 2
            End If
            
            '���һҳ����������ҳ
            If i <> rsTmp.RecordCount Then
                If Val(rsTmp!���� & "") < IIF(optReport(0).value, mlngRows����, mlngRows����) Then
                    mstrTurnPages = mstrTurnPages & "," & rsTmp!ҳ��
                End If
            End If
      
            Set imgIco(i).Picture = imgIcon(intIco).Picture
            imgIco(i).ToolTipText = IIF(intIco = 0, "�Ѵ�ӡ", IIF(intIco = 1, "������", "δ��ӡ"))
            imgChk(i).Visible = IIF(intIco = 0, False, True)
            picPaper(i).Visible = True
            picPaperB(i).Visible = True
            rsTmp.MoveNext
        Next
        
        '�ж��Ƿ�Ӧ����ʾ����ϴδ�ӡ��ť
        rsTmp.Filter = "��ӡ>0"
        cmdClsLastPrint.Visible = Not rsTmp.EOF
        For i = mintPageCount + 1 To Val(picPaper(0).Tag)
            imgChk(i).Visible = False
            picPaper(i).Visible = False
            picPaperB(i).Visible = False
        Next
        
        If mstrTurnPages <> "" Then mstrTurnPages = Mid(mstrTurnPages, 2)
        
        Set rsTmp = GetStopedAdvice(True)
        
        If mlngPrintType = 1 Then
            If optReport(optҽ��_����).value Then
                lblStopPrint.Caption = "��У��/ֹͣ/ȷ��ֹͣ��ҽ����Ҫ��ӡ��"
            Else
                lblStopPrint.Caption = "��У�Ե�ҽ����Ҫ��ӡ��"
            End If
        Else
            lblStopPrint.Caption = "��ȷ��ֹͣ��ҽ����Ҫ��ӡ��"
        End If
            
        lblStopPrint.Visible = rsTmp.RecordCount > 0
         
        cmdPreview.Enabled = mintPageCount > 0
        cmdPrint.Enabled = mintPageCount > 0
        
        lngTmp = IntEx(mintPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
 
        picContainer(pic����_ֽ��).Height = lngTmp * 3450
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        
        If mintPageCount = 0 Then
            lblTotal.Caption = IIF(optReport(optҽ��_����).value, "����", "��ʱ") & "ҽ�������ޡ�"
        Else
            lblTotal.Caption = IIF(optReport(optҽ��_����).value, "����", "��ʱ") & "ҽ��������" & mintPageCount & "ҳ��"
        End If
        lblTotal.Visible = True
        If mlngPrintedMaxPage <> 0 Then txtClearPage.Text = mlngPrintedMaxPage
    ElseIf tbsMain.SelectedItem.Key = "ҽ���״�" Then
        
        Set rsTmp = GetStopedAdvice(False)
        
        mintStopPageCount = rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            lblNum(i + 1000).Caption = Val(rsTmp!ҳ�� & "")
            lblNum(i + 1000).ToolTipText = "��" & Val(rsTmp!ҳ�� & "") & "ҳ"
            picPaper(i + 1000).Visible = True
            picPaperB(i + 1000).Visible = True
            imgChk(i + 1000).Visible = False
            rsTmp.MoveNext
        Next
        
        For i = mintStopPageCount + 1 To Val(picPaperB(0).Tag)
            imgChk(i + 1000).Visible = False
            picPaper(i + 1000).Visible = False
            picPaperB(i + 1000).Visible = False
        Next
        
        cmdPrint.Enabled = mintStopPageCount <> 0
        cmdPreview.Enabled = mintStopPageCount <> 0
        
        lngTmp = IntEx(mintStopPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
        picContainer(pic�״�_ֽ��).Height = lngTmp * 3450
        
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        If mintStopPageCount = 0 Then
            lblTotal.Caption = "�״�ҽ�������ޡ�"
        Else
            lblTotal.Caption = "�״�ҽ��������" & mintStopPageCount & "ҳ��"
        End If
        lblTotal.Visible = True
        cmdClsLastPrint.Visible = False
    ElseIf tbsMain.SelectedItem.Key = "��ӡѡ��" Then
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        vsc.Visible = False
        lblTotal.Visible = False
        cmdClsLastPrint.Visible = False
    End If
    fraClear.Visible = cmdClsLastPrint.Visible
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetStopedAdvice(ByVal blnOnlyCheckExists As Boolean) As ADODB.Recordset
'���ܣ���ȡ��ǰ������Ҫͣ����ӡ�ļ�¼��
'������blnOnlyCheckExists-ֻ����Ƿ����ͣ����ӡ
    Dim strSQL As String
    
    If optReport(optҽ��_����).value Then
        If blnOnlyCheckExists Then
            strSQL = "Select 1 From ����ҽ����ӡ A, ����ҽ����¼ B Where A.ҽ��id = B.ID And A.��Ч = 0 And A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And a.��ӡʱ�� is not null and (B.ȷ��ͣ��ʱ�� Is Not Null And" & vbNewLine & _
                "     Not Exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� = 2) " & _
                IIF(mlngPrintType = 1, "Or B.ִ����ֹʱ�� Is Not Null And Not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in (1,2))  or b.У��ʱ�� is not null and not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in(1,2,3))", "") & ") And Rownum<2"
        Else
            strSQL = _
                "Select Distinct ҳ��" & vbNewLine & _
                "From (Select A.ҽ��id, Max(A.ҳ��) As ҳ��" & vbNewLine & _
                "       From ����ҽ����ӡ A, ����ҽ����¼ B" & vbNewLine & _
                "       Where A.ҽ��id = B.ID And A.��Ч = 0 And A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And a.��ӡʱ�� is not null And (B.ȷ��ͣ��ʱ�� Is Not Null And" & vbNewLine & _
                "             Not Exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� = 2) " & _
                IIF(mlngPrintType = 1, "Or B.ִ����ֹʱ�� Is Not Null And Not exists(Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ���  in (1,2))  or b.У��ʱ�� is not null and not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in(1,2,3)) ", "") & ")" & vbNewLine & _
                "       Group By A.ҽ��id)" & vbNewLine & _
                "Order By ҳ��"
        End If
    
    Else
        If blnOnlyCheckExists Then
            strSQL = "Select 1 From ����ҽ����ӡ A, ����ҽ����¼ B Where A.ҽ��id = B.ID And A.��Ч = 1 And A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And a.��ӡʱ�� is not null " & vbNewLine & _
                IIF(mlngPrintType = 1, " and b.У��ʱ�� is not null and not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in(1,2,3))", " and 1=0") & " And Rownum<2"
        Else
            strSQL = _
                "Select Distinct ҳ��" & vbNewLine & _
                "From (Select A.ҽ��id, Max(A.ҳ��) As ҳ��" & vbNewLine & _
                "       From ����ҽ����ӡ A, ����ҽ����¼ B" & vbNewLine & _
                "       Where A.ҽ��id = B.ID And A.��Ч = 1 And A.����id = [1] And A.��ҳid = [2] And Nvl(A.Ӥ��, 0) = [3] And a.��ӡʱ�� is not null " & vbNewLine & _
                IIF(mlngPrintType = 1, " and b.У��ʱ�� is not null and not exists (Select 1 From ����ҽ����ӡ S Where S.ҽ��id = A.ҽ��id And S.��ӡ��� in(1,2,3)) ", " and 1=0") & vbNewLine & _
                "       Group By A.ҽ��id)" & vbNewLine & _
                "Order By ҳ��"
        End If
        
        If mlngPrintType = 0 Then
            strSQL = "Select 1 as ҳ�� From dual where 0=1"
        End If
        
    End If
    On Error GoTo errH
    
    Set GetStopedAdvice = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex)

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdvicePrint(ByVal intMode As Integer)
'���ܣ�ִ��ҽ������ӡ����Ԥ��
'������intMode=1-Ԥ��,2-��ӡ
    Dim lngBegin As Long, lngEnd As Long
    Dim lng�к� As Long, strReport As String
    Dim colSegment As Collection
    Dim col�����ӡ As Collection
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim intIndex As Long
    Dim varArr As Variant
    Dim int��ҩ���� As Integer
    
    'ȷ������ı�����
    strReport = IIF(optReport(optҽ��_����).value, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
    If optReport(optҽ��_����).value Then
        int��ҩ���� = IIF(mint��ҩ���г��� = 0, 0, 1)
    Else
        int��ҩ���� = IIF(mint��ҩ�������� = 0, 0, 1)
    End If
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "�����ӡ" Then 'ҽ������
        'ֻ���ڴ�ӡ����ҽ��������ܽ�����ѡ��δ���ֻ������ѡ��
        '����ѡ������Զ���ҳ�ŷֶ�
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
        
        For i = 1 To mintPageCount
            If imgChk(i).Visible Then
                If lngBegin = 0 Then
                    lngBegin = i: lngEnd = i
                ElseIf i = lngEnd + 1 Then
                    lngEnd = i
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = i: lngEnd = i
                End If
            End If
        Next
        
        If lngBegin <> 0 And lngEnd <> 0 Then
            colSegment.Add lngBegin & "-" & lngEnd
        End If
        
        If colSegment.Count = 0 Then
            MsgBox "��ѡ����Ҫ��ӡ��ҽ����ҳ�ŷ�Χ��", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "��һ��ֻѡ��һ����������һ��ҳ�ŷ�Χ����Ԥ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ҳ�������ܴ��ڻ�ҳ��ӡ�������������һ�δ�ӡ��
        If mstrTurnPages <> "" Then
            Set col�����ӡ = New Collection
            For i = 1 To colSegment.Count '�ֶε��ô�ӡ
            
                lngBegin = Split(colSegment(i), "-")(0)
                lngEnd = Split(colSegment(i), "-")(1)
                
                varArr = Split(mstrTurnPages, ",")
                For j = 0 To UBound(varArr)
                    If lngBegin <= Val(varArr(j)) And Val(varArr(j)) <= lngEnd Then
                        col�����ӡ.Add lngBegin & "-" & Val(varArr(j))
                        lngBegin = Val(varArr(j)) + 1
                    End If
                Next
                
                If lngBegin <= lngEnd Then col�����ӡ.Add lngBegin & "-" & lngEnd
            Next
            Set colSegment = col�����ӡ
        End If
        
        For i = 1 To colSegment.Count '�ֶε��ô�ӡ
        
            mintPrintCount = 0 '���ڷ�ֹԤ��ʱ����ظ���ӡ
            
            lng�к� = 0
            lngBegin = Split(colSegment(i), "-")(0)
            lngEnd = Split(colSegment(i), "-")(1)
            
            '������ֻ�ᴦ��һ��
            If mintMid = lngBegin Then
                If mbln����ҳ Then '����ҳ�����ΰ���������������к�
                    strSQL = "select max(�к�)+1 as �к� from ����ҽ����ӡ where ��ӡʱ�� is not null and ����id=[1] and ��ҳid=[2] and nvl(Ӥ��,0)=[3] and ��Ч=[4] and ҳ��=[5]"
                    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex, IIF(optReport(0).value, 0, 1), lngBegin)
                    If Not rsTmp.EOF Then
                        lng�к� = Val(rsTmp!�к� & "")
                        If chkSeqPage.value = 1 And chkSeqPage.Visible Then lng�к� = 0
                    End If
                End If
            End If
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "Ӥ��=" & cboBaby.ListIndex, "��ӡģʽ=" & mlngPrintType, "ͣ����ӡ=0", "��ʼ�к�=" & lng�к�, _
                "StartPageNum=" & lngBegin, "��ʼҳ��=" & lngBegin, "����ҳ��=" & lngEnd, "��ҩ����=" & int��ҩ����, "PressWorkFirst=" & IIF(lng�к� <> 0, 1, 0), intMode)
        Next
    ElseIf tbsMain.SelectedItem.Key = "ҽ���״�" Then
        '����ѡ������Զ���ҳ�ŷֶ�
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
    
        For i = 1 To mintStopPageCount
            intIndex = 1000 + i
            If imgChk(intIndex).Visible Then
                If lngBegin = 0 Then
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                ElseIf Val(lblNum(intIndex).Caption) = lngEnd + 1 Then
                    lngEnd = Val(lblNum(intIndex).Caption)
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                End If
            End If
        Next
        
        If lngBegin <> 0 Then colSegment.Add lngBegin & "-" & lngEnd

        If colSegment.Count = 0 Then
            MsgBox "��ѡ����Ҫ�״��ҽ����ҳ�ŷ�Χ��", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "��һ��ֻѡ��һ����������һ��ҳ�ŷ�Χ����Ԥ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        For i = 1 To colSegment.Count '��ҳ�Ŷε����״�
            mintPrintCount = 0 '���ڷ�ֹԤ��ʱ����ظ���ӡ
            lngBegin = Split(colSegment(i), "-")(0): lngEnd = Split(colSegment(i), "-")(1)
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID, "Ӥ��=" & cboBaby.ListIndex, "��ӡģʽ=" & mlngPrintType, "ͣ����ӡ=1", "��ʼ�к�=1", _
                "StartPageNum=" & lngBegin, "��ʼҳ��=" & lngBegin, "����ҳ��=" & lngEnd, "��ҩ����=" & int��ҩ����, "PressWork=1", intMode)
        Next
    End If
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
'���ܣ���ʼ��ӡ�¼�����ʼ��ҽ����ӡ��Ϣ��¼��
    
    If tbsMain.SelectedItem.Key = "�����ӡ" Then
        'Ԥ��ʱ����ظ���ӡ���
        If mintPrintCount > 0 Then
            MsgBox "�Ѿ���ӡ���ˣ�Ҫ�����´�ӡ����ʹ���ش��ܡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        'Ԥ��ʱ���ݴ�ӡ���
        If TotalPages < 0 Then
            MsgBox "Ϊ��֤��Ч����������ѡ���ȫ��ҳ����д�ӡ��", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "ҽ��ID", adBigInt
        mrsPrint.Fields.Append "ҳ��", adBigInt
        mrsPrint.Fields.Append "�к�", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    ElseIf tbsMain.SelectedItem.Key = "ҽ���״�" Then
        'Ԥ��ʱ����ظ���ӡ���
        If mintPrintCount > 0 Then
            MsgBox "�Ѿ���ӡ���ˣ������ظ���ӡ��", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        'Ԥ��ʱ���ݴ�ӡ���
        If TotalPages < 0 Then
            MsgBox "Ϊ��֤��Ч�����״���ѡ���ȫ��ҳ����д�ӡ��", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "ҽ��ID", adBigInt
        mrsPrint.Fields.Append "ҳ��", adBigInt
        mrsPrint.Fields.Append "�к�", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'���ܣ�������ӡ�¼���д�벡��ҽ����ӡ����
    Dim curDate As Date, strSQL As String
    
    If tbsMain.SelectedItem.Key = "�����ӡ" Then
        '����ҽ����ӡλ�ü�¼
        curDate = zldatabase.Currentdate
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSQL = "zl_����ҽ����ӡ_Update(" & ZVal(mrsPrint!ҽ��ID) & "," & mrsPrint!ҳ�� & "," & mrsPrint!�к� & "," & _
                    mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(0).value, 0, 1) & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & "'," & mlngPrintType & ")"
                Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                mrsPrint.MoveNext
                
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    ElseIf tbsMain.SelectedItem.Key = "ҽ���״�" Then
        '���ҽ��ͣ��ʱ�����״��־
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSQL = "Zl_����ҽ����ӡ_Update(" & mrsPrint!ҽ��ID & "," & mrsPrint!ҳ�� & "," & mrsPrint!�к� & ",null,null,null,null,null,null," & mlngPrintType & ",1)"
                Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                mrsPrint.MoveNext
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
 
    Set mrsPrint = Nothing
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'���ܣ��������ݴ�ӡ�¼�����¼ҽ����ӡ������
'˵�����������������Ҫ��ӡʱ���ǲ��ἤ����¼���
    If tbsMain.SelectedItem.Key = "�����ӡ" Then
        If Page >= 1 And Row >= 1 Then
            'mrsPrint.Filter = "ҽ��ID=" & ID 'NULL�᷵��Ϊ0
            mrsPrint.Filter = "ҽ��ID=" & ID & " and ҳ�� =" & Page & " and �к�=" & Row
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!ҽ��ID = ID
                mrsPrint!ҳ�� = Page
                mrsPrint!�к� = Row
            End If
            mrsPrint.Update
        End If
    ElseIf tbsMain.SelectedItem.Key = "ҽ���״�" Then
        If ID > 0 And Page >= 1 And Row >= 1 Then
            mrsPrint.Filter = "ҽ��ID=" & ID
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!ҽ��ID = ID
                mrsPrint!ҳ�� = Page
                mrsPrint!�к� = Row
                mrsPrint.Update
            End If
        End If
    End If
End Sub

Private Function GetReportRows(ByVal lngSys As Long, ByVal strReport As String, Optional ByVal intFormat As Integer = 1) As Long
'���ܣ���ȡָ����������Ҫ������Ŀɴ�ӡ��������
'������lngSys=ϵͳ��ţ�Ϊ0��ʾ������
'      strReport=������
'      intFormat=�����ʽ��,ȱʡΪ1
'���أ�0��ʾû��������
'˵����
'  1.��������д��ڶ����������������һ����Ϊ��Ҫ���
'  2.�������������ɴ�ӡ������ָ����֮�����������
    Dim rsTable As ADODB.Recordset
    Dim rsColumn As ADODB.Recordset
    Dim strSQL As String, i As Long, j
    Dim blnHead As Boolean, blnBody As Boolean
    Dim lngBodyH As Long, lngHeadH As Long
    
    On Error GoTo errH
    
    strSQL = "Select A.ID as ����ID,B.ID,B.W,B.H,B.�и�,B.����" & _
        " From zlReports A,zlRPTItems B" & _
        " Where A.ID=B.����ID And B.����=4 And Nvl(A.ϵͳ,0)=[1] And A.���=[2] And B.��ʽ��=[3]" & _
        " Order by B.W*B.H Desc"
    Set rsTable = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngSys, strReport, intFormat)
    If rsTable.EOF Then Exit Function
    
    strSQL = "Select ���,��ͷ,���� From zlRPTItems Where ����ID=[1] And ��ʽ��=[2] And �ϼ�ID=[3] And ����=6 Order by ���"
    Set rsColumn = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTable!����ID), intFormat, Val(rsTable!ID))
    If rsColumn.EOF Then Exit Function
    
    '���´�������Զ��屨���еķ�����д
    '----------------------------------
    '�����ͷ�߶�:�Ե�һ��Ϊ׼
    For i = 0 To UBound(Split(rsColumn!��ͷ, "|"))
        lngHeadH = lngHeadH + Val(Split(Split(rsColumn!��ͷ, "|")(i), "^")(1))
    Next
    
    '�������߶�
    blnHead = False: blnBody = False
    rsColumn.MoveFirst
    Do While Not rsColumn.EOF
        i = UBound(Split(rsColumn!��ͷ, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(rsColumn!��ͷ, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or Not IsNull(rsColumn!����)
        rsColumn.MoveNext
    Loop
    If Not blnHead And blnBody Then '���б���
        lngBodyH = rsTable!H
    Else
        If rsTable!H - lngHeadH + 15 < 0 Then
            lngBodyH = 0
        Else
            lngBodyH = rsTable!H - lngHeadH + 15
        End If
    End If
    
    '�������
    GetReportRows = Int(lngBodyH / rsTable!�и�) * NVL(rsTable!����, 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Insert��ӡ��¼()
'���ܣ����ɽ�Ҫ��ӡ��ҽ����¼��Ҫ����ͣ����ӡ��ҽ��������/����������ҽ����Ӥ���Ƿֿ��ģ�����������
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    '���˺�Ӥ����Ҫ�жϣ�����ҽ��������ʱҪ�����������������ʱҽ����������ת��ҳ�ʹ�ӡ������ǵ����
    '�ж��Ƿ�Ҫ���ɳ����ӡ�ļ�¼������ѭ����j ��ʾ��Ч��i ��ʾӤ�����i=0ʱ��ʾ����
    For j = 0 To 1
        lngRows = IIF(j = 0, mlngRows����, mlngRows����)
        For i = 0 To cboBaby.ListCount - 1
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & i & "," & j & ",null,null,3)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ����ӡ_Insert(" & mlng����ID & "," & mlng��ҳID & "," & i & "," & j & "," & lngRows & ")"
        Next
    Next
    
    '�ύ����
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub LoadAllPaper()
'���ܣ���������������ͼƬֽ��
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer, i As Integer
    
    On Error GoTo errH
    
    For i = 1 To 4
        Load picContainer(i)
        picContainer(i).Width = 6160
        picContainer(i).Height = 3400
    Next

    Set picContainer(1).Container = Me
    Set picContainer(3).Container = Me

    Set picContainer(2).Container = picContainer(1)
    Set picContainer(4).Container = picContainer(3)
    
    picContainer(1).Top = 1720
    picContainer(1).Left = 350
    picContainer(3).Top = 1720
    picContainer(3).Left = 350
    
    picContainer(2).Top = 0
    picContainer(4).Top = 0
    picContainer(2).Left = 0
    picContainer(4).Left = 0
    
    For i = 1 To 4
        picContainer(i).Visible = True
        picContainer(i).ZOrder 0
    Next
    
    vsc.Left = 6520
    vsc.Height = 3300
    vsc.Top = 1820
    vsc.Width = 200
    vsc.ZOrder 0

    strSQL = "select max(a.ҳ��) as ҳ�� from ����ҽ����ӡ a where a.����id=[1] and a.��ҳid=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
 
    intTmp = Val(rsTmp!ҳ�� & "") * 2
    picPaper(0).Tag = intTmp
    picPaperB(0).Tag = intTmp
    
    For i = 1 To intTmp
        Call LoadPaper(0, i)
        Call LoadPaper(1, i)
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPaper(ByVal intCt As Integer, ByVal intNum As Integer)
'���ܣ�����ͼֽ�ţ�Ŀǰ֧�����ҳ�� 999ҳ
'������intCt������0��������ӡfraPrint(0)��2��ͣ����ӡfraPrint(1)��intNum ҳ��
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim intRow As Integer
    Dim intIndex As Integer
    
    On Error GoTo errH
    
    intIndex = intNum + 1000 * intCt
    
    intRow = 1 + (intNum - 1) \ 7
    
    lngLeft = ((intNum - 1) Mod 7) * (picPaper(0).Width + 200)
    lngTop = (intRow - 1) * (picPaper(0).Height + 250)
    
    '��ͼƬ
    Load picPaper(intIndex)
    Load picPaperB(intIndex)
 
    '����ͼ������ͼƬ
    Set picPaperB(intIndex).Container = picContainer(2 + intCt * 2)
    Set picPaper(intIndex).Container = picContainer(2 + intCt * 2)
 
    picPaper(intIndex).Left = lngLeft
    picPaper(intIndex).Top = lngTop
    picPaper(intIndex).Width = picPaper(0).Width
    picPaper(intIndex).Height = picPaper(0).Height
    picPaper(intIndex).BackColor = picPaper(0).BackColor
    picPaper(intIndex).Visible = False
    picPaper(intIndex).ZOrder 0

    picPaperB(intIndex).Left = picPaper(intIndex).Left + 50
    picPaperB(intIndex).Top = picPaper(intIndex).Top + 50
    picPaperB(intIndex).Width = picPaper(0).Width
    picPaperB(intIndex).Height = picPaper(0).Height
    picPaperB(intIndex).BackColor = picPaperB(0).BackColor
    picPaperB(intIndex).Visible = False
    
    'ֽ�ϵ�ͼ��
    Load imgIco(intIndex)
    Set imgIco(intIndex).Container = picPaper(intIndex)
    Set imgIco(intIndex).Picture = imgIcon(0).Picture
    imgIco(intIndex).Left = (picPaper(intIndex).Width - imgIco(intIndex).Width) / 2
    imgIco(intIndex).Top = 260
    imgIco(intIndex).Visible = True
    imgIco(intIndex).ZOrder 1
    
    Load lblNum(intIndex)
    Set lblNum(intIndex).Container = picPaper(intIndex)
    lblNum(intIndex).Visible = True
    lblNum(intIndex).Caption = intNum
    lblNum(intIndex).ToolTipText = "��" & intNum & "ҳ"
    lblNum(intIndex).FontSize = lblNum(0).FontSize
    lblNum(intIndex).Left = (picPaper(intIndex).Width - lblNum(intIndex).Width) / 2
    lblNum(intIndex).Top = imgIco(intIndex).Height + imgIco(intIndex).Top + 10
    lblNum(intIndex).BackColor = picPaper(0).BackColor
    
    '��ѡͼƬ�������пؼ��ɼ���
    Load imgChk(intIndex)
    Set imgChk(intIndex).Container = picPaper(intIndex)
    Set imgChk(intIndex).Picture = imgChk(0).Picture '�̶�
    imgChk(intIndex).Width = 240
    imgChk(intIndex).Height = 240
    imgChk(intIndex).Left = picPaper(0).Width - imgChk(intIndex).Width
    imgChk(intIndex).Top = -10
    imgChk(intIndex).Visible = False
    imgChk(intIndex).ZOrder 1
    
    Exit Sub
errH:
    If 1 = 2 Then
        Resume
    End If
    err.Clear
End Sub
   
Private Sub UnLoadPaper()
    Dim i As Integer
    
    On Error Resume Next
    
    '��ж�������ڵ��Ƽ���ж������
    For i = 1 To Val(picPaper(0).Tag)
        Unload imgChk(i)
        Unload imgIco(i)
        Unload lblNum(i)
        Unload picPaperB(i)
        Unload picPaper(i)
    Next
    
    For i = 1 To Val(picPaperB(0).Tag)
        Unload imgChk(i + 1000)
        Unload imgIco(i + 1000)
        Unload lblNum(i + 1000)
        Unload picPaperB(i + 1000)
        Unload picPaper(i + 1000)
    Next
    
    For i = 1 To 4
        Unload picContainer(i)
    Next
    
    err.Clear
End Sub

Private Sub imgIco_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub lblNum_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub imgChk_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub picPaper_Click(Index As Integer)
    Dim blnTmp As Boolean
    Dim i As Integer
    
    blnTmp = imgChk(Index).Visible
    imgChk(Index).Visible = Not blnTmp
    
    If Not (Index > 1000 Or mintMid = 0 Or mintMid > Index) Then
        If blnTmp Then
            For i = Index + 1 To mintPageCount
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        Else
            For i = mintMid To Index - 1
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        End If
    End If
    
    If mbln����ҳ And Index < 1000 Then
        If mintMid = 1 Then
            chkSeqPage.Visible = imgChk(mintMid).Visible
        Else
            chkSeqPage.Visible = imgChk(mintMid).Visible And Not imgChk(mintMid - 1).Visible
        End If
    End If
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    If tbsMain.SelectedItem.Key = "�����ӡ" Then
        picContainer(pic����_ֽ��).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    Else
        picContainer(pic�״�_ֽ��).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub cmdClsLastPrint_Click()
'���ܣ������ӡ��¼
    Call ClearPrintRs(True)
End Sub

Private Sub txtClearPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtClearPage_Validate(False)
        cmdClear.SetFocus
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtClearPage_Validate(Cancel As Boolean)
    Dim lngTmp As Long
    lngTmp = Val(txtClearPage.Text)
    If lngTmp = 0 Then
        txtClearPage.Text = 1
    ElseIf lngTmp > mlngPrintedMaxPage Then
        txtClearPage.Text = mlngPrintedMaxPage
    Else
        txtClearPage.Text = lngTmp
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearPrintRs(False)
End Sub

Private Sub ClearPrintRs(ByVal bln�ϴδ�ӡ As Boolean)
'���ܣ���ĳҳ��ʼ�����ӡ
'������bln�ϴδ�ӡ ��true ����ϴδ�ӡ��false ���ָ��ҳ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim strλ�� As String, str��ӡ�� As String
    Dim lngTmp As Long, strTmp As String, str��ӡʱ�� As String
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Long
    Dim lngҳ�� As Long
    
    If MsgBox("ȷʵҪ����Ѵ�ӡ��ҽ����¼���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    arrSQL = Array()
    
    On Error GoTo errH
    
    lngRows = IIF(optReport(optҽ��_����).value, mlngRows����, mlngRows����)
    
    If bln�ϴδ�ӡ Then
        If optReport(optҽ��_����).value Then
            '��������ʱ���ж�
            strSQL = "select max(��ӡʱ��) as ʱ�� from ����ҽ����ӡ Where ����id=[1] And ��ҳid=[2] And Nvl(Ӥ��,0)=[3] And ��Ч=0"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex)
            If Not IsNull(rsTmp!ʱ��) Then
                If mdat����ʱ�� > rsTmp!ʱ�� Then
                    MsgBox "�ϴδ�ӡ������֮ǰ��Ҫ�Ȼ����������������ӡ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        strSQL = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(optҽ��_����).value, 0, 1) & ",null,null,1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
         
        '���ɼ�¼
        strSQL = "Zl_����ҽ����ӡ_Insert(" & mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(optҽ��_����).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    Else
        If mdat����ʱ�� <> CDate("1900-01-01") And optReport(optҽ��_����).value Then
            strSQL = "select max(ҳ��) as ҳ�� from ����ҽ����ӡ a " & _
                " where a.����id=[1] and a.��ҳid=[2] and nvl(a.Ӥ��,0)=[3] and a.��Ч=[4] and ��ӡʱ��<[5]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex, 0, mdat����ʱ��)
            lngҳ�� = Val(rsTmp!ҳ�� & "")
        End If
    
        If Val(txtClearPage.Text) <= lngҳ�� Then
            If MsgBox("�����ӡ��ҽ�����а���������ǰ��������ݣ���Ҫ������Ȼ�������������ѡ �� ���������һ������֮���ӡ�����ݣ�ѡ �� ��������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        strSQL = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(optҽ��_����).value, 0, 1) & "," & Val(txtClearPage.Text) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        '���ɼ�¼
        strSQL = "Zl_����ҽ����ӡ_Insert(" & mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(optҽ��_����).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        If Val(txtClearPage.Text) > 1 Then
            strSQL = "Select a.��ӡʱ��, a.��ӡ��" & vbNewLine & _
                " From ����ҽ����ӡ A,����ҽ����¼ b Where a.ҽ��id=b.id and b.������� in ('5','6') and a.����id =[1] And a.��ҳid =[2]" & _
                " and a.Ӥ��=[3] And a.��Ч =[4] And a.ҳ�� =[5] and a.�к�=[6]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex, IIF(optReport(optҽ��_����).value, 0, 1), _
                Val(txtClearPage.Text) - 1, lngRows)
            If Not rsTmp.EOF Then
                str��ӡ�� = rsTmp!��ӡ�� & ""
                strTmp = Format(rsTmp!��ӡʱ��, "yyyy-MM-dd HH:mm:ss")
                strTmp = "To_Date('" & strTmp & "','YYYY-MM-DD HH24:MI:SS')"
                str��ӡʱ�� = strTmp
            End If
        End If
    End If
    
    '�ύ����
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    If str��ӡ�� <> "" Then
        strSQL = "Select a.ҽ��id,a.ҳ��,a.�к� From ����ҽ����ӡ A Where a.��ӡʱ�� is null and a.����id=[1] And a.��ҳid=[2] and a.Ӥ��=[3] And a.��Ч=[4] And a.ҳ��=[5]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, cboBaby.ListIndex, IIF(optReport(optҽ��_����).value, 0, 1), Val(txtClearPage.Text) - 1)
        If Not rsTmp.EOF Then
            arrSQL = Array()
            For i = 1 To rsTmp.RecordCount
                strSQL = "zl_����ҽ����ӡ_Update(" & ZVal(rsTmp!ҽ��ID) & "," & rsTmp!ҳ�� & "," & rsTmp!�к� & "," & _
                    mlng����ID & "," & mlng��ҳID & "," & cboBaby.ListIndex & "," & IIF(optReport(optҽ��_����).value, 0, 1) & "," & _
                    str��ӡʱ�� & ",'" & str��ӡ�� & "'," & mlngPrintType & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                rsTmp.MoveNext
            Next
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
    
    Call RefreshFace
    
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
