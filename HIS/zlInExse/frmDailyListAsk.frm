VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailyListAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѯ��������"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboPage 
      Height          =   300
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2745
      Width           =   1410
   End
   Begin VB.Frame fraUnit 
      Caption         =   "���˲�����"
      Height          =   1215
      Left            =   3360
      TabIndex        =   16
      Top             =   1440
      Width           =   1920
      Begin VB.OptionButton optUnit 
         Caption         =   "���˵�ǰ����"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1500
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "�з��õĲ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.Frame fraTime 
      Caption         =   "��ѯ���ð�"
      Height          =   1215
      Left            =   3840
      TabIndex        =   13
      Top             =   105
      Width           =   1440
      Begin VB.OptionButton opttime 
         Caption         =   "�Ǽ�ʱ��"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   380
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opttime 
         Caption         =   "����ʱ��"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   800
         Width           =   1020
      End
   End
   Begin VB.Frame fraState 
      Caption         =   "����״̬"
      Height          =   1215
      Left            =   1680
      TabIndex        =   12
      Top             =   1440
      Width           =   1440
      Begin VB.CheckBox chkInOut 
         Caption         =   "��Ժ����"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox chkInOut 
         Caption         =   "��Ժ����"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "��������"
      Height          =   1230
      Left            =   135
      TabIndex        =   11
      Top             =   1425
      Width           =   1440
      Begin VB.CheckBox chkPatiType 
         Caption         =   "��ҽ������"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "ҽ������"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1155
      End
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   "����ʱ�䷶Χ"
      Height          =   1215
      Left            =   135
      TabIndex        =   8
      Top             =   105
      Width           =   3600
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   900
         TabIndex        =   1
         Top             =   750
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   84017155
         CurrentDate     =   36257.9583333333
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   84017155
         CurrentDate     =   36257.9597337963
      End
      Begin VB.Label lblEnd 
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblBegin 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5385
      TabIndex        =   7
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5385
      TabIndex        =   6
      Top             =   330
      Width           =   1100
   End
   Begin VB.Label lblPage 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   2805
      Width           =   720
   End
End
Attribute VB_Name = "frmDailyListAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mbytInFun As Byte    '0-һ���嵥�е���,1-���˷��ò�ѯ�е���
Public mdatBegin As Date
Public mdatEnd As Date
Public mlngPageID As Long
Public mlng����ID As Long

Public mblnAskOk As Boolean
Public mstrPrivs As String
Public mlngModul As Long
Public mblnDateMoved As Boolean '��ǰ��ѡ�����������Ƿ��ں����ݱ���

Private Sub chkInOut_Click(Index As Integer)
    If chkInOut(Index).Value = 0 Then
        If chkInOut((Index + 1) Mod 2).Value = 0 Then
            chkInOut(Index).Value = 1
        End If
    End If
End Sub

Private Sub chkPatiType_Click(Index As Integer)
    If chkPatiType(Index).Value = 0 Then
        If chkPatiType((Index + 1) Mod 2).Value = 0 Then
            chkPatiType(Index).Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnAskOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim lngTmp As Long
    Dim blnHavePara As Boolean
    
        
    If dtpBegin.Value >= dtpEnd.Value Then
        MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
        Exit Sub
    End If
    blnHavePara = InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʼʱ��", Format(Me.dtpBegin.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara
    zlDatabase.SetPara "����ʱ��", Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, mlngModul, blnHavePara

    lngTmp = DateDiff("d", Me.dtpEnd.Value, zlDatabase.Currentdate)
    zlDatabase.SetPara "�������", lngTmp, glngSys, mlngModul, blnHavePara
    lngTmp = DateDiff("d", Me.dtpBegin.Value, Me.dtpEnd.Value)
    zlDatabase.SetPara "��ʼ���", lngTmp, glngSys, mlngModul, blnHavePara
    
    
    If mbytInFun = 0 Then
        zlDatabase.SetPara "��ҽ������", chkPatiType(0).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "ҽ������", chkPatiType(1).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "��Ժ����", chkInOut(0).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "��Ժ����", chkInOut(1).Value, glngSys, mlngModul, blnHavePara
        zlDatabase.SetPara "���˲���ģʽ", IIf(optUnit(0).Value = True, 0, 1), glngSys, mlngModul, blnHavePara
                
        '�����ڼ�
        zlDatabase.SetPara "����ʱ��", IIf(opttime(1).Value, 1, 0), glngSys, mlngModul, blnHavePara
        
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), "yyyy-MM-dd HH:mm:ss"), , , Me.Caption)
    End If
    
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    If cboPage.Visible Then
        mlngPageID = Val(cboPage.ItemData(cboPage.ListIndex))
    End If
    mblnAskOk = True
    Me.Hide
End Sub

Private Sub dtpBegin_Change()
    If Me.dtpBegin.Value > Me.dtpEnd.Value Then
        Me.dtpEnd.Value = Me.dtpBegin.Value
    End If
End Sub

Private Sub dtpEnd_Change()
    If Me.dtpBegin.Value > Me.dtpEnd.Value Then
        Me.dtpBegin.Value = Me.dtpEnd.Value
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim lngTmp As Long
    Dim strStartTime  As String
    Dim strEndTime As String, blnParSet As Boolean
    
    On Error Resume Next
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0
    
    strEndTime = zlDatabase.GetPara("����ʱ��", glngSys, mlngModul, "23:59:59", Array(lblEnd, dtpEnd), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys, mlngModul, 0, Array(lblEnd, dtpEnd), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpEnd.Value = CDate(Format(zlDatabase.Currentdate() - lngTmp, "yyyy-MM-dd") & " " & strEndTime)
    
    strStartTime = zlDatabase.GetPara("��ʼʱ��", glngSys, mlngModul, "00:00:00", Array(lblBegin, dtpBegin), blnParSet)
    lngTmp = Val(zlDatabase.GetPara("��ʼ���", glngSys, mlngModul, 0, Array(lblBegin, dtpBegin), blnParSet))
    If lngTmp > 7 Then lngTmp = 7
    Me.dtpBegin.Value = CDate(Format(Me.dtpEnd.Value - lngTmp, "yyyy-MM-dd") & " " & strStartTime)
    
    If mbytInFun = 0 Then
        '�����ڼ�
        lngTmp = IIf(zlDatabase.GetPara("����ʱ��", glngSys, mlngModul, 0, Array(opttime(0), opttime(1)), blnParSet) = "1", 1, 0)
        opttime(lngTmp).Value = True
        
        chkPatiType(0).Value = IIf(zlDatabase.GetPara("��ҽ������", glngSys, mlngModul, "1", Array(chkPatiType(0)), blnParSet) = "1", 1, 0)
        chkPatiType(1).Value = IIf(zlDatabase.GetPara("ҽ������", glngSys, mlngModul, "1", Array(chkPatiType(1)), blnParSet) = "1", 1, 0)
        
        lngTmp = IIf(zlDatabase.GetPara("���˲���ģʽ", glngSys, mlngModul, "0", Array(optUnit(0), optUnit(1)), blnParSet) = "1", 1, 0)
        optUnit(lngTmp).Value = True
        
        If InStr(";" & mstrPrivs, ";��Ժ���˲�ѯ;") = 0 Then
            chkInOut(0).Enabled = False
            chkInOut(1).Enabled = False
            chkInOut(0).Value = 1
            chkInOut(1).Value = 0
        Else
            chkInOut(0).Enabled = True
            chkInOut(1).Enabled = True
            chkInOut(0).Value = IIf(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "1", Array(chkInOut(0)), blnParSet) = "1", 1, 0)
            chkInOut(1).Value = IIf(zlDatabase.GetPara("��Ժ����", glngSys, mlngModul, "1", Array(chkInOut(1)), blnParSet) = "1", 1, 0)
        End If
        lblPage.Visible = False
        cboPage.Visible = False
        Me.Height = 3150
    Else
        fraType.Visible = False
        fraState.Visible = False
        fraTime.Visible = False
        fraUnit.Visible = False
        cmdOk.Left = fraTime.Left
        cmdCancel.Left = fraTime.Left
        Me.Width = Me.Width - fraTime.Width
        Me.Height = Me.Height - fraType.Height - 100
        lblPage.Top = fraType.Top
        cboPage.Top = fraType.Top - 30
        Call LoadסԺ����(mlng����ID, mlngPageID)
    End If
End Sub

Private Sub LoadסԺ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim strSql As String, rsPage As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Distinct ��ҳID From ������ҳ Where ����ID = [1] And �������� = 0 Order By ��ҳID Desc"
    Set rsPage = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    cboPage.Clear
    cboPage.AddItem "����סԺ"
    cboPage.ItemData(cboPage.NewIndex) = 0
    Do While Not rsPage.EOF
        cboPage.AddItem "��" & Val(NVL(rsPage!��ҳID)) & "��סԺ"
        cboPage.ItemData(cboPage.NewIndex) = Val(NVL(rsPage!��ҳID))
        If Val(NVL(rsPage!��ҳID)) = lng��ҳID Then cboPage.ListIndex = cboPage.NewIndex
        rsPage.MoveNext
    Loop
    If cboPage.ListIndex < 0 Then cboPage.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
End Sub

