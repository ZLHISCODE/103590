VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinanceSuperviseStandbyMoneyEdit 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "���ý����õ�"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtBackTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3795
      Width           =   2145
   End
   Begin VB.TextBox txtBackPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3795
      Width           =   1785
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5295
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1035
      Width           =   2040
   End
   Begin VB.ComboBox cboPerson 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Text            =   "cboPerson"
      Top             =   1890
      Width           =   2040
   End
   Begin VB.TextBox txtMemo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2820
      Width           =   6120
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3315
      Width           =   2145
   End
   Begin VB.TextBox txtInputPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3315
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6315
      TabIndex        =   13
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5055
      TabIndex        =   12
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   350
      Left            =   90
      TabIndex        =   14
      Top             =   4800
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   2340
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   101056515
      CurrentDate     =   41520
   End
   Begin VB.TextBox txtMoney 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2355
      Width           =   2160
   End
   Begin VB.Label lblBackTime 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ�ʱ��"
      Height          =   210
      Left            =   4395
      TabIndex        =   21
      Top             =   3855
      Width           =   840
   End
   Begin VB.Label lblBackPerson 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   210
      Left            =   690
      TabIndex        =   20
      Top             =   3855
      Width           =   630
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "���ý��(&M)"
      Height          =   210
      Left            =   4125
      TabIndex        =   4
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   15
      X2              =   10455
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5070
      TabIndex        =   17
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label lblPerson 
      AutoSize        =   -1  'True
      Caption         =   "������(&P)"
      Height          =   210
      Left            =   375
      TabIndex        =   0
      Top             =   1965
      Width           =   945
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "ժ  Ҫ(&Z)"
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   2910
      Width           =   945
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��(&T)"
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ�ʱ��"
      Height          =   210
      Left            =   4395
      TabIndex        =   10
      Top             =   3375
      Width           =   840
   End
   Begin VB.Label lblInputPerson 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ���"
      Height          =   210
      Left            =   690
      TabIndex        =   8
      Top             =   3360
      Width           =   630
   End
   Begin VB.Line linMain 
      BorderColor     =   &H8000000C&
      X1              =   -285
      X2              =   10155
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Label lblTittle 
      Alignment       =   2  'Center
      Caption         =   "���ý����õ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   135
      TabIndex        =   16
      Top             =   210
      Width           =   7170
   End
End
Attribute VB_Name = "frmFinanceSuperviseStandbyMoneyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlngID As Long
Private mstr������ As String
Public Enum gEditCard
    EM_ED_���� = 0
    EM_ED_�ϸ� = 1
    EM_ED_�鿴 = 2
End Enum
Private mEditType As gEditCard
Private mrsChargePerson As ADODB.Recordset
Private mblnFirst As Boolean, mblnOK As Boolean
Private mblnChange  As Boolean
Private mblnUnload As Boolean

Public Function EditCard(ByVal frmMain As Object, _
    ByVal EditType As gEditCard, _
    ByVal lngModuel As Long, ByVal strPrivs As String, _
    ByVal str������ As String, Optional ByVal lngID As Long = 0) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '       EditType-�༭����
    '       lngID-�ݴ�ID(�鿴ʱ����)
    '       str������-������
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 15:37:32
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModuel: mstrPrivs = strPrivs
    mlngID = lngID: mblnOK = False: mstr������ = str������
    If frmMain Is Nothing Then
         Me.Show vbModal
    Else
         Me.Show vbModal, frmMain
    End If
    EditCard = mblnOK
End Function

Private Sub ClearCtrlData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2013-10-12 15:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtMemo.Text = "": txtMoney.Text = ""
    txtInputPerson.Text = ""
    txtTime.Text = ""
    txtBackPerson.Text = ""
    txtBackTime.Text = ""
End Sub
Private Sub SetCtrlEnabled()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enabled����
    '����:���˺�
    '����:2013-10-12 16:02:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    Dim lngBackColor As Long
    
    blnEnabled = mEditType = EM_ED_���� Or mEditType = EM_ED_�ϸ�
    lngBackColor = IIf(blnEnabled, &H80000005, &H8000000F)
    txtMemo.Enabled = blnEnabled: txtMemo.BackColor = lngBackColor
    txtMoney.Enabled = blnEnabled: txtMoney.BackColor = lngBackColor
    cboNO.Enabled = blnEnabled: cboNO.BackColor = lngBackColor
    dtpDate.Enabled = blnEnabled
    dtpDate.Value = zlDatabase.Currentdate
    cboPerson.Enabled = blnEnabled: cboPerson.BackColor = lngBackColor
    txtInputPerson.Enabled = False
    txtTime.Enabled = False
    txtBackPerson.Enabled = False
    txtBackTime.Enabled = False
    cmdOK.Visible = mEditType <> EM_ED_�鿴
End Sub

Private Function LoadCardData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؿ�Ƭ����
    '����:���سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 15:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblMoney As Double
    On Error GoTo errHandle
    If mEditType = EM_ED_���� Or mEditType = EM_ED_�ϸ� Then
        Call ClearCtrlData
        txtInputPerson.Text = UserInfo.����
        txtTime.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        dblMoney = Val(zlDatabase.GetPara("ȱʡ���ñ��ý��", glngSys, mlngModule, 1000, Array(txtMoney, lblMoney), InStr(1, mstrPrivs, ";��������;") > 0))
        txtMoney.Text = Format(dblMoney, "0.00")
        If txtMoney.Enabled Then txtMoney.BackColor = &H80000005
        If mEditType = EM_ED_�ϸ� Then
            lblTittle.Caption = "���ý����õ�(�ϸ�)"
        Else
            lblTittle.Caption = "���ý����õ�"
        End If
        LoadCardData = LoadPerson: Exit Function
    End If

    strSQL = "" & _
    "   Select ID,�ս�ID,��¼����,NO,���㷽ʽ,���,�տ�Ա as ������,����ʱ��, " & _
    "           �ջ���,�ջ�ʱ��,��ע,�Ǽ���,�Ǽ�ʱ��  " & _
    "   From ��Ա�ݴ��¼ " & _
    "   Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ�ָ���ı��ý����ü�¼,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    cmdCancel.Caption = "�˳�(&X)"
    With cboPerson
        .Clear
        .AddItem Nvl(rsTemp!������)
        .ListIndex = .NewIndex
    End With
    dtpDate.Value = Format(rsTemp!����ʱ��, "yyyy-mm-dd")
    txtTime.Text = Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    txtBackTime.Text = Format(rsTemp!�ջ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    txtInputPerson.Text = Nvl(rsTemp!�Ǽ���)
    txtBackPerson.Text = Nvl(rsTemp!�ջ���)
    txtMoney.Text = Format(Val(Nvl(rsTemp!���)), "###0.00;-###0.00;0.00;-0.00")
    txtMoney.Tag = Nvl(rsTemp!���㷽ʽ)
    txtMemo.Text = Nvl(rsTemp!��ע)
    cboNO.AddItem Nvl(rsTemp!NO)
    cboNO.ListIndex = cboNO.NewIndex
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Sub cboPerson_Click()
    mblnChange = True
End Sub
  

Private Sub cboPerson_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    If KeyAscii <> 13 Then Exit Sub
    
    If cboPerson.Locked Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strText = UCase(cboPerson.Text)
    If cboPerson.ListIndex <> -1 Then
        '�����б�ʱ,�����ı�������������
        If strText <> UCase(cboPerson.List(cboPerson.ListIndex)) Then Call zlcontrol.CboSetIndex(cboPerson.hWnd, -1)
    End If
    If strText = "" Then cboPerson.ListIndex = -1: Exit Sub
    If cboPerson.ListIndex >= 0 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    intIdx = -1
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsChargePerson)
    strCompents = Replace(gstrLike, "%", "*") & strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0 '0-�������ȫ����
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1 '1-�������ȫ��ĸ
    Else
        intInputType = 2 '2-����
    End If
    mrsChargePerson.Filter = 0: iCount = 0
    With mrsChargePerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsChargePerson.EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!���)) = Val(strText) Then
                    If iCount = 0 Then strResult = Nvl(!����)
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Val(mrsChargePerson!���) Like strText & "*" Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                 End If
                 
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strText Then
                    If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                    If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                    If CheckPersonExists(Nvl(!����)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                End If
            End Select
            mrsChargePerson.MoveNext
        Loop
    End With
    
    If iCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
    'ֱ�Ӷ�λ
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckPersonExists(strResult, True) Then zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
     If rsTemp.RecordCount = 0 Then
        'δ�ҵ�
        rsTemp.Close: Set rsTemp = Nothing
        KeyAscii = 0: zlcontrol.TxtSelAll cboPerson: Exit Sub
     End If
     
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = "���"
    Case 1 '����ȫƴ��
        rsTemp.Sort = "����"
    Case Else
        '����ѡ������
        rsTemp.Sort = "���"
    End Select
    '����ѡ����
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, cboPerson, rsTemp, True, "", "", rsReturn) Then
        If cboPerson.Enabled Then cboPerson.SetFocus
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '���ж�λ
                If CheckPersonExists(Nvl(rsReturn!����), True) Then
                    'zlCommFun.PressKey vbKeyTab
                End If
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
End Sub

Private Sub cboPerson_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
    If cboPerson.Text <> "" Then
        If cbo.FindIndex(cboPerson, zlStr.NeedName(cboPerson.Text), True) = -1 Then cboPerson.ListIndex = -1: cboPerson.Text = ""
    End If
    If cboPerson.Text = "" Then Call cboPerson_KeyPress(vbKeyReturn)
    '�����ݣ���������
    If cboPerson.ListIndex = -1 And cboPerson.ListCount <> 0 Then Cancel = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim strNO As String
    If isValied = False Then Exit Sub
    If SaveData(strNO) = False Then Exit Sub
    Call BillPrint(strNO)
    MsgBox "���ý𷢷ųɹ�!", vbOKOnly + vbInformation, gstrSysName
    Call LoadCardData
    If cboPerson.Enabled And cboPerson.Visible Then cboPerson.SetFocus
    mblnChange = False: mblnOK = True
End Sub

Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�տ��վݴ�ӡ
    '����:���˺�
    '����:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "���ý����õ�") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("���ý����õ���ӡ��ʽ", glngSys, mlngModule))     'ʹ��ҽ��վ����ز���
    Case 0    '����ӡ
        Exit Sub
    Case 1    '��������ӡ
        blnPrint = True
    Case 2    'ѡ���ӡ
        If MsgBox("���Ƿ�Ҫ��ӡ�ɿ��վݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500_1", Me, "NO=" & strNO, 2)
End Sub

Private Function SaveData(ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:strNo-���ݱ���ɹ������ص��ݺ�
    '����:���ݱ���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 17:07:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strSQL As String
    On Error GoTo errHandle
    
    lngID = zlDatabase.GetNextId("��Ա�ݴ��¼")
    strNO = zlDatabase.GetNextNo(141)
    '    Zl_��Ա�ݴ��¼_Insert
    strSQL = "Zl_��Ա�ݴ��¼_Insert("
    '  Id_In       In ��Ա�ݴ��¼.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In       In ��Ա�ݴ��¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  ���_In     In ��Ա�ݴ��¼.���%Type,
    strSQL = strSQL & "" & Val(txtMoney.Text) & ","
    '  ������_In   In ��Ա�ݴ��¼.�տ�Ա%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cboPerson.Text) & "',"
    '  ����ʱ��_In In ��Ա�ݴ��¼.����ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(dtpDate.Value, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ��ע_In     In ��Ա�ݴ��¼.��ע%Type,
    strSQL = strSQL & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & Trim(txtMemo.Text) & "'") & ","
    '  �Ǽ���_In   In ��Ա�ݴ��¼.�Ǽ���%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In In ��Ա�ݴ��¼.�Ǽ�ʱ��%Type
    strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ��¼����_In In ��Ա�ݴ��¼.��¼����%Type
    If mEditType = EM_ED_�ϸ� Then
        strSQL = strSQL & "" & 1 & ")"
    ElseIf mEditType = EM_ED_���� Then
        strSQL = strSQL & "" & 11 & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function isValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-12 16:56:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If zlCommFun.ActualLen(txtMemo.Text) > 50 Then
         MsgBox "ժҪ����,���ֻ������25���ַ���50������", vbInformation, gstrSysName
         If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
         Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        MsgBox "ժҪ�в��ܰ���������!", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
    If cboPerson.ListIndex < 0 Then
        MsgBox "δѡ��������!", vbInformation, gstrSysName
        If cboPerson.Visible And cboPerson.Enabled Then cboPerson.SetFocus
        Exit Function
    End If
    
'    If Val(txtMoney.Text) = 0 Then
'        MsgBox "��������Ľ��!", vbInformation, gstrSysName
'        If txtMoney.Visible And txtMoney.Enabled Then txtMoney.SetFocus
'        Exit Function
'    End If
    
    If Val(txtMoney.Text) > 99999999 Or Val(txtMoney.Text) < 0 Then
        MsgBox "����Ľ�������0-99999999��Χ֮��!", vbInformation, gstrSysName
        If txtMoney.Visible And txtMoney.Enabled Then txtMoney.SetFocus
        Exit Function
    End If
    If Format(dtpDate.Value, "yyyy-MM-dd HH:mm:ss") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        MsgBox "������������ڲ��ܴ���" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "!", vbInformation, gstrSysName
        If dtpDate.Visible And dtpDate.Enabled Then dtpDate.SetFocus
        Exit Function
    End If
    
    If mEditType = EM_ED_�ϸ� Then
        strSQL = "Select 1 From ��Ա�ɿ���� where �տ�Ա=[1] and ����=1 and ���<>0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlStr.NeedName(cboPerson.Text)))
        If Not rsTemp.EOF Then
            MsgBox "������:" & Trim(cboPerson.Text) & " �Ѿ������տ��¼,�޷������ϸڱ��ý�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select 1 From ��Ա�ݴ��¼ where �տ�Ա=[1] and �ջ�ʱ�� is null And MOD(��¼����,10)=1  and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlStr.NeedName(cboPerson.Text)))
    If Not rsTemp.EOF Then
        If MsgBox("������:" & Trim(cboPerson.Text) & " �Ѿ����ù����ý�,�Ƿ��������?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
            If cboPerson.Visible And cboPerson.Enabled Then cboPerson.SetFocus
            Exit Function
        End If
    End If
    
'    strSQL = "Select Count(1) As ���� From ��Ա�ݴ��¼ Where �տ�Ա=[1] And �ջ�ʱ�� Is Null And ��¼����=11"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(NeedName(cboPerson.Text)))
'    If Val(Nvl(rsTemp!����)) < 1 Then
'        strSQL = "Select 1 From ��Ա�ɿ���� where �տ�Ա=[1] and ����=1 and ���<>0"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(NeedName(cboPerson.Text)))
'        If Not rsTemp.EOF Then
'            MsgBox "������:" & Trim(cboPerson.Text) & " �Ѿ������տ��¼,�޷����ñ��ý�", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500_1", Me
End Sub
Private Sub dtpDate_Change()
    mblnChange = True
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnload Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cboPerson Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Function LoadPerson() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�Ա��Ϣ
    '���:blnFilter-�Ƿ���й���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct A.ID,A.���,A.����,A.����,M.���� as ��������,a.�Ա�" & _
    "   From ��Ա�� A,��Ա����˵�� B, ������Ա C,���ű� M" & _
    "   Where A.id = B.��ԱID And B.��Ա���� In ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���')  " & _
    "               And A.ID=C.��ԱID and C.����ID=M.ID(+) And C.ȱʡ(+)=1 " & _
    "               And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
    "   Order By ���"
    Set mrsChargePerson = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ�Ա��Ϣ")
    If mrsChargePerson.RecordCount = 0 Then
        MsgBox "������һ����Ա����Ϊ: " & vbCrLf & _
                      "     ����Һ�Ա,�����շ�Ա,Ԥ���տ�Ա,סԺ����Ա,��Ժ�Ǽ�Ա,�����Ǽ��� " & vbCrLf & _
                      "���շ���Ա,����[��Ա����]�н�������!", vbInformation + vbOKOnly, gstrSysName
        cboPerson.Clear
        Exit Function
    End If
    With cboPerson
        .Clear
        Do While Not mrsChargePerson.EOF
            .AddItem Nvl(mrsChargePerson!���) & "-" & Nvl(mrsChargePerson!����)
            .ItemData(.NewIndex) = Val(Nvl(mrsChargePerson!ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            If Nvl(mrsChargePerson!����) = mstr������ Then .ListIndex = .NewIndex
            mrsChargePerson.MoveNext
        Loop
        'If .ListCount <> 0 Then .ListIndex = 0
    End With
    LoadPerson = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    mblnFirst = True
    Call SetCtrlEnabled
    mblnUnload = Not LoadCardData
    If mblnUnload Then Exit Sub
    mblnChange = False
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub
 
Private Sub txtMemo_GotFocus()
    zlCommFun.OpenIme True
    zlcontrol.TxtSelAll txtMemo
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    zlcontrol.TxtCheckKeyPress txtMemo, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtMoney_Change()
    mblnChange = True
End Sub
Private Function CheckPersonExists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ������շ�Ա�����б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:���ڷ���true,���򷵻�False
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str���� Then
            If blnLocateItem Then cboPerson.ListIndex = i
            CheckPersonExists = True
            Exit Function
        End If
    Next
End Function

Private Sub txtMoney_GotFocus()
    zlCommFun.OpenIme False
    zlcontrol.TxtSelAll txtMoney
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    zlcontrol.TxtCheckKeyPress txtMoney, KeyAscii, m���ʽ
End Sub
