VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#3.0#0"; "zlIDKind.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBlackListRecordFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSplit 
      Height          =   90
      Left            =   -45
      TabIndex        =   23
      Top             =   3210
      Width           =   7455
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   1
      Left            =   1860
      TabIndex        =   4
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   2
      Left            =   1860
      TabIndex        =   7
      Top             =   1245
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.ComboBox cbo����ԭ�� 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   960
      TabIndex        =   14
      Text            =   "cbo����ԭ��"
      Top             =   2085
      Width           =   5400
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "������ʱ���ѯ"
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   6
      Top             =   1335
      Width           =   1695
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   90
      TabIndex        =   19
      Top             =   3435
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5265
      TabIndex        =   18
      Top             =   3405
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4050
      TabIndex        =   17
      Top             =   3405
      Width           =   1100
   End
   Begin VB.TextBox txt������ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4245
      MaxLength       =   18
      TabIndex        =   12
      Top             =   1710
      Width           =   2100
   End
   Begin VB.TextBox txt�Ǽ��� 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   960
      MaxLength       =   64
      TabIndex        =   10
      Top             =   1710
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   0
      Left            =   4245
      TabIndex        =   2
      Top             =   480
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "������ʱ���ѯ"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   540
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   1
      Left            =   4245
      TabIndex        =   5
      Top             =   870
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "������ʱ���ѯ"
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   930
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   2
      Left            =   4245
      TabIndex        =   8
      Top             =   1275
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin zlIDKind.PatiIdentify patiFind 
      Height          =   345
      Left            =   960
      TabIndex        =   16
      Top             =   2505
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBlackListRecordFilter.frx":0000
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   2
      InputAppearance =   2
      ShowSortName    =   -1  'True
      DefaultCardType =   "���￨"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   540
      TabIndex        =   15
      Top             =   2595
      Width           =   360
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   2
      Left            =   3990
      TabIndex        =   22
      Top             =   1335
      Width           =   180
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   1
      Left            =   3990
      TabIndex        =   21
      Top             =   930
      Width           =   180
   End
   Begin VB.Label lbl����ԭ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ԭ��"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   0
      Left            =   3990
      TabIndex        =   20
      Top             =   540
      Width           =   180
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   3630
      TabIndex        =   11
      Top             =   1770
      Width           =   540
   End
   Begin VB.Label lbl�Ǽ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ǽ���"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1770
      Width           =   540
   End
End
Attribute VB_Name = "frmBlackListRecordFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcllFilter As Collection
Private mlngModule As Long
Private mlng����ID As Long
Public Function zlShowEdit(ByVal frmMain As Object, ByVal lngModule As Long, ByRef cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�༭����ԭ��
    '���:frmMain-���õ�������
    '    cllFilter-��������
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 17:01:16
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    Set mcllFilter = cllFilter: mlngModule = lngModule
    If mcllFilter Is Nothing Then Set mcllFilter = New Collection
    mblnOK = False
    Me.Show 1, frmMain
    Set cllFilter = mcllFilter
    zlShowEdit = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub LoadDefalutFilterValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ����ֵ
    '����:���˺�
    '����:2018-02-28 14:07:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysdate As Date
    Dim i As Integer
    
    On Error GoTo errHandle
    
    dtSysdate = zlDatabase.Currentdate
    For i = 0 To dtpBegin.UBound
        dtpBegin(i).MaxDate = Format(dtSysdate, "yyyy-MM-dd 23:59:59")
        dtpBegin(i).Value = Format(dtSysdate - 7, "yyyy-MM-dd 00:00:00")
        dtpEnd(i).Value = dtpBegin(i).MaxDate
        dtpEnd(i).MaxDate = dtpBegin(i).MaxDate
    Next
    
    chkDate(0).Value = 1: chkDate(1).Value = 0: chkDate(2).Value = 0
    txt�Ǽ���.Text = ""
    txt������.Text = ""
    cbo����ԭ��.Text = ""
    patiFind.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function GetConsFilter(ByRef cllFilter_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ѯ����
    '���:
    '����:cllFilter-������ص�������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-14 14:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    Set cllFilter_Out = New Collection
    If patiFind.Text <> "" And Val(patiFind.Tag) = 0 Then
        MsgBox "δ�ҵ�ָ���Ĳ��ˣ����ڲ��˴����س����Ҳ���!", vbInformation + vbOKOnly, gstrSysName
        patiFind.SetFocus
        Exit Function
    End If
    
    
    '��ʱ�������������Ϊ������¼��һ��С�������������ݿ��Բ�ʹ������
'    If Val(patiFind.Tag) = 0 And chkDate(0).Value <> 1 And chkDate(1).Value <> 1 And chkDate(2).Value <> 1 Then
'        MsgBox "���ˡ�����ʱ�䡢����ʱ�䡢����ʱ�����Ҫ����Ҫ����һ��������", vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    
    If Val(patiFind.Tag) <> 0 Then
        cllFilter_Out.Add Array("����ID", Val(patiFind.Tag)), "����ID"
    End If
    If Trim(txt�Ǽ���.Text) <> "" Then
        cllFilter_Out.Add Array("�Ǽ���", Trim(txt�Ǽ���.Text)), "�Ǽ���"
    End If
    If Trim(txt������.Text) <> "" Then
        cllFilter_Out.Add Array("������", Trim(txt������.Text)), "������"
    End If
    
    If Trim(cbo����ԭ��.Text) <> "" Then
        cllFilter_Out.Add Array("����ԭ��", Trim(cbo����ԭ��.Text)), "����ԭ��"
    End If
    If chkDate(0).Value = 1 Then
        cllFilter_Out.Add Array("����ʱ��", dtpBegin(0).Value, dtpEnd(0).Value), "����ʱ��"
    End If
    If chkDate(1).Value = 1 Then
        cllFilter_Out.Add Array("����ʱ��", dtpBegin(1).Value, dtpEnd(1).Value), "����ʱ��"
    End If
    If chkDate(2).Value = 1 Then
        cllFilter_Out.Add Array("����ʱ��", dtpBegin(1).Value, dtpEnd(1).Value), "����ʱ��"
    End If
    GetConsFilter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo����ԭ��_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll cbo����ԭ��
    
End Sub

Private Sub cbo����ԭ��_LostFocus()
     zlCommFun.OpenIme False
End Sub

Private Sub chkDate_Click(Index As Integer)
    dtpBegin(Index).Enabled = chkDate(Index).Value = 1
    dtpEnd(Index).Enabled = chkDate(Index).Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cllFilter As Collection
    If GetConsFilter(cllFilter) = False Then Exit Sub
    Set mcllFilter = cllFilter
    mblnOK = True
    Unload Me
End Sub

Private Sub dtpEnd_Change(Index As Integer)
    dtpBegin(Index).MaxDate = dtpEnd(Index).Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is patiFind Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub patiFind_Change()
    mlng����ID = 0
    patiFind.Tag = ""
End Sub

Private Sub patiFind_FindPatiArfter(ByVal objCard As zlOneCardComLib.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlOneCardComLib.clsPatientInfo, objCardData As zlOneCardComLib.clsPatientInfo, strErrMsg As String, blnCancel As Boolean)
    Dim cllFilter As Collection, lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����ID
    End If
    mlng����ID = lngPatiID
    patiFind.Tag = lngPatiID
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    
End Sub

Private Sub LoadDataFromcllFilter()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ����ã�����ȱʡ����
    '����:���˺�
    '����:2018-02-28 14:07:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysdate As Date, rsTemp As ADODB.Recordset
    Dim i As Integer, lng����id As Long, strSQL As String
    Dim varData As Variant
    
    On Error GoTo errHandle
    For i = 1 To mcllFilter.Count
        varData = mcllFilter(i)
        Select Case varData(0)
        Case "����ID"
            lng����id = Val(varData(1))
            strSQL = "Select ���� From ������Ϣ where ����ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����id)
            If Not rsTemp.EOF Then
                patiFind.Text = Nvl(rsTemp!����)
                patiFind.Tag = lng����id
            End If
        Case "����ʱ��"
            dtpBegin(1).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(1).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(1).Value = 1
        Case "����ʱ��"
            dtpBegin(2).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(2).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(2).Value = 1
        Case "����ʱ��"
            dtpBegin(0).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(0).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(0).Value = 1
        Case "����ԭ��"
            cbo����ԭ��.Text = varData(1)
        Case "�Ǽ���"
            txt�Ǽ���.Text = varData(1)
        Case "������"
            txt������.Text = varData(1)
        End Select
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2018-11-14 14:28:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objCards As zlOneCardComLib.Cards, i As Integer, strFind As String, strKindstr As String
    Dim objCard As zlOneCardComLib.Card
    
    strSQL = "Select ����,����,���� From ���ò�����Ϊԭ��  Order by  ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With cbo����ԭ��
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            rsTemp.MoveNext
        Loop
        .ListIndex = -1
    End With
    
    strKindstr = "��|��������￨|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|�ֻ���|0"
    Call patiFind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindstr, gstrProductName)
    Set objCards = patiFind.objIDKind.Cards
    If Not objCards Is Nothing Then
        strFind = Val(zlDatabase.GetPara("�ϴβ������", glngSys, mlngModule, ""))  '����ȱʡ��
        If strFind <> "" Then
            For i = 1 To objCards.Count
                Set objCard = objCards(i)
                If objCard.���� = strFind Then
                    If patiFind.GetKindIndex(objCard.�ӿ����) >= 0 Then
                        patiFind.IDKindIDX = i + 1
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    Call LoadDefalutFilterValue
End Sub

Private Sub cmdDef_Click()
    Call LoadDefalutFilterValue
End Sub

Private Sub Form_Load()
    Call InitFace   '��ʼ������
    Call LoadDataFromcllFilter
End Sub

Private Sub patiFind_FindPatiBefore(ByVal objCard As zlOneCardComLib.Card, blnCard As Boolean, strShowText As String, objCardData As zlOneCardComLib.clsPatientInfo, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, rsTmp As ADODB.Recordset
    Dim strName As String, lng����id As Long
    If objCard.���� Like "*��*��*" And Not blnCard And InStr("-*+/.", Left(Trim(patiFind.Text), 1)) = 0 Then
        If zlstr.ActualLen(patiFind.Text) < 4 Then
            MsgBox "���ݲ�����������,�����������������ֻ��ĸ��ַ�!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
        strName = Trim(patiFind.Text)
        strPati = _
            " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
            "           A.סԺ��,B.���� as ����,A.��ǰ���� as ����," & _
            "           A.��������,A.���֤��,A.��ͥ��ַ,decode(Nvl(A.��Ժ,0),0,'','��') As ��Ժ��־ " & _
            " From ������Ϣ A,���ű� B " & _
            " Where  Rownum <101 And A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And A.���� Like [1] " & _
            "   Order by A.����"
        vRect = zlControl.GetControlRect(patiFind.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, _
            vRect.Left, vRect.Top, patiFind.Height, blnCancel, False, True, _
            strName & "%", "bytSize=1")
        If Not rsTmp Is Nothing Then
            If Nvl(rsTmp!ID) = 0 Then '�����²���
                blnCancel = True: Exit Sub
            Else '�Բ���ID��ȡ
                lng����id = Nvl(rsTmp!ID)
            End If
        Else 'ȡ��ѡ��
            blnCancel = True: Exit Sub
        End If
        If zlGetPati(lng����id, objCardData) = False Then blnCancel = True: Exit Sub
        blnFindPatied = True
    End If
End Sub

Private Sub txt������_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt������
End Sub

Private Sub txt������_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt�Ǽ���_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt�Ǽ���
End Sub

Private Sub txt�Ǽ���_LostFocus()
    zlCommFun.OpenIme False
End Sub

Public Function zlGetPati(ByVal lng����id As String, ByRef objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID,���»�ȡ����
    '����:�Ϸ�,����true,���򷵻�False
    '����:����
    '����:2019-09-04 10:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    Set objPati = New clsPatientInfo
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.����id, a. �����, a.סԺ��, a.����֤��, a.�ѱ�, a.ҽ�Ƹ��ʽ,p.���� as ҽ�Ƹ��ʽ����, a. ����, a.�Ա�, a. ����, a.��������, a.�����ص�, a.���֤��, a.����֤��,  " & _
    "        a.ְҵ, a.����, a.����, a.����, a.ѧ��, a.����״��, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.�໤��, a.��ϵ������, a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, " & _
    "        a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�,a.���ڵ�ַ, a.���ڵ�ַ�ʱ�, a.����, '' as ����, '' as ����" & _
    "   From ������Ϣ A,ҽ�Ƹ��ʽ P" & _
    "   Where  a.ҽ�Ƹ��ʽ=P.����(+) And ����ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����id)
    If rsTemp.EOF Then Exit Function
    objPati.����ID = rsTemp!����ID
    objPati.����� = Nvl(rsTemp!�����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.�Ա� = Nvl(rsTemp!�Ա�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.�������� = Format(rsTemp!��������, "yyyy-mm-dd")
    objPati.������ַ = Nvl(rsTemp!�����ص�)
    objPati.���֤�� = Nvl(rsTemp!���֤��)
    objPati.����֤�� = Nvl(rsTemp!����֤��)
    objPati.ְҵ = Nvl(rsTemp!ְҵ)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ѧ�� = Nvl(rsTemp!ѧ��)
    objPati.����״�� = Nvl(rsTemp!����״��)
    objPati.���� = Nvl(rsTemp!����)
    objPati.��ͥ��ַ = Nvl(rsTemp!��ͥ��ַ)
    objPati.��ͥ�绰 = Nvl(rsTemp!��ͥ�绰)
    objPati.��ͥ�ʱ� = Nvl(rsTemp!��ͥ��ַ�ʱ�)
    objPati.�໤�� = Nvl(rsTemp!�໤��)
    objPati.��ϵ�� = Nvl(rsTemp!��ϵ������)
    objPati.��ϵ�˹�ϵ = Nvl(rsTemp!��ϵ�˹�ϵ)
    objPati.��ϵ�˵�ַ = Nvl(rsTemp!��ϵ�˵�ַ)
    objPati.��ϵ�˵绰 = Nvl(rsTemp!��ϵ�˵绰)
    objPati.������λ = Nvl(rsTemp!������λ)
    objPati.������λ�绰 = Nvl(rsTemp!��λ�绰)
    objPati.������λ�ʱ� = Nvl(rsTemp!��λ�ʱ�)
    objPati.������λ������ = Nvl(rsTemp!��λ������)
    objPati.������λ�������ʻ� = Nvl(rsTemp!��λ�ʺ�)
    objPati.���ڵ�ַ = Nvl(rsTemp!���ڵ�ַ)
    objPati.���ڵ�ַ�ʱ� = Nvl(rsTemp!���ڵ�ַ�ʱ�)
    objPati.���� = Nvl(rsTemp!����)
    objPati.���� = Nvl(rsTemp!����)
    objPati.ҽ�Ƹ��ʽ���� = Nvl(rsTemp!ҽ�Ƹ��ʽ����)
    objPati.ҽ�Ƹ��ʽ = Nvl(rsTemp!ҽ�Ƹ��ʽ)
    zlGetPati = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

