VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˻�����Ϣ����"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton optType 
      Caption         =   "סԺ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3390
      TabIndex        =   12
      Top             =   2085
      Width           =   870
   End
   Begin VB.OptionButton optType 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2115
      TabIndex        =   11
      Top             =   2085
      Width           =   855
   End
   Begin VB.ComboBox cmbNum 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030A
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":030C
      TabIndex        =   14
      Text            =   "cmbNum"
      Top             =   2475
      Width           =   2070
   End
   Begin VB.ComboBox cboAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   2  'OFF
      Left            =   2115
      TabIndex        =   8
      Top             =   1590
      Width           =   1350
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030E
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   675
      Width           =   2070
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   100
      TabIndex        =   1
      Top             =   210
      Width           =   2070
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2010
      TabIndex        =   15
      Top             =   3210
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3345
      TabIndex        =   16
      Top             =   3210
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   15
      TabIndex        =   17
      Top             =   2925
      Width           =   5310
   End
   Begin MSMask.MaskEdBox medBirthdayTime 
      Height          =   360
      Left            =   3480
      TabIndex        =   6
      Top             =   1140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medBirthdayDate 
      Bindings        =   "frmModiPatiBaseInfo.frx":0312
      Height          =   360
      Left            =   2115
      TabIndex        =   5
      Top             =   1140
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "YYYY-MM-DD"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   10
      Top             =   2085
      Width           =   960
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Һŵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   13
      Top             =   2535
      Width           =   960
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   495
      Picture         =   "frmModiPatiBaseInfo.frx":031D
      Top             =   345
      Width           =   480
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1500
      TabIndex        =   7
      Top             =   1650
      Width           =   480
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1530
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mstrģ�� As String
Private mint���� As Integer
Private mstrInfo As String
Private mblnChange As Boolean
Private mblnDrop As Boolean
Private mrsTmp As New ADODB.Recordset

Public Function ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strģ�� As String, ByRef strInfo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:lng����ID-����ID
    '     lng����ID=������0��ʾĳһ��סԺ����ҳId(�����Զ���λ��Ҫ�޸ĵ�ĳһ��סԺ)������0��ʾ��Ҫ�û��ֹ�ѡ�������ﻹ��סԺ
    '     strģ��=���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '����:
    '����:
    '����:������
    '����:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mstrģ�� = strģ��
    mblnChange = False
    mblnOK = False
    '��ȡ���˻�����Ϣ
    If Not LoadPatiBaseInfo Then ShowMe = False: Exit Function
    
    Me.Show 1, frmParent
    strInfo = Trim(mstrInfo)
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    txtName.Text = ""
    txtName.MaxLength = GetColumnLength("������Ϣ", "����")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.ListIndex = 0
    txtAge.MaxLength = GetColumnLength("������Ϣ", "����")
    
    cboSex.Clear
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "�Ա�")
    Do While Not rsTmp.EOF
        cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
            cboSex.ItemData(cboSex.NewIndex) = 1
        End If
    rsTmp.MoveNext
    Loop
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo ErrHand
    
    If mlng����ID <> 0 Then 'סԺ����
        strSQL = " Select Nvl(a.����, b.����) ����, Nvl(a.�Ա�, b.�Ա�) �Ա�,a.����,B.��������" & vbNewLine & _
                " From ������ҳ a, ������Ϣ b" & vbNewLine & _
                " Where a.����id = b.����id And a.����id = [1] And a.��ҳid = [2]"
    Else
        strSQL = "Select ����,�Ա�,����,�������� From ������Ϣ Where ����ID=[1]"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˻�����Ϣ", mlng����ID, mlng����ID)
    
    mblnChange = False
    
    If Not rsTmp.EOF Then
        '������Ϣ��ʼ��
        Call InitDicts
        
        txtName.Text = zlCommFun.Nvl(rsTmp!����)
        
        cboSex.ListIndex = GetCboIndex(cboSex, Nvl(rsTmp!�Ա�))
        If cboSex.ListIndex = -1 And Not IsNull(rsTmp!�Ա�) Then
            cboSex.AddItem rsTmp!�Ա�, 0
            cboSex.ListIndex = cboSex.NewIndex
        End If
           
        Call LoadOldData("" & rsTmp!����, txtAge, cboAge)
        mblnChange = False
        medBirthdayDate.Text = Format(IIf(IsNull(rsTmp!��������), "____-__-__", rsTmp!��������), "YYYY-MM-DD")
        mblnChange = True
        
        If Not IsNull(rsTmp!��������) Then
            If CDate(medBirthdayDate.Text) - CDate(rsTmp!��������) <> 0 Then medBirthdayTime.Text = Format(rsTmp!��������, "HH:MM")
        Else
            medBirthdayTime.Text = "__:__"
            mblnChange = False
            medBirthdayDate.Text = ReCalcBirth(Val(txtAge.Text), cboAge.Text)
            mblnChange = True
        End If
    Else
        MsgBox "��ȡ���˻�����Ϣʧ��,����ȷ��Ҫ������Ϣ�����Ĳ��ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call LoadPatiData
    
    mblnChange = True
    
    LoadPatiBaseInfo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadPatiData()
'-----------------------------------------------
'����:��ȡ���˾����¼��Ϣ(סԺ����������¼)
'
'-----------------------------------------------
    Dim strSQL As String
    Dim bln���� As Boolean, blnסԺ As Boolean
    
    On Error GoTo ErrHand
    strSQL = _
        " Select 1 ����,ID Id, No, to_char(�Ǽ�ʱ��,'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��" & vbNewLine & _
        " From ���˹Һż�¼" & vbNewLine & _
        " Where ����id = [1] And Mod(��¼״̬, 2) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select 2 ����,��ҳId Id, '' || ��ҳid No, to_char(�Ǽ�ʱ��,'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��" & vbNewLine & _
        " From ������ҳ" & vbNewLine & _
        " Where ����id = [1] And Nvl(��ҳid, 0) <> 0"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����¼", mlng����ID)
    
    optType(0).Enabled = True
    optType(1).Enabled = True
    cmbNum.Clear
    If mrsTmp.RecordCount > 0 Then
        mrsTmp.Filter = "����=1"
        bln���� = mrsTmp.RecordCount > 0
        mrsTmp.Filter = "����=2"
        blnסԺ = mrsTmp.RecordCount > 0
        
        mblnChange = True
        If bln���� = True And blnסԺ = True Then
            If mlng����ID <> 0 Then
                optType(1).Value = True
            Else
                optType(0).Value = True
            End If
        Else
            If bln���� = True Then
                optType(0).Value = True
                optType(1).Enabled = False
            Else
                optType(1).Value = True
                optType(0).Enabled = False
            End If
        End If
    Else
        mblnChange = False
        '���˴�δ�ҺŻ�סԺ
        optType(0).Value = True
        optType(0).Enabled = False
        optType(1).Enabled = False
        lblType.Enabled = False
        lblNum.Enabled = False
        cmbNum.Enabled = False
        mblnChange = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboAge_LostFocus()
    If Trim(txtAge.Text) = "" Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If Not IsDate(medBirthdayDate.Text) Then
        mblnChange = False
        medBirthdayDate.Text = ReCalcBirth(Val(txtAge.Text), cboAge.Text)
        mblnChange = True
    End If
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboSex.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboSex.hwnd, KeyAscii)
    If lngIdx <> -2 Then cboSex.ListIndex = lngIdx
End Sub

Private Sub cmbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmbNum.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then
        mblnDrop = SendMessage(cmbNum.hwnd, &H157, 0, 0) = 1
    End If
End Sub

Private Sub cmbNum_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cmbNum.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cmbNum.Text)
        If cmbNum.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cmbNum.List(cmbNum.ListIndex) Then Call zlControl.CboSetIndex(cmbNum.hwnd, -1)
        End If
        If strText = "" Then
            cmbNum.ListIndex = -1
        ElseIf cmbNum.ListIndex = -1 Then
            intIdx = -1
            strFilter = "����=" & IIf(optType(0).Value = True, 1, 2)
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsTmp)
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsTmp.Filter = strFilter: iCount = 0
            With mrsTmp
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsTmp.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�Ӷ�λ
                        If Nvl(!NO) = strText Then strResult = Nvl(!NO): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!NO)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!NO)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!NO)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!NO) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!NO) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call zlDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    End Select
                    mrsTmp.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!NO)
            'ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckExists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                If optType(0).Value = True Then
                    rsTemp.Sort = "�Ǽ�ʱ�� DESC"
                Else
                    rsTemp.Sort = "ID DESC"
                End If
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1101, cmbNum, rsTemp, True, "", "����", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheckExists(Nvl(rsReturn!NO), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cmbNum: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cmbNum.ListIndex = -1 Then
            cmbNum.Text = ""
            Exit Sub
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
            ElseIf intIdx <> cmbNum.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cmbNum.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If optType(0).Value = True Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        Else
            If InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmbNum_Validate(Cancel As Boolean)
    If cmbNum.Text <> "" Then
        If GetCboIndex(cmbNum, NeedName(cmbNum.Text)) = -1 Then cmbNum.ListIndex = -1: cmbNum.Text = ""
    End If
    If cmbNum.Text = "" And cmbNum.Enabled = True Then '˵��¼�����Ϣ���������б���
        MsgBox "��ѡ��" & IIf(optType(0).Value = True, "�Һŵ���", "סԺ����"), vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Function isCheckExists(ByVal strNo As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cmbNum.ListCount - 1
        If NeedName(cmbNum.List(i)) = strNo Then
            If blnLocateItem Then cmbNum.ListIndex = i
            isCheckExists = True
            Exit Function
        End If
    Next
End Function


Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'���ܣ��������У��ͱ���
    Dim strSQL As String, strInfo As String
    Dim str���� As String, str�������� As String, str�Ա� As String
    Dim lngTmp As Long
    Dim cmdTmp As New ADODB.Command
    Dim cmdPara As New ADODB.Parameter
    Dim D�������� As Date
    '��һ�������ݺϷ���У��
    If Trim(txtName.Text) = "" Then
        MsgBox "�������벡�˵�������", vbInformation, gstrSysName
        If txtName.Enabled And txtName.Visible Then txtName.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "����ȷ�����˵��Ա�", vbInformation, gstrSysName
        If cboSex.Enabled And cboSex.Visible Then cboSex.SetFocus: Exit Sub
    End If
    
    If Not IsDate(medBirthdayDate.Text) Then
        MsgBox "������ȷ���벡�˵ĳ������ڣ�", vbInformation, gstrSysName
        If medBirthdayDate.Enabled And medBirthdayDate.Visible Then medBirthdayDate.SetFocus: Exit Sub
    End If
    
    If Trim(txtAge.Text) = "" Then
        MsgBox "�������벡�˵����䣡", vbInformation, gstrSysName
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    
    If IsDate(medBirthdayDate.Text) Then
        lngTmp = GetOldAcademic(CDate(medBirthdayDate.Text), cboAge.Text)
        If (lngTmp <> 0 And lngTmp <> Val(txtAge.Text)) Or (lngTmp = 0 And lngTmp <> Val(txtAge.Text) And Not CDate(medBirthdayDate.Text) = CDate(0) And InStr(" ������", cboAge.Text) > 1) Then
            strInfo = ""
            If lngTmp = 0 Then strInfo = ReCalcOld(CDate(medBirthdayDate.Text), cboAge, 0, False)
            If strInfo = "" Then
                strInfo = lngTmp & cboAge.Text
            End If
            If MsgBox("����ͳ������ڲ�һ�£�" & medBirthdayDate.Text & "��������Ӧ����" & strInfo & "��" & _
                vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If Not CheckTextLength("����", txtName) Then Exit Sub
    If Not CheckTextLength("����", txtAge) Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If cmbNum.Enabled And cmbNum.ListIndex = -1 Then
        MsgBox "����ѡ��" & IIf(optType(0).Value = True, "�Һŵ���", "סԺ����") & "��", vbInformation, gstrSysName
        If cmbNum.Enabled And cmbNum.Visible Then cmbNum.SetFocus: Exit Sub
    End If
    
    
    If medBirthdayTime = "__:__" Then
        str�������� = IIf(IsDate(medBirthdayDate.Text), "TO_Date('" & medBirthdayDate.Text & "','YYYY-MM-DD')", "NULL")
        D�������� = CDate(Format(medBirthdayDate.Text, "YYYY-MM-DD"))
    Else
        str�������� = IIf(IsDate(medBirthdayDate.Text), "TO_Date('" & medBirthdayDate.Text & " " & medBirthdayTime.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
        D�������� = CDate(Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:mm:ss"))
    End If
    If InStr(1, cboSex.Text, "-") <> 0 Then
        str�Ա� = Split(cboSex.Text, "-")(1)
    Else
        str�Ա� = cboSex.Text
    End If
    
    str���� = Trim(txtAge.Text)
    If IsNumeric(str����) Then str���� = str���� & cboAge.Text
    
    If cmbNum.Enabled = True Then
        mint���� = IIf(optType(1).Value = True, 2, 1)
        mlng����ID = Val(cmbNum.ItemData(cmbNum.ListIndex))
    Else
        mint���� = 1
        mlng����ID = 0
    End If
    
    '�ڶ��������ݱ���
     On Error GoTo ErrHand
    Set cmdTmp = New ADODB.Command
    strSQL = "Zl_������Ϣ_������Ϣ����("
'   ����id_In ������Ϣ�䶯.����id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    Set cmdPara = cmdTmp.CreateParameter("����ID", adVarNumeric, adParamInput, 18, mlng����ID)
    cmdTmp.Parameters.Append cmdPara
'   ����id_In Number := Null,
    strSQL = strSQL & "" & mlng����ID & ","
    Set cmdPara = cmdTmp.CreateParameter("����ID", adVarNumeric, adParamInput, 18, mlng����ID)
    cmdTmp.Parameters.Append cmdPara
'   ģ��_In   ������Ϣ�䶯.�䶯ģ��%Type,
    strSQL = strSQL & "'" & mstrģ�� & "',"
    Set cmdPara = cmdTmp.CreateParameter("�䶯ģ��", adVarChar, adParamInput, 100, mstrģ��)
    cmdTmp.Parameters.Append cmdPara
'   ����_In   ������Ϣ.����%Type,
    strSQL = strSQL & "'" & Trim(txtName.Text) & "',"
    Set cmdPara = cmdTmp.CreateParameter("����", adVarChar, adParamInput, 100, Trim(txtName.Text))
    cmdTmp.Parameters.Append cmdPara
'   �Ա�_In   ������Ϣ.�Ա�%Type,
    strSQL = strSQL & "'" & str�Ա� & "',"
    Set cmdPara = cmdTmp.CreateParameter("�Ա�", adVarChar, adParamInput, 100, str�Ա�)
    cmdTmp.Parameters.Append cmdPara
'   ����_In   ������Ϣ.����%Type
    strSQL = strSQL & "'" & str���� & "',"
    Set cmdPara = cmdTmp.CreateParameter("����", adVarChar, adParamInput, 100, str����)
    cmdTmp.Parameters.Append cmdPara
'   ��������_In ������Ϣ.��������%Type,
    strSQL = strSQL & "" & str�������� & ","
'   ����_In   number(1)  --1-����;2-סԺ
    Set cmdPara = cmdTmp.CreateParameter("��������", adDBTimeStamp, adParamInput, , D��������)
    cmdTmp.Parameters.Append cmdPara
    strSQL = strSQL & "" & mint���� & ","
    Set cmdPara = cmdTmp.CreateParameter("����", adVarNumeric, adParamInput, 1, mint����)
    cmdTmp.Parameters.Append cmdPara
'   ˵��_Out    Out ������Ϣ�䶯.˵��%Type --����
    strSQL = strSQL & "" & "" & ")"
    Set cmdPara = cmdTmp.CreateParameter("˵��", adLongVarChar, adParamOutput, 4000)
    cmdTmp.Parameters.Append cmdPara
    cmdTmp.ActiveConnection = gcnOracle
    cmdTmp.CommandType = adCmdStoredProc
    cmdTmp.CommandText = "Zl_������Ϣ_������Ϣ����"
    Call SQLTest(App.ProductName, "Zl_������Ϣ_������Ϣ����", strSQL)
    cmdTmp.Execute
    Call SQLTest
    mstrInfo = Nvl(cmdTmp.Parameters("˵��"), "")
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        If ActiveControl.Name <> txtName.Name And ActiveControl.Name <> txtAge.Name And ActiveControl.Name <> cmbNum.Name Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub medBirthdayDate_Change()
    If IsDate(medBirthdayDate.Text) And mblnChange Then
        mblnChange = False
        medBirthdayDate.Text = Format(CDate(medBirthdayDate.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        
        txtAge.Text = ReCalcOld(CDate(medBirthdayDate.Text), cboAge)
    End If
End Sub

Private Sub medBirthdayDate_GotFocus()
    Call OpenIme
    SelAll medBirthdayDate
End Sub

Private Sub medBirthdayDate_LostFocus()
    If medBirthdayDate.Text <> "____-__-__" And Not IsDate(medBirthdayDate.Text) Then
        medBirthdayDate.SetFocus
    End If
End Sub

Private Sub medBirthdayTime_GotFocus()
    Call OpenIme
    SelAll medBirthdayTime
End Sub

Private Sub medBirthdayTime_KeyPress(KeyAscii As Integer)
    If Not IsDate(medBirthdayDate) Then
        KeyAscii = 0
        medBirthdayTime.Text = "__:__"
    End If
End Sub

Private Sub medBirthdayTime_Validate(Cancel As Boolean)
    If medBirthdayTime.Text <> "__:__" And Not IsDate(medBirthdayTime.Text) Then
        medBirthdayTime.SetFocus
        Cancel = True
    End If
End Sub

Private Sub optType_Click(Index As Integer)
    If mblnChange = False Or mrsTmp Is Nothing Then Exit Sub
    If mrsTmp.State = adStateClosed Then Exit Sub
     
    If Index = 0 Then
        lblNum.Caption = "�Һŵ���"
        mrsTmp.Filter = "����=1"
    ElseIf Index = 1 Then
        lblNum.Caption = "סԺ����"
        mrsTmp.Filter = "����=2"
    End If
    If Index = 0 Or Index = 1 Then
        cmbNum.Clear
        Do While Not mrsTmp.EOF
            cmbNum.AddItem Nvl(mrsTmp!NO)
            cmbNum.ItemData(cmbNum.NewIndex) = Val(mrsTmp!ID)
            If Index = 1 And mlng����ID = Val(mrsTmp!ID) Then
                cmbNum.ListIndex = cmbNum.NewIndex
            End If
        mrsTmp.MoveNext
        Loop
        
        If cmbNum.ListIndex = -1 And cmbNum.ListCount > 0 Then cmbNum.ListIndex = 0
    End If
End Sub

Private Sub txtAge_GotFocus()
    Call zlCommFun.OpenIme
    zlControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        cboAge.ListIndex = -1: cboAge.Visible = False
    ElseIf cboAge.Visible = False Then
        cboAge.ListIndex = 0: cboAge.Visible = True
    End If
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            Call CheckInputLen(txtName, KeyAscii)
        End If
    Else
        If Trim(txtName.Text) = "" Then
            Exit Sub
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme
    txtName.Text = Trim(txtName.Text)
End Sub
