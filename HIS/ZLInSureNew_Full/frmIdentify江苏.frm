VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboҽ�Ʒ�ʽ 
      Height          =   300
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   150
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5535
      TabIndex        =   3
      Top             =   4695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4410
      TabIndex        =   2
      Top             =   4695
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   4035
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   6495
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   1455
         TabIndex        =   33
         Top             =   3510
         Width           =   4740
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   4125
         TabIndex        =   31
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   1095
         TabIndex        =   29
         Top             =   3045
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   4125
         TabIndex        =   27
         Top             =   2595
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   1095
         TabIndex        =   25
         Top             =   2595
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   4125
         TabIndex        =   23
         Top             =   2145
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   1095
         TabIndex        =   21
         Top             =   2145
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1095
         TabIndex        =   19
         Top             =   1695
         Width           =   5100
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4125
         TabIndex        =   17
         Top             =   1230
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1095
         TabIndex        =   15
         Top             =   1230
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4125
         TabIndex        =   13
         Top             =   780
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1095
         TabIndex        =   11
         Top             =   780
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4125
         TabIndex        =   9
         Top             =   330
         Width           =   2070
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1095
         TabIndex        =   7
         Top             =   330
         Width           =   2070
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���������ⲡ"
         Height          =   180
         Index           =   14
         Left            =   315
         TabIndex        =   32
         Top             =   3585
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ְ��"
         Height          =   180
         Index           =   13
         Left            =   3330
         TabIndex        =   30
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   12
         Left            =   315
         TabIndex        =   28
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   11
         Left            =   3690
         TabIndex        =   26
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   10
         Left            =   315
         TabIndex        =   24
         Top             =   2670
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���"
         Height          =   180
         Index           =   9
         Left            =   3330
         TabIndex        =   22
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   8
         Left            =   675
         TabIndex        =   20
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Index           =   7
         Left            =   315
         TabIndex        =   18
         Top             =   1770
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��λID"
         Height          =   180
         Index           =   6
         Left            =   3510
         TabIndex        =   16
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���֤��"
         Height          =   180
         Index           =   5
         Left            =   315
         TabIndex        =   14
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   4
         Left            =   3330
         TabIndex        =   12
         Top             =   855
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   3
         Left            =   675
         TabIndex        =   10
         Top             =   855
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   8
         Top             =   405
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����ID"
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   6
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdComp 
      Caption         =   "��֤(&S)"
      Height          =   350
      Left            =   3120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   1100
   End
   Begin VB.TextBox txtNo 
      Height          =   300
      Left            =   1005
      MaxLength       =   10
      TabIndex        =   1
      Top             =   150
      Width           =   2070
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "����֤��"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strVoucherID As String, mbytType As Byte, mlng����ID As Long, intReturn As Long
Private strArrInfo(20) As String, sngArrInfo(20) As Single, iLoop As Long
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, lng����ID As Long) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show vbModal
    gintҽ�Ʒ�ʽ = cboҽ�Ʒ�ʽ.ListIndex + 1
    
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cboҽ�Ʒ�ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdComp_Click()
    Dim int��Ч���� As Integer
    Dim lng����ID As Long
    Dim datCurr As Date
    Dim strReadCard As String, intRecType As Long
    Dim intInsID As Long, i As Long, rsTemp As New ADODB.Recordset
    strReadCard = GetSetting(appName:="ZLSOFT", Section:="ҽ����Ϣ", Key:="ReadCard", Default:="0")
    gblnReadCard = Not strReadCard = "0"
    
    If mbytType = 1 Then
        intRecType = 1
    Else
        intRecType = 0
    End If
    
    If strReadCard = "0" Then
        If IsNumeric(txtNO.Text) = False Or Len(txtNO.Text) <> 10 Then
            MsgBox "������10λ����֤�ţ�����֤���еġ�-���������룩��", vbInformation, gstrSysName
            txtNO.SetFocus
            Exit Sub
        Else
            strVoucherID = Left(txtNO.Text, 5) & "-" & Right(txtNO.Text, 5)
        End If
    Else
        strVoucherID = ""
    End If
    
    If mbytType = 1 Or mbytType = 3 Then
        gstrRecCode = String(12, " ")
        intReturn = FGetRecCode(intRecType, gstrRecCode)
        If intReturn <> 0 Then
            MsgBox "�ڻ�ȡ�շ���ˮ��ʱ��������δ��ô�����Ϣ��", vbExclamation, gstrSysName
            cmdOK.Enabled = False
            Exit Sub
        End If
    ElseIf mbytType = 0 Then                    '��ˢ��������ʱ��������������������ö�����ʽ����Ҫ�����޸�
        gstrSQL = "Select * From �����ʻ� where ����=[1] And ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_����, strVoucherID)
        If rsTemp.EOF Then
            MsgBox "�Ҳ������˵ĹҺ���Ϣ����ȷ�ϸò����Ƿ�ҺŻ򿨺������Ƿ���ȷ��", vbInformation, "����"
            Exit Sub
        End If
        gstrRecCode = Nvl(rsTemp!����֤��)
    Else
        gstrRecCode = String(12, " ")
    End If
    
    If InStr(gstrRecCode, Chr(0)) > 0 Then gstrRecCode = Left(gstrRecCode, InStr(gstrRecCode, Chr(0)) - 1)
'===============================================================================================================
'���ܣ���ȡ�α��˵Ļ�����Ϣ���ʻ���֧����Ϣ
'��ڲ��������ͣ�0����/1סԺ��,�շ���ˮ��,����֤��
'���ڲ�����0����ID,1����,2����,3��������,4���֤����,5��λID,6��λ����,7�Ա�(��/Ů),8��Ա���,9��������,10����,
'          11�������,12����ְ��,13�������ⲡ��(�ѻ�����'δ֪'/����δ����ʱ����'δ����'/������������������ⲡ��),
'          14����(����),15�����ۼ�סԺ����,16�ʻ�����,17�ʻ���֧,18֧���汾��,19����ͳ��֧���ۼ�,20����󲡻���֧���ۼ�,
'          21���깫��Ա����/��ҵ����֧���ۼ�,22������ͨ��������ۼ�,23������ͨ����������Χ�ڷ����ۼ�,
'          24������������������Χ�ڷ����ۼ�,25�������סԺ������Χ�ڷ����ۼ�,26������ͨסԺ�����ۼ�,
'          27������ͨסԺ������Χ�ڷ����ۼ�,28�����ͥ����סԺ������Χ�ڷ����ۼ�,29����1,30����2,
'          31���괢���ʻ�֧���ۼ�,32������������֧���ۼ�,33�����ֽ�֧���ۼ�,34�ʻ����
'===============================================================================================================
'            psCardID        : pChar;     0         //O����[C16]
'        psName          : pChar;    1         //O����[C8]
'        psAreaCode      : pChar;     2        //O��������[C3]
'        psQueryID       : pChar;     3        //O���֤����[C18]
'        psUnitID        : pChar;     4         //O��λID[C8]
'        psUnitName      : pChar;    5         //O��λ����[C50]
'        psSex           : pChar;    6         //O�Ա�[C2](��/Ů)
'        psKind          : pChar;    7         //O��Ա���[C4]
'        psBirthday      : pChar;     8         //O��������[C10](YYYY-MM-DD)
'        psNational      : pChar;     9         //O����[C20]
'        psIndustry      : pChar;     10         //O�������[C20]
'        psDuty          : pChar;    11         //O����ְ��[C30]
'        psChronic       :pChar;     12     // O�������ⲡ��[C200](�ѻ�����'δ֪'/����δ����ʱ����'δ����'/������������������ⲡ��)
'        psOthers1       :pChar;     13     // O����(����)[C200]
    strArrInfo(0) = String(16, " ")
    strArrInfo(1) = String(8, " ")
    strArrInfo(2) = String(3, " ")
    strArrInfo(3) = String(18, " ")
    strArrInfo(4) = String(8, " ")
    strArrInfo(5) = String(50, " ")
    strArrInfo(6) = String(2, " ")
    strArrInfo(7) = String(4, " ")
    strArrInfo(8) = String(10, " ")
    strArrInfo(9) = String(20, " ")
    strArrInfo(10) = String(20, " ")
    strArrInfo(11) = String(30, " ")
    strArrInfo(12) = String(200, " ")
    strArrInfo(13) = String(200, " ")
    intReturn = FGetCardInfo(intRecType, gstrRecCode, strVoucherID, intInsID, strArrInfo(0), strArrInfo(1), _
        strArrInfo(2), strArrInfo(3), strArrInfo(4), strArrInfo(5), strArrInfo(6), strArrInfo(11), strArrInfo(8), _
        strArrInfo(9), strArrInfo(10), strArrInfo(7), strArrInfo(12), strArrInfo(13), sngArrInfo(19), sngArrInfo(0), _
        sngArrInfo(1), sngArrInfo(2), sngArrInfo(3), sngArrInfo(4), sngArrInfo(5), sngArrInfo(6), sngArrInfo(7), _
        sngArrInfo(8), sngArrInfo(9), sngArrInfo(10), sngArrInfo(11), sngArrInfo(12), sngArrInfo(13), sngArrInfo(14), _
        sngArrInfo(15), sngArrInfo(16), sngArrInfo(17), sngArrInfo(18))
    If intReturn <> 0 Then
        MsgBox "�ڻ�ȡ�ղ�����Ϣʱ��������δ��ô�����Ϣ��", vbExclamation, gstrSysName
        cmdOK.Enabled = False
        txtNO.SetFocus
        Exit Sub
    End If
    
    '����ǹҺ�����м�飬���ղ�����ָ�������ڲ������ٴιҺ�
    If mbytType = 3 Or mbytType = 0 Then
        '��ȡ�Һ�����
        int��Ч���� = 2
        datCurr = zlDatabase.Currentdate()
        'ȡ����ID
        gstrSQL = "select ����ID from �����ʻ� where ����=[1] and ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����id", TYPE_����, Substr((strArrInfo(0)), 1, 11))
        If rsTemp.RecordCount <> 0 Then
            lng����ID = rsTemp!����ID
            gstrSQL = " Select MAX(����ID) AS ����ID From ������ü�¼" & _
                      " Where ��¼����=1 and �շ���� in('5','6','7') And ��¼״̬=1 And ����ID=[1]" & _
                      " And �Ǽ�ʱ�� Between to_date('" & Format(DateAdd("d", -1 * int��Ч����, datCurr), "yyyy-MM-dd") & " 00:00:00" & "','yyyy-MM-dd hh24:mi:ss')" & _
                      " And to_date('" & Format(datCurr, "yyyy-MM-dd") & " 23:59:59" & "','yyyy-MM-dd hh24:mi:ss')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ч�������һ����Ч�ĹҺ�����", lng����ID)
            If Nvl(rsTemp!����ID, 0) <> 0 Then
                If MsgBox("3�����ѹҺž����һ�Σ��Ƿ������ٴξ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If
    
    For iLoop = 0 To 20
        If InStr(strArrInfo(iLoop), Chr(0)) > 0 Then
            strArrInfo(iLoop) = Left(strArrInfo(iLoop), InStr(strArrInfo(iLoop), Chr(0)) - 1)
        End If
    Next
    txtInfo(0).Text = intInsID
    For i = 1 To txtInfo.UBound
        txtInfo(i).Text = strArrInfo(i - 1)
    Next
    
    cmdOK.Enabled = True
    cboҽ�Ʒ�ʽ.SetFocus
End Sub

Private Sub cmdOK_Click()
    If Me.txtInfo(0).Text = "" Then
        MsgBox "����ȷ��������ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����;9.˳���;
    '10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
    
'���ڲ�����0����ID,1����,2����,3��������,4���֤����,5��λID,6��λ����,7�Ա�(��/Ů),8��Ա���,9��������,10����,
'          11�������,12����ְ��,13�������ⲡ��(�ѻ�����'δ֪'/����δ����ʱ����'δ����'/������������������ⲡ��),
'          14����(����),0�����ۼ�סԺ����,1�ʻ�����,2�ʻ���֧,3֧���汾��,4����ͳ��֧���ۼ�,5����󲡻���֧���ۼ�,
'          6���깫��Ա����/��ҵ����֧���ۼ�,7������ͨ��������ۼ�,8������ͨ����������Χ�ڷ����ۼ�,
'          9������������������Χ�ڷ����ۼ�,10�������סԺ������Χ�ڷ����ۼ�,11������ͨסԺ�����ۼ�,
'          12������ͨסԺ������Χ�ڷ����ۼ�,13�����ͥ����סԺ������Χ�ڷ����ۼ�,14����1,15����2,
'          16���괢���ʻ�֧���ۼ�,17������������֧���ۼ�,18�����ֽ�֧���ۼ�,19�ʻ����
    
    mstrPatient = "": mstrOther = ""
    mstrPatient = mstrPatient & strArrInfo(0) & ";"             '����
    mstrPatient = mstrPatient & strVoucherID & ";"              'ҽ����
    mstrPatient = mstrPatient & ";"                             '����
    mstrPatient = mstrPatient & strArrInfo(1) & ";"             '����
    mstrPatient = mstrPatient & strArrInfo(6) & ";"             '�Ա�
    mstrPatient = mstrPatient & strArrInfo(8) & ";"             '��������
    mstrPatient = mstrPatient & strArrInfo(3) & ";"             '���֤��
    mstrPatient = mstrPatient & strArrInfo(5) & ";"             '��λ����
 
    mstrOther = mstrOther & ";"                                 '����
    mstrOther = mstrOther & ";"                                 '˳���
    mstrOther = mstrOther & strArrInfo(11) & ";"                '10��Ա���
    mstrOther = mstrOther & sngArrInfo(18) & ";"                '11�ʻ����
    mstrOther = mstrOther & ";"                                 '12��ǰ״̬
    mstrOther = mstrOther & ";"                                 '13����ID
    mstrOther = mstrOther & strArrInfo(7) & ";"                 '14��ְ
    mstrOther = mstrOther & ";"                                 '15����֤��
    mstrOther = mstrOther & ";"                                 '16�����
    mstrOther = mstrOther & ";"                                 '17�Ҷȼ�
    mstrOther = mstrOther & sngArrInfo(0) & ";"                '18�ʻ������ۼ�
    mstrOther = mstrOther & sngArrInfo(1) & ";"                '19�ʻ�֧���ۼ�
    mstrOther = mstrOther & ";"                                 '20����ͳ���ۼ�
    mstrOther = mstrOther & sngArrInfo(4) & ";"                '21ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & sngArrInfo(19) & ";"                '22סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                 '23��������
    
    Me.Hide
End Sub

Private Sub Form_Load()
    cboҽ�Ʒ�ʽ.AddItem "��ͨ����"
    cboҽ�Ʒ�ʽ.AddItem "��ͨסԺ"
    cboҽ�Ʒ�ʽ.AddItem "���ⲡ"
    cboҽ�Ʒ�ʽ.AddItem "��������"
    cboҽ�Ʒ�ʽ.AddItem "����"
    cboҽ�Ʒ�ʽ.ListIndex = 0
End Sub

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO.Text)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdComp_Click
    End If
End Sub
