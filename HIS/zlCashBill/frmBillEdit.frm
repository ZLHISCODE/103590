VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillEdit 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ʊ�����õ�"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraCheck 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4815
      TabIndex        =   31
      Top             =   3435
      Width           =   3030
      Begin VB.OptionButton optResult 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   24
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton optResult 
         Caption         =   "���"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   23
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1335
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "�˶Ա�ע(&D)"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblResult 
         Caption         =   "�˶Խ��(&C)"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   4530
      TabIndex        =   27
      Top             =   5910
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   5940
      TabIndex        =   28
      Top             =   5910
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -210
      TabIndex        =   30
      Top             =   5790
      Width           =   8295
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   120
      TabIndex        =   29
      Top             =   5910
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   120
      TabIndex        =   32
      Top             =   750
      Width           =   7725
      Begin VB.ComboBox cbo��� 
         Height          =   360
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2565
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   360
         Left            =   4020
         TabIndex        =   35
         Top             =   1155
         Width           =   285
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   330
         Left            =   2955
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   330
      End
      Begin VB.ComboBox cmb������ 
         Height          =   360
         Left            =   1380
         TabIndex        =   13
         Text            =   "cmb������"
         Top             =   1635
         Width           =   1920
      End
      Begin VB.ComboBox cmbʹ�÷�ʽ 
         Height          =   360
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1635
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   17
         Top             =   2115
         Width           =   1920
      End
      Begin VB.ComboBox cmbƱ�� 
         Height          =   360
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1920
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   1
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1155
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   2
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   3
         Left            =   4650
         MaxLength       =   2
         TabIndex        =   10
         Top             =   1155
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   4
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   5
         Left            =   1395
         MaxLength       =   20
         TabIndex        =   5
         Top             =   705
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   5805
         TabIndex        =   19
         Top             =   2115
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   180617219
         CurrentDate     =   37007
      End
      Begin VB.Label lblUserType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�����(&K)"
         Height          =   240
         Left            =   3735
         TabIndex        =   2
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&G)"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�÷�ʽ(&M)"
         Height          =   240
         Index           =   1
         Left            =   4350
         TabIndex        =   14
         Top             =   1695
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���(&R)"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2175
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��(&D)"
         Height          =   240
         Index           =   3
         Left            =   4350
         TabIndex        =   18
         Top             =   2175
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ������(&K)"
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   0
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���뷶Χ(&B)"
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   1215
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   240
         Index           =   5
         Left            =   4350
         TabIndex        =   34
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label lbl˵�� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   1860
         Left            =   0
         TabIndex        =   20
         Top             =   3015
         Width           =   4605
      End
      Begin VB.Label Label2 
         Caption         =   "��ϸ���"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   2685
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&P)"
         Height          =   240
         Index           =   7
         Left            =   480
         TabIndex        =   4
         Top             =   765
         Width           =   840
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�����õ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2220
      TabIndex        =   21
      Top             =   240
      Width           =   2700
   End
End
Attribute VB_Name = "frmBillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytInFun As Byte '0-Ʊ�����õ�,1-�˶����õ�
Private mstrPrivs As String
Private mlng����ID As Long
Private mstr��� As String
Private mlng���ID As Long

Private mblnIsBIll As Boolean '��ǰƱ���Ƿ�ΪƱ��
Private mintƱ�� As gBillType

Private mblnOK As Boolean
Private mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Private mstr��С���� As String
Private mstr������ As String
Private mstrƱ�ݳ��� As String '��ʾ����Ʊ�ݵĺ��볤�ȣ���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨  77777
Private mlng���� As Long       '��ǰƱ������ĳ���
Private mblnҩ��  As Boolean
Private mrsPerson As ADODB.Recordset
Private mlngPreID As Long
Private mlngModule As Long
Private mbln���ȷ������ As Boolean      '33725
Private mrs���� As ADODB.Recordset
Private mrs�ֶ� As ADODB.Recordset
Private mstr��⿪ʼ�� As String, mstr�������� As String
Private mint�ϴ�Ʊ�� As gBillType
Private mblnNotClick As Boolean
Private mstrPreType(1 To 7) As String '�ϴ�ѡ������
Private mcllCardProperty As Collection  '���ų���,ǰ׺�ı�,����,���ſ���

Private Function Select���Ʊ��(ByVal intƱ�� As gBillType, ByVal objCtl As Object, _
    ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����Ʊ��
    '���:objCtl-�ؼ�(Ŀǰֻ֧���ı���)
    '     strKey-����Ľ�ֵ
    '     intƱ��-��ǰѡ���Ʊ��
    '����:
    '����:���ҳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-18 11:08:09
    '����:33725
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strWhere As String
    Dim str��ʼ���� As String, intǰ׺ As Integer, lng���� As Long, str��ֹ���� As String
    Dim blnCancel As Boolean, sngX As Single, sngY As Single, lngH As Long, i As Long
    Dim vRect As RECT, strSearch1 As String, blnFind As Boolean
    Dim str��� As String
    
    mlng���ID = 0
    If Not mbln���ȷ������ Then zlCommFun.PressKey vbKeyTab: Exit Function
    'zlDatabase.ShowSQLSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmMain=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Dim strʹ����� As String
    
    strSearch1 = strKey
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        If zlCommFun.IsNumOrChar(strKey) Then
            If intƱ�� = gBillType.���ѿ� Then
                strWhere = " And (A.ID=[3] or A.��ʼ���� like upper([2]) or A.��ֹ���� like upper([2]))"
            Else
                strWhere = " And (A.ID=[3] or A.��ʼ���� like upper([2]) or A.��ֹ���� like upper([2]))"
            End If
        Else
            strWhere = " And (A.�Ǽ��� like upper([2]) or A.��ע like upper([2]) )"
        End If
        strKey = GetMatchingSting(strKey, False)
    End If
    
    Select Case intƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        strWhere = strWhere & " And nvl(A.ʹ�����,'LXH')=[4]"
        str��� = Trim(cbo���.Text)
        If str��� = "" Then str��� = "LXH"
        strʹ����� = " A.ʹ�����,"
    Case gBillType.Ԥ���վ�
        If cbo���.ListIndex < 0 Then Exit Function
        '58071
        strWhere = strWhere & " And nvl(A.ʹ�����,'0')=[4]"
        str��� = cbo���.ItemData(cbo���.ListIndex)
        strʹ����� = " decode(nvl(A.ʹ�����,'0'),'0','','1','����','סԺ') as  ʹ�����,"
    Case gBillType.���￨
        If cbo���.ListIndex < 0 Then Exit Function
        strWhere = strWhere & " And nvl(A.ʹ�����,'0')=[4]"
        str��� = cbo���.ItemData(cbo���.ListIndex)
        strʹ����� = " nvl(A.ʹ�����,'���￨') as ʹ�����,"
    End Select
    
    If intƱ�� = gBillType.���ѿ� Then
        If cbo���.ListIndex < 0 Then Exit Function
        str��� = cbo���.ItemData(cbo���.ListIndex)
        
        gstrSQL = _
            "Select A.Id, A.���� as �������,A.�ӿڱ�� as ʹ�����ID, nvl(M.����,'���ѿ�') As ʹ�����," & vbNewLine & _
            "       A.ǰ׺�ı�,A.��ʼ����, A.��ֹ����, A.�������, A.ʣ������," & vbNewLine & _
            "       A.��ע, A.�Ǽ���, A.�Ǽ�ʱ��" & vbNewLine & _
            "From ���ѿ�����¼ A, ���ѿ����Ŀ¼ M" & vbNewLine & _
            "Where a.�ӿڱ�� = m.���(+) And a.�ӿڱ��=[4]" & vbNewLine & _
            "      And nvl(A.ʣ������,0)>0 And A.�Ƿ���ڿ�=1"
    Else
        gstrSQL = _
            "  Select A.Id, A.���� as �������,A.ʹ����� as ʹ�����ID," & strʹ����� & "A.ǰ׺�ı�,  " & _
            "          A.��ʼ����, A.��ֹ����, A.�������, A.ʣ������, A.��ע, A.�Ǽ���, A.�Ǽ�ʱ�� " & _
            "  From Ʊ������¼ A " & IIf(intƱ�� = 5, ",ҽ�ƿ���� M", "") & _
            "  Where nvl(A.ʣ������,0)>0 And A.Ʊ��=[1] And A.����Ʊ��=1  " & strWhere & _
                   IIf(intƱ�� = 5, " And to_number(nvl(A.ʹ�����,'0'))=M.ID(+) ", "")
    End If
    
   '���궨λ
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    sngX = vRect.Left - 15
    sngY = vRect.Top
    lngH = objCtl.Height
     
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����Ʊ��ѡ��", False, "", "", False, False, True, _
        sngX, sngY, lngH, blnCancel, False, False, intƱ��, strKey, Val(strSearch1), str���)
    
   If blnCancel = True Then
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "δ�ҵ�����������Ʊ������¼,����"
        Call zlControl.ControlSetFocus(objCtl, True)
        If UCase(TypeName(objCtl)) = UCase("TextBox") Then zlControl.TxtSelAll objCtl
        Exit Function
    End If
    Call zlControl.ControlSetFocus(objCtl, True)
    objCtl.Text = NVL(rsTemp!�������)
    objCtl.Tag = NVL(rsTemp!�������)
    mlng���ID = rsTemp!ID
    
    If intƱ�� = gBillType.���ѿ� Then
        str��ʼ���� = NVL(rsTemp!��ʼ����): str��ֹ���� = NVL(rsTemp!��ֹ����)
    Else
        str��ʼ���� = NVL(rsTemp!��ʼ����): str��ֹ���� = NVL(rsTemp!��ֹ����)
    End If
    txtEdit(1).Text = Trim(NVL(rsTemp!ǰ׺�ı�))
    intǰ׺ = Len(txtEdit(1).Text)
    txtEdit(2).Text = Trim(Mid(str��ʼ����, intǰ׺ + 1))
    txtEdit(2).Tag = txtEdit(2).Text
    lng���� = Len(txtEdit(2).Text)
    txtEdit(3).Text = NVL(rsTemp!ǰ׺�ı�)
    txtEdit(4).Text = Mid(str��ֹ����, intǰ׺ + 1)
    txtEdit(4).Tag = txtEdit(4).Text
    blnFind = False
    With cbo���
        mblnNotClick = True
        For i = 0 To .ListCount - 1
            If intƱ�� = gBillType.Ԥ���վ� Or intƱ�� = gBillType.���￨ Or intƱ�� = gBillType.���ѿ� Then
              If .ItemData(i) = Val(NVL(rsTemp!ʹ�����ID)) Then
                    blnFind = True
                    .ListIndex = i: Exit For
              End If
            Else
                If Trim(.List(i)) = Trim(NVL(rsTemp!ʹ�����ID)) Then
                    blnFind = True
                    .ListIndex = i: Exit For
                End If
            End If
        Next
        
        If blnFind = False _
            And Not (intƱ�� = gBillType.Ԥ���վ� Or intƱ�� = gBillType.���￨ Or intƱ�� = gBillType.���ѿ�) Then
            .AddItem NVL(rsTemp!ʹ�����ID, " ")
            .ListIndex = .NewIndex
        End If
        .Tag = .Text
        mblnNotClick = False
    End With
    
    mstr��⿪ʼ�� = str��ʼ����: mstr�������� = str��ֹ����:
    Call Load�ֶ�Ʊ��(intƱ��, Trim(objCtl.Text), mstr��⿪ʼ��, mstr��������)
    Dim varTemp As Variant
    If mrs�ֶ�.RecordCount <> 0 Then
        mrs�ֶ�.MoveFirst
        varTemp = Split(NVL(mrs�ֶ�!Ʊ�ݷ�Χ) & "-", "-")
        If varTemp(1) = "" Then varTemp(1) = varTemp(0)
        txtEdit(2).Text = Mid(varTemp(0), intǰ׺ + 1)
        txtEdit(4).Text = Mid(varTemp(1), intǰ׺ + 1)
    Else
        txtEdit(2).Text = "": txtEdit(4).Text = ""
    End If
    '103428:���ϴ���2017/2/15��ҽ�ƿ�Ҫ�������õ�Ʊ�ݳ���ȷ�����볤��
    If intƱ�� = gBillType.���￨ Or intƱ�� = gBillType.���ѿ� Then
        mlng���� = zlStr.ActualLen(mstr��⿪ʼ��)
        Call txtEdit_Change(1)
    End If
    
    txtEdit(2).Tag = txtEdit(2).Text: txtEdit(4).Tag = txtEdit(4).Text
    zlCommFun.PressKey vbKeyTab
    rsTemp.Close
    Select���Ʊ�� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Load�ֶ�Ʊ��(ByVal intƱ�� As gBillType, ByVal str���� As String, _
    ByVal str��⿪ʼ�� As String, ByVal str�����ֹ�� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���طֶ�Ʊ������
    '����:���˺�
    '����:2010-11-18 17:27:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strKey As String, str��ʼ���� As String, intǰ׺ As Integer, lng���� As Long
    
    On Error GoTo errHandle
    Call Init�ֶ�Ʊ��(mrs�ֶ�)
    intǰ׺ = Len(txtEdit(1).Text): lng���� = Len(str��⿪ʼ��) - intǰ׺
    '��ȡ��ǰ���ε�����ź���С���
    If intƱ�� = gBillType.���ѿ� Then
        gstrSQL = _
            "Select ��ʼ���� As ��ʼ����,nvl(��ֹ����,��ʼ����) as ��ֹ���� From ���ѿ������¼ Where ���ID=[3]" & vbNewLine & _
            "Union ALL " & _
            "Select ��ʼ���� As ��ʼ����,nvl(��ֹ����,��ʼ����) as ��ֹ����" & vbNewLine & _
            "From ���ѿ����ü�¼" & vbNewLine & _
            "Where ����=[1] And �ӿڱ��=(Select Max(�ӿڱ��) From ���ѿ�����¼ Where id=[3])" & vbNewLine & _
                    IIf(mlng����ID <> 0, " And ID<>[2] ", "") & _
            "Order by ��ʼ����"
    Else
        gstrSQL = _
            "Select ��ʼ����,nvl(��ֹ����,��ʼ����) as ��ֹ���� From Ʊ�ݱ����¼ Where ���ID=[3]" & vbNewLine & _
            "Union ALL " & _
            "Select ��ʼ����,nvl(��ֹ����,��ʼ����) as ��ֹ����" & vbNewLine & _
            "From Ʊ�����ü�¼" & vbNewLine & _
            "Where ����=[1] And Ʊ��=(Select Max(Ʊ��) From Ʊ������¼ Where id=[3])" & vbNewLine & _
                    IIf(mlng����ID <> 0, " And ID<>[2] ", "") & _
            "Order by ��ʼ����"
    End If
    Set mrs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����, mlng����ID, mlng���ID)
    If Not mrs����.EOF Then
        i = 1
        str��ʼ���� = str��⿪ʼ��
        Do While Not mrs����.EOF
            If str��ʼ���� < NVL(mrs����!��ʼ����) Then
                strKey = txtEdit(1).Text & _
                    zlStr.LPAD(zlStr.Increase(Mid(NVL(mrs����!��ʼ����), intǰ׺ + 1), True), lng����, "0", True)
                If strKey <> str��ʼ���� Then
                    strKey = str��ʼ���� & "-" & strKey
                End If
                mrs�ֶ�.AddNew
                mrs�ֶ�!ID = i
                mrs�ֶ�!��� = i
                mrs�ֶ�!Ʊ�ݷ�Χ = strKey
                mrs�ֶ�.Update
                i = i + 1
            End If
            str��ʼ���� = txtEdit(1).Text & _
                zlStr.LPAD(zlStr.Increase(Mid(NVL(mrs����!��ֹ����), intǰ׺ + 1), False), lng����, "0", True)
            mrs����.MoveNext
        Loop
        strKey = str�����ֹ��
        If str��ʼ���� <= strKey And str��ʼ���� <> "" Then
            If str��ʼ���� <> strKey Then
                strKey = str��ʼ���� & "-" & strKey
            End If
            mrs�ֶ�.AddNew
            mrs�ֶ�!ID = i
            mrs�ֶ�!��� = i
            mrs�ֶ�!Ʊ�ݷ�Χ = strKey
            mrs�ֶ�.Update
        End If
    Else
        mrs�ֶ�.AddNew
        mrs�ֶ�!ID = 1
        mrs�ֶ�!��� = 1
        mrs�ֶ�!Ʊ�ݷ�Χ = str��⿪ʼ�� & "-" & str�����ֹ��
        mrs�ֶ�.Update
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Init�ֶ�Ʊ��(rs�ֶ� As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�ֶ�Ʊ�ŵ����ݽṹ
    '����:���˺�
    '����:2010-11-18 14:33:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rs�ֶ� = New ADODB.Recordset
    With rs�ֶ�
        If .State = adStateOpen Then .Close
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "Ʊ�ݷ�Χ", adLongVarChar, 200, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub

Private Sub InitContext()
    Dim dtCurrnet As Date
    
    mblnҩ�� = (glngSys \ 100 = 8)
    
    mstr��С���� = ""
    mstr������ = ""
    
    dtCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpDate.Value = dtCurrnet
    dtpDate.MaxDate = dtCurrnet
    
    cmbƱ��.Clear
    If mblnҩ�� = False Then
        If zlStr.IsHavePrivs(mstrPrivs, "�շ��վ�") Then
            cmbƱ��.AddItem "1-�շ��վ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 1
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "Ԥ���վ�") _
          Or (zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
          Or zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��")) Then
            cmbƱ��.AddItem "2-Ԥ���վ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 2
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "�����վ�") Then
          cmbƱ��.AddItem "3-�����վ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 3
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "�Һ��վ�") Then
          cmbƱ��.AddItem "4-�Һ��վ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 4
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "ҽ�ƿ�") Then
           cmbƱ��.AddItem "5-ҽ�ƿ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 5
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "���ѿ�") Then
           cmbƱ��.AddItem "6-���ѿ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 6
        End If
'        cmbƱ��.AddItem "1-�շ��վ�":        cmbƱ��.ItemData(cmbƱ��.NewIndex) = 1
'        cmbƱ��.AddItem "2-Ԥ���վ�":        cmbƱ��.ItemData(cmbƱ��.NewIndex) = 2
'        cmbƱ��.AddItem "3-�����վ�":        cmbƱ��.ItemData(cmbƱ��.NewIndex) = 3
'        cmbƱ��.AddItem "4-�Һ��վ�":        cmbƱ��.ItemData(cmbƱ��.NewIndex) = 4
'        cmbƱ��.AddItem "5-ҽ�ƿ�":          cmbƱ��.ItemData(cmbƱ��.NewIndex) = 5
'        cmbƱ��.AddItem "6-���ѿ�":          cmbƱ��.ItemData(cmbƱ��.NewIndex) = 6
    Else
        cmbƱ��.AddItem "1-�շ��վ�": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 1
        cmbƱ��.AddItem "5-��Ա��": cmbƱ��.ItemData(cmbƱ��.NewIndex) = 5
    End If
    
    cmbʹ�÷�ʽ.Clear
    cmbʹ�÷�ʽ.AddItem "1-����"
    cmbʹ�÷�ʽ.AddItem "2-����"
    cmbʹ�÷�ʽ.ListIndex = 0
    
    '��ʼ��Ʊ�ݴ�ӡ
    'On Error Resume Next
    'BillInit gcnOracle
End Sub

Private Sub cbo���_Click()
    Dim blnChange As Boolean
    
    If mintƱ�� = gBillType.�Һ��վ� Then Exit Sub
    mblnChange = True
    If mintƱ�� = gBillType.�շ��վ� Or mintƱ�� = gBillType.�����վ� Then
        If cbo���.Tag = Trim(cbo���.Text) Then Exit Sub
        cbo���.Tag = Trim(cbo���.Text)
    Else
        If Val(cbo���.Tag) = cbo���.ItemData(cbo���.ListIndex) Then Exit Sub
        cbo���.Tag = cbo���.ItemData(cbo���.ListIndex)
    End If
    
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        If cbo���.ListIndex >= 0 Then
            mlng���� = mcllCardProperty(cbo���.ListIndex + 1)(0)
            If mlng���� = 1 Or mlng���� = 2 Then
                txtEdit(1).Text = ""
            End If
            Call txtEdit_Change(1)
        End If
    End If
    If mblnNotClick Then GoTo hdYLK

    If mbytInFun = 0 And mlng����ID = 0 Then
        txtEdit(5).Text = ""
        txtEdit(1).Text = ""
        txtEdit(2).Text = ""
        txtEdit(3).Text = ""
        txtEdit(4).Text = ""
    End If
hdYLK:
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        txtEdit(1).Text = UCase(mcllCardProperty(cbo���.ListIndex + 1)(1))
        txtEdit(1).Enabled = mcllCardProperty(cbo���.ListIndex + 1)(1) = "" And Not mbln���ȷ������ And mlng���� > 2
        txtEdit(3).Enabled = txtEdit(1).Enabled
        txtEdit(1).BackColor = IIf(txtEdit(1).Enabled, txtEdit(2).BackColor, cmdOK.BackColor)
        txtEdit(3).BackColor = txtEdit(1).BackColor
    End If
    cmdOK.Enabled = True
End Sub

Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmb������_Validate(Cancel As Boolean)
    If cmb������.ListIndex < 0 Then zlControl.CboLocate cmb������, mlngPreID, True
    If cmb������.ListIndex < 0 And cmb������.Text <> "" Then cmb������.Text = ""
End Sub

Private Sub cmbƱ��_Click()
    'ѡ����Ӧ����Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strWhere As String
    
    On Error GoTo errHandle
    '115348:���ϴ�,2017/10/24,Ʊ��δ�ı䲻ˢ�½�����Ϣ
    If Val(cmbƱ��.Tag) = cmbƱ��.ItemData(cmbƱ��.ListIndex) Then Exit Sub
    cmbƱ��.Tag = cmbƱ��.ItemData(cmbƱ��.ListIndex)
    cbo���.Tag = ""
    mblnChange = True
    
    mintƱ�� = cmbƱ��.ItemData(cmbƱ��.ListIndex)
    mblnIsBIll = CurrentIsBill(mintƱ��)
    If mblnIsBIll Then
        lblTitle.Caption = "Ʊ�����õ�"
        lbl(6).Caption = "���뷶Χ(&B)"
        lblUserType.Caption = "ʹ�����(&K)"
    Else
        lblTitle.Caption = IIf(mintƱ�� = gBillType.���￨, "ҽ�ƿ����õ�", "���ѿ����õ�")
        lbl(6).Caption = "���ŷ�Χ(&B)"
        lblUserType.Caption = "�����(&K)"
    End If
    
    Call LoadCombox
    If mintƱ�� = gBillType.Ԥ���վ� Or mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        If cbo���.ListIndex < 0 Then
            mstrPreType(mintƱ��) = ""
        Else
            mstrPreType(mintƱ��) = cbo���.ItemData(cbo���.ListIndex)
        End If
    Else
        mstrPreType(mintƱ��) = cbo���.Text
    End If
    '�õ���ǰƱ������ĳ���
'    mlng���� = Val(Mid(mstrƱ�ݳ���, mmintƱ��, 1))
'    If mlng���� = 0 Then
'        mlng���� = 10
'    End If
    If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
        If cbo���.ListIndex >= 0 Then
            mlng���� = mcllCardProperty(cbo���.ListIndex + 1)(0)
        End If
    Else
        mlng���� = Val(Split(mstrƱ�ݳ���, "|")(mintƱ�� - 1))
    End If
    If mlng���� = 1 Or mlng���� = 2 Then
        '������ǰ׺
        txtEdit(1).Enabled = False
        txtEdit(1).Text = ""
        txtEdit(3).Enabled = False
        txtEdit(3).Text = ""
    Else
        txtEdit(1).Enabled = True
        txtEdit(3).Enabled = True
        If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
            txtEdit(1).Enabled = mcllCardProperty(cbo���.ListIndex + 1)(1) = ""
            txtEdit(3).Enabled = txtEdit(1).Enabled
        End If
    End If
    Call txtEdit_Change(1)
    
    Select Case mintƱ��
        Case gBillType.�շ��վ�      '1-�շ��վ�
            strWhere = " And B.��Ա����='�����շ�Ա'"
        Case gBillType.Ԥ���վ�      '2-Ԥ���վ�
            strWhere = " And B.��Ա���� in ('Ԥ���տ�Ա','��Ժ�Ǽ�Ա')"
        Case gBillType.�����վ�      '3-�����վ�
            strWhere = " And B.��Ա����='סԺ����Ա'"
        Case gBillType.�Һ��վ�      '4-�Һ��վ�
            strWhere = " And B.��Ա����='����Һ�Ա'"
        Case gBillType.���￨, gBillType.���ѿ�     '5-ҽ�ƿ� ���߳�Ϊ  ��Ա��
            strWhere = " And B.��Ա���� in ('�����Ǽ���','��Ժ�Ǽ�Ա')"
        Case Else
            Exit Sub
    End Select
    gstrSQL = _
        "Select distinct A.ID,A.���, A.����,A.����" & vbNewLine & _
        "From ��Ա�� A,��Ա����˵�� B" & vbNewLine & _
        "Where A.ID=B.��ԱID " & strWhere & vbNewLine & _
        "      And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        "      And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
        "Order By A.����"
    Set mrsPerson = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    cmb������.Clear
    Do Until mrsPerson.EOF
        cmb������.AddItem mrsPerson("����")
        cmb������.ItemData(cmb������.NewIndex) = Val(NVL(mrsPerson!ID))
        mrsPerson.MoveNext
    Loop
    If cmb������.ListCount > 0 Then cmb������.ListIndex = 0
    
    With cmbƱ��
        If mint�ϴ�Ʊ�� <> .ItemData(.ListIndex) Then
            If mintƱ�� = gBillType.���ѿ� Then
                gstrSQL = "Select 1 From ���ѿ�����¼ Where Rownum < 2"
            Else
                gstrSQL = "Select 1 From Ʊ������¼ Where Ʊ��=[1] And Rownum < 2"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .ItemData(.ListIndex))
            mbln���ȷ������ = Not rsTmp.EOF
            If mbln���ȷ������ Then
                txtEdit(1).Text = "": txtEdit(2).Text = "":
                txtEdit(3).Text = "": txtEdit(4).Text = "":
                txtEdit(5).Text = ""
                If mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
                    txtEdit(1).Text = UCase(mcllCardProperty(cbo���.ListIndex + 1)(1))
                    txtEdit(3).Text = UCase(mcllCardProperty(cbo���.ListIndex + 1)(1))
                End If
            End If
            mint�ϴ�Ʊ�� = .ItemData(.ListIndex)
        End If
    End With
    Call SetCtrlEnable
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmbƱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmb������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    
     '���˺� ����:27378 ����:2010-01-27 16:20:02
    Dim strAllCaption As String
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cmb������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
       If mrsPerson Is Nothing Then Exit Sub
    If zlPersonSelect(Me, mlngModule, cmb������, mrsPerson, cmb������.Text, True, strAllCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
 
End Sub

Private Sub cmb������_Click()
    If cmb������.ListIndex >= 0 Then mlngPreID = cmb������.ItemData(cmb������.ListIndex)
    mblnChange = True
    cmdOK.Enabled = True
End Sub

Private Sub cmbʹ�÷�ʽ_Click()
    mblnChange = True
    cmdOK.Enabled = True
End Sub

Private Sub cmbʹ�÷�ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdSel_Click()
    If SelectƱ�� = False Then Exit Sub
    zlControl.ControlSetFocus cmb������
    cmdOK.Enabled = True
End Sub
Private Sub cmd����_Click()
    If Select���Ʊ��(mintƱ��, txtEdit(5), "") = False Then
        Exit Sub
    End If
    cmdOK.Enabled = True
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mint�ϴ�Ʊ�� = -1
    mbln���ȷ������ = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lng����ID As Long
    Dim strUserName As String
    
    If ValidateContent() = False Then Exit Sub
    If mbytInFun = 0 Then '104831
        If Val(zlDatabase.GetPara("����" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "ǩ��ȷ��", glngSys, mlngModule, 0)) = 1 Then
            '����:40775
            strUserName = zlDatabase.UserIdentify(Me, "������ǩ���û��������룡", glngSys, mlngModule, "")
            If strUserName = "" Then Exit Sub
            If strUserName <> cmb������.Text Then
                MsgBox "��������ǩ���˲�һ�£����ܼ�����", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If Save(lng����ID, strUserName) = False Then Exit Sub
    
    mblnOK = True
    If mbytInFun = 0 Then
        If mlng����ID <> 0 Then
            '�޸�
            mblnChange = False
            Unload Me
            Exit Sub
        Else
            '��������
            txtEdit(2).Text = ""
            txtEdit(4).Text = ""
            '�����:115671,����,2017/11/15,����Ʊ�ݺ�ȷ����ť����,ʹ���ͣ����Ʊ�ֵ��������С�
            cmdOK.Enabled = False
            If cmbƱ��.Enabled And cmbƱ��.Visible Then cmbƱ��.SetFocus
            If mstr��⿪ʼ�� <> "" Then
                Call Load�ֶ�Ʊ��(mintƱ��, Trim(txtEdit(5).Text), mstr��⿪ʼ��, mstr��������)
                If mrs�ֶ�.RecordCount <> 0 Then
                    Dim varTemp As Variant
                    mrs�ֶ�.MoveFirst
                    varTemp = Split(NVL(mrs�ֶ�!Ʊ�ݷ�Χ) & "-", "-")
                    If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                    txtEdit(2).Text = Mid(varTemp(0), Len(txtEdit(1).Text) + 1)
                    txtEdit(4).Text = Mid(varTemp(1), Len(txtEdit(1).Text) + 1)
                Else
                    txtEdit(5).Text = "": zlControl.ControlSetFocus txtEdit(5)
                End If
            End If
        End If
        mblnChange = False
    Else
        mblnChange = False
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub optResult_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And txtEdit(1).Text <> txtEdit(3).Text Then txtEdit(3).Text = txtEdit(1).Text
    If Index = 3 And txtEdit(1).Text <> txtEdit(3).Text Then txtEdit(1).Text = txtEdit(3).Text
    If Index = 1 Or Index = 3 Then
         
        txtEdit(2).MaxLength = mlng���� - LenB(StrConv(txtEdit(1).Text, vbFromUnicode))
        txtEdit(4).MaxLength = txtEdit(2).MaxLength
    End If
    If Index = 5 Then
        txtEdit(Index).Tag = "": Set mrs�ֶ� = Nothing
        mlng���ID = 0
    End If
    Call ShowSum
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 1 Or Index = 3 Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
End Sub

Private Function SelectƱ��() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ������Ʊ�ݺ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-11-18 16:24:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����ѡ����
    Dim rsReturn As ADODB.Recordset, varTemp As Variant
    If Not mbln���ȷ������ Then
        SelectƱ�� = True: Exit Function
    End If
    
    On Error GoTo errHandle
    If mrs�ֶ� Is Nothing Then
        ShowMsgbox "����ȷ��������Σ����飡"
        zlControl.ControlSetFocus txtEdit(5)
        Exit Function
    End If
    
    mrs�ֶ�.Filter = 0
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtEdit(2), mrs�ֶ�, True, "", "ID", rsReturn) Then
        If rsReturn.RecordCount <> 0 Then
            varTemp = Split(rsReturn!Ʊ�ݷ�Χ & "-", "-")
            If varTemp(1) = "" Then varTemp(1) = varTemp(0)
            txtEdit(2).Text = Mid(varTemp(0), Len(txtEdit(1).Text) + 1)
            txtEdit(4).Text = Mid(varTemp(1), Len(txtEdit(3).Text) + 1)
            zlControl.ControlSetFocus cmb������
        End If
    End If
    mrs�ֶ�.Filter = 0
    
    SelectƱ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 5 And mbln���ȷ������ Then
            If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
            If Select���Ʊ��(mintƱ��, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
                Exit Sub
            End If
            Exit Sub
        End If
        If Not (Index = 2 Or Index = 4) Then
            If Trim(txtEdit(Index)) = "" Then
                If SelectƱ�� = False Then Exit Sub
            End If
        End If
        zlCommFun.PressKey vbKeyTab: Exit Sub
    Else
        
    End If
    If Index = 1 Or Index = 3 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
    Else
        If Not (Index = 5 And mbln���ȷ������) Then
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Function ValidateContent() As Boolean
'����:����������ݵ��Ƿ���Ч
'����:��Ч�򷵻�True,���򷵻�False
    Dim i As Integer, strTemp As String
    Dim str��� As String, strNote As String
    Dim rsTmp As New ADODB.Recordset
    Dim bln�������� As Boolean '�����:43366
    Dim lng���� As Long, byt�������� As Byte
    Dim strName As String
    
    On Error GoTo errHandle
    strName = IIf(mblnIsBIll, "����", "����")
    ValidateContent = False
    If mbytInFun = 0 Then
        If cmbƱ��.ListIndex < 0 Then
            MsgBox "����ѡ��Ҫ���õ�Ʊ�֡�", vbExclamation, gstrSysName
            If cmbƱ��.Visible And cmbƱ��.Enabled Then cmbƱ��.SetFocus
            Exit Function
        End If
        '�ַ������
        For i = 1 To 4
            If zlCommFun.StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength) = False Then
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        Next
        
        For i = 1 To Len(txtEdit(2).Text)
            strTemp = Mid(txtEdit(2), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "��ʼ" & strName & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                zlControl.TxtSelAll txtEdit(2)
                Exit Function
            End If
        Next
        For i = 1 To Len(txtEdit(4).Text)
            strTemp = Mid(txtEdit(4), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "��ֹ" & strName & "�к��з������ַ�����ĸֻ����Ϊǰ׺��", vbExclamation, gstrSysName
                txtEdit(4).SetFocus
                zlControl.TxtSelAll txtEdit(4)
                Exit Function
            End If
        Next
        If mbln���ȷ������ Then
            If txtEdit(5).Tag = "" Then
                    MsgBox "�������δѡ��,�������á�", vbExclamation, gstrSysName
                    zlControl.ControlSetFocus txtEdit(5)
                    Exit Function
            End If
        End If
        If Len(txtEdit(2).Text) <> txtEdit(2).MaxLength Then
            If Not mbln���ȷ������ And (mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ�) Then
                lng���� = mcllCardProperty(cbo���.ListIndex + 1)(0)
                byt�������� = mcllCardProperty(cbo���.ListIndex + 1)(3)
                Select Case byt��������
                    Case 0
                        MsgBox "��ʼ" & strName & "�ĳ��Ȳ�����Ӧ����" & lng���� & "λ!", vbExclamation, gstrSysName
                        txtEdit(2).SetFocus
                        zlControl.TxtSelAll txtEdit(2)
                        Exit Function
                    Case 2
                        If MsgBox("��ʼ" & strName & "�ĳ�������" & lng���� & "λ,�Ƿ������", vbExclamation + vbYesNo, gstrSysName) = vbNo Then
                            txtEdit(2).SetFocus
                            zlControl.TxtSelAll txtEdit(2)
                            Exit Function
                        End If
                End Select
            Else
                MsgBox "��ʼ" & strName & "�ĳ��Ȳ�����Ӧ����" & mlng���� & "λ��", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                zlControl.TxtSelAll txtEdit(2)
                Exit Function
            End If
        End If
        If Len(txtEdit(2).Text) = 0 Then
            MsgBox "��ʼ" & strName & "����Ϊ�ա�", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
        If Len(txtEdit(2).Text) <> Len(txtEdit(4).Text) Then
            MsgBox "��ֹ" & strName & "�ĳ���Ҫ�Ϳ�ʼ" & strName & "����ͬ��", vbExclamation, gstrSysName
            txtEdit(4).SetFocus
            zlControl.TxtSelAll txtEdit(4)
            Exit Function
        End If
        If txtEdit(2).Text > txtEdit(4).Text Then
            MsgBox "��ʼ" & strName & "����С����ֹ" & strName & "��", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
        If txtEdit(2).Text = "0000000000" And txtEdit(4).Text = "9999999999" Then
            MsgBox "����ʹ�����" & strName & "��Χ��", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
        If mstr��С���� <> "" Then
            If Len(txtEdit(2).Text) <> Len(txtEdit(2).Tag) Then
                MsgBox "�������õ���" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "�Ѿ�ʹ�ã�" & strName & "���Ȳ��ܸı䡣" & vbCrLf & _
                    strName & "����Ӧ����" & Len(txtEdit(1).Text & txtEdit(2).Tag) & "λ��", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                zlControl.TxtSelAll txtEdit(2)
                Exit Function
            End If
            If txtEdit(1).Text & txtEdit(2).Text > mstr��С���� Then
                MsgBox "�������õ���" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "�Ѿ�ʹ�ã�" & vbCrLf & _
                    "��ʼ" & strName & "���ֻ���Ե�" & mstr��С���� & "��", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                zlControl.TxtSelAll txtEdit(2)
                Exit Function
            End If
            If txtEdit(3).Text & txtEdit(4).Text < mstr������ Then
                MsgBox "�������õ���" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "�Ѿ�ʹ�ã�" & vbCrLf & _
                    strName & "�Ѿ��õ�" & mstr������ & "����ֹ" & strName & "�����������", vbExclamation, gstrSysName
                txtEdit(2).SetFocus
                zlControl.TxtSelAll txtEdit(2)
                Exit Function
            End If
        End If
        If cmb������.Text = "" Then
            MsgBox "�����˲���Ϊ�ա�", vbExclamation, gstrSysName
            cmb������.SetFocus
            Exit Function
        End If
        
        '�����:43366,54259
        If Len(CalcTotal) > 11 Then
            bln�������� = True
        ElseIf Len(CalcTotal) < 11 Then
            bln�������� = False
        ElseIf CalcTotal > "9999999999" Then
            bln�������� = True
        ElseIf CalcTotal < "9999999999" Then
            bln�������� = False
        End If
        
        '�������������Ƿ����
        If bln�������� Then
            MsgBox strName & "���õ������쳣�������鿪ʼ����" & strName & "����ȷ�ԡ�", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
        
        
'        '�����뷶Χ�Ƿ����
'        If CalcTotal > 999999999# Then
'            MsgBox strName & "���õ������쳣�������鿪ʼ����" & strName & "����ȷ�ԡ�", vbExclamation, gstrSysName
'            txtEdit(2).SetFocus
'            SelAll txtEdit(2)
'            Exit Function
'        End If
        '����Ƿ���ʹ�����
        
        If mintƱ�� = gBillType.�շ��վ� Or mintƱ�� = gBillType.Ԥ���վ� _
            Or mintƱ�� = gBillType.�����վ� Or mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
            If cbo���.ListIndex < 0 Then
                MsgBox "ע��:" & vbCrLf & IIf(mintƱ�� = gBillType.Ԥ���վ�, "   Ԥ�����", "    ʹ�����") & "û��ѡ����ѡ��", vbInformation, gstrSysName
                zlControl.ControlSetFocus cbo���: Exit Function
                Exit Function
            End If
            If mintƱ�� = gBillType.Ԥ���վ� Or mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
                str��� = cbo���.ItemData(cbo���.ListIndex)
            End If
        End If
        '�ж������Ƿ��ظ�
        If mintƱ�� = gBillType.���ѿ� Then
            gstrSQL = _
                "Select ������, �Ǽ�ʱ��, ��ʼ���� As ��ʼ����, ��ֹ���� As ��ֹ����, Nvl(ʣ������,0) ʣ������ " & _
                "From ���ѿ����ü�¼ " & _
                "Where ID<>[3] And �ӿڱ�� =[6] " & _
                "      And (��ʼ����<=[1] and ��ֹ����>=[1] or ��ʼ����<=[2] and ��ֹ����>=[2]) And length(��ʼ����)=length([1]) And Nvl(����,'0')=[7]"
        Else
            gstrSQL = _
                "Select ������, �Ǽ�ʱ��, ��ʼ����, ��ֹ����, Nvl(ʣ������,0) ʣ������ " & _
                "From Ʊ�����ü�¼ " & _
                "Where ID<>[3] And Ʊ��=[4]" & _
                        IIf(mintƱ�� = gBillType.�շ��վ� Or mintƱ�� = gBillType.�����վ�, " and nvl(ʹ�����,'LXH')=[5]", _
                        IIf(mintƱ�� = gBillType.Ԥ���վ� Or mintƱ�� = gBillType.���￨, " And nvl(ʹ�����,'2') =[6]", "")) & _
                "      And (��ʼ����<=[1] and ��ֹ����>=[1] or ��ʼ����<=[2] and ��ֹ����>=[2]) And length(��ʼ����)=length([1]) And Nvl(����,'0')=[7]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtEdit(1).Text & txtEdit(2).Text, _
            txtEdit(3).Text & txtEdit(4).Text, mlng����ID, Left(cmbƱ��.Text, 1), _
            IIf(Trim(cbo���.Text) = "", "LXH", cbo���.Text), str���, NVL(txtEdit(5).Text, "0"))
        If rsTmp.RecordCount > 0 Then
            strNote = "�뱾������" & IIf(mblnIsBIll, "Ʊ��", "��") & "�����ص������ü�¼����,��������" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & ",�ص������ü�¼����:" & vbCrLf
            Do While Not rsTmp.EOF
                strNote = strNote & rsTmp!������ & "��" & Format(rsTmp!�Ǽ�ʱ��, "yyyy-mm-dd") & "������" & rsTmp!��ʼ���� & "��" & rsTmp!��ֹ���� & "��" & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "." & vbCrLf
                rsTmp.MoveNext
            Loop
            MsgBox strNote, vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If zlCommFun.ActualLen(txtRemarks.Text) > txtRemarks.MaxLength Then
            MsgBox "��ע��Ϣ��������" & txtRemarks.MaxLength & "���ַ�!", vbExclamation, gstrSysName
            If txtRemarks.Enabled Then txtRemarks.SetFocus
            Exit Function
        End If
    End If
    ValidateContent = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Save(ByRef lng����ID As Long, ByVal strUserName As String) As Boolean
'����:����༭������
'����:lng����ID-����ʱ�����¼�¼������ID
'����ֵ:�ɹ�����True,����ΪFalse
    Dim strTemp As String, strSQL As String, str��� As String
    
    On Error GoTo errHandle
    str��� = ""
    Select Case mintƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        str��� = Trim(cbo���.Text)
    Case gBillType.Ԥ���վ�
        str��� = cbo���.ItemData(cbo���.ListIndex)
        If Val(str���) = 0 Then str��� = ""
    Case gBillType.���￨, gBillType.���ѿ�
        str��� = cbo���.ItemData(cbo���.ListIndex)
        If Val(str���) = 0 Then str��� = ""
    End Select
    
    str��� = IIf(str��� = "", "NULL", "'" & str��� & "'")
    If mintƱ�� = gBillType.���ѿ� Then
        If mbytInFun = 0 Then
            If mlng����ID = 0 Then '����
                lng����ID = zlDatabase.GetNextId("���ѿ����ü�¼")
                'Zl_���ѿ����ü�¼_Insert
                strSQL = "Zl_���ѿ����ü�¼_Insert("
                '  Id_In       ���ѿ����ü�¼.Id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  �ӿڱ��_In ���ѿ����ü�¼. �ӿڱ��%Type,
                strSQL = strSQL & "" & str��� & ","
                '  ������_In   ���ѿ����ü�¼.������%Type,
                strSQL = strSQL & "'" & cmb������.Text & "',"
                '  ǰ׺�ı�_In ���ѿ����ü�¼.ǰ׺�ı�%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & "',"
                '  ��ʼ����_In ���ѿ����ü�¼.��ʼ����%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & txtEdit(2).Text & "',"
                '  ��ֹ����_In ���ѿ����ü�¼.��ֹ����%Type,
                strSQL = strSQL & "'" & txtEdit(3).Text & txtEdit(4).Text & "',"
                '  ʹ�÷�ʽ_In ���ѿ����ü�¼.ʹ�÷�ʽ%Type,
                strSQL = strSQL & "'" & Left(cmbʹ�÷�ʽ.Text, 1) & "',"
                '  �Ǽ�ʱ��_In ���ѿ����ü�¼.�Ǽ�ʱ��%Type := Null,
                strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
                '  �Ǽ���_In   ���ѿ����ü�¼.�Ǽ���%Type := Null,
                strSQL = strSQL & "'" & txtEdit(0).Text & "',"
                '  ʣ������_In ���ѿ����ü�¼.ʣ������%Type := Null,
                strSQL = strSQL & "" & CalcTotal & ","
                '  ����_In     ���ѿ����ü�¼.����%Type := Null,
                strSQL = strSQL & "'" & txtEdit(5).Text & "',"
                '  ǩ����_In   ���ѿ����ü�¼.ǩ����%Type := Null
                strSQL = strSQL & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ","
                '  ���id_In   ���ѿ����ü�¼.���id%Type := Null
                strSQL = strSQL & IIf(mlng���ID = 0, "NULL", mlng���ID) & ")"
            Else '�޸�
                'Zl_���ѿ����ü�¼_Update
                strSQL = "Zl_���ѿ����ü�¼_Update("
                '  Id_In       ���ѿ����ü�¼.Id%Type,
                strSQL = strSQL & "" & mlng����ID & ","
                '  �ӿڱ��_In ���ѿ����ü�¼.�ӿڱ��%Type,
                strSQL = strSQL & "" & str��� & ","
                '  ������_In   ���ѿ����ü�¼.������%Type,
                strSQL = strSQL & "'" & cmb������.Text & "',"
                '  ��ʼ����_In ���ѿ����ü�¼.��ʼ����%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & txtEdit(2).Text & "',"
                '  ��ֹ����_In ���ѿ����ü�¼.��ֹ����%Type,
                strSQL = strSQL & "'" & txtEdit(3).Text & txtEdit(4).Text & "',"
                '  ǰ׺�ı�_In ���ѿ����ü�¼.ǰ׺�ı�%Type := Null,
                strSQL = strSQL & "'" & txtEdit(1).Text & "',"
                '  ʹ�÷�ʽ_In ���ѿ����ü�¼.ʹ�÷�ʽ%Type := 1,
                strSQL = strSQL & "'" & Left(cmbʹ�÷�ʽ.Text, 1) & "',"
                '  �Ǽ�ʱ��_In ���ѿ����ü�¼.�Ǽ�ʱ��%Type := Null,
                strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
                '  �Ǽ���_In   ���ѿ����ü�¼.�Ǽ���%Type := Null,
                strSQL = strSQL & "'" & txtEdit(0).Text & "',"
                '  ����_In     ���ѿ����ü�¼.����%Type := Null,
                strSQL = strSQL & "'" & txtEdit(5).Text & "',"
                '  ǩ����_In   ���ѿ����ü�¼.ǩ����%Type := Null
                strSQL = strSQL & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ","
                '  ���id_In   ���ѿ����ü�¼.���id%Type := Null
                strSQL = strSQL & IIf(mlng���ID = 0, "NULL", mlng���ID) & ")"
            End If
        Else '�˶����õ�
            ' Zl_���ѿ����ü�¼_Check
            strSQL = " Zl_���ѿ����ü�¼_Check("
            '  Id_In       ���ѿ����ü�¼.Id%Type,
            strSQL = strSQL & "" & mlng����ID & ","
            '  �˶Խ��_In ���ѿ����ü�¼.�˶Խ��%Type,
            strSQL = strSQL & "" & IIf(optResult(0).Value, 1, 0) & ","
            '  �˶���_In   ���ѿ����ü�¼.�˶���%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ��ע_In     ���ѿ����ü�¼.��ע%Type,
            strSQL = strSQL & "'" & txtRemarks.Text & "',"
            '  �˶�ģʽ_In ���ѿ����ü�¼.�˶�ģʽ%Type
            strSQL = strSQL & "" & 0 & ")"
        End If
    Else
        If mbytInFun = 0 Then
            If mlng����ID = 0 Then
                '����
                lng����ID = zlDatabase.GetNextId("Ʊ�����ü�¼")
                'Zl_Ʊ�����ü�¼_Insert
                strSQL = "Zl_Ʊ�����ü�¼_Insert("
                '  Id_In       In Ʊ�����ü�¼.ID%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  Ʊ��_In     In Ʊ�����ü�¼.Ʊ��%Type,
                strSQL = strSQL & "" & Left(cmbƱ��.Text, 1) & ","
                '  ʹ�����_In In Ʊ�����ü�¼.ʹ�����%Type,
                strSQL = strSQL & "" & str��� & ","
                '  ������_In   In Ʊ�����ü�¼.������%Type,
                strSQL = strSQL & "'" & cmb������.Text & "',"
                '  ǰ׺�ı�_In In Ʊ�����ü�¼.ǰ׺�ı�%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & "',"
                '  ��ʼ����_In In Ʊ�����ü�¼.��ʼ����%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & txtEdit(2).Text & "',"
                '  ��ֹ����_In In Ʊ�����ü�¼.��ֹ����%Type,
                strSQL = strSQL & "'" & txtEdit(3).Text & txtEdit(4).Text & "',"
                '  ʹ�÷�ʽ_In In Ʊ�����ü�¼.ʹ�÷�ʽ%Type,
                strSQL = strSQL & "'" & Left(cmbʹ�÷�ʽ.Text, 1) & "',"
                '  �Ǽ�ʱ��_In In Ʊ�����ü�¼.�Ǽ�ʱ��%Type := Null,
                strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
                '  �Ǽ���_In   In Ʊ�����ü�¼.�Ǽ���%Type := Null,
                strSQL = strSQL & "'" & txtEdit(0).Text & "',"
                '  ʣ������_In In Ʊ�����ü�¼.ʣ������%Type := Null,
                strSQL = strSQL & "" & CalcTotal & ","
                '  ����_In     In Ʊ�����ü�¼.����%Type := Null,
                strSQL = strSQL & "'" & txtEdit(5).Text & "',"
                '  ǩ����_In   In Ʊ�����ü�¼.ǩ����%Type := Null
                strSQL = strSQL & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ","
                 '  ���id_In   In Ʊ�����ü�¼.���id%Type := Null
                strSQL = strSQL & IIf(mlng���ID = 0, "NULL", mlng���ID) & ")"
            Else
                '�޸�
                'Zl_Ʊ�����ü�¼_Update
                strSQL = "Zl_Ʊ�����ü�¼_Update("
                '  Id_In       In Ʊ�����ü�¼.ID%Type,
                strSQL = strSQL & "" & mlng����ID & ","
                '  ʹ�����_In In Ʊ�����ü�¼.ʹ�����%Type,
                strSQL = strSQL & "" & str��� & ","
                '  ������_In   In Ʊ�����ü�¼.������%Type,
                strSQL = strSQL & "'" & cmb������.Text & "',"
                '  ��ʼ����_In In Ʊ�����ü�¼.��ʼ����%Type,
                strSQL = strSQL & "'" & txtEdit(1).Text & txtEdit(2).Text & "',"
                '  ��ֹ����_In In Ʊ�����ü�¼.��ֹ����%Type,
                strSQL = strSQL & "'" & txtEdit(3).Text & txtEdit(4).Text & "',"
                '  ǰ׺�ı�_In In Ʊ�����ü�¼.ǰ׺�ı�%Type := Null,
                strSQL = strSQL & "'" & txtEdit(1).Text & "',"
                '  ʹ�÷�ʽ_In In Ʊ�����ü�¼.ʹ�÷�ʽ%Type := 1,
                strSQL = strSQL & "'" & Left(cmbʹ�÷�ʽ.Text, 1) & "',"
                '  �Ǽ�ʱ��_In In Ʊ�����ü�¼.�Ǽ�ʱ��%Type := Null,
                strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
                '  �Ǽ���_In   In Ʊ�����ü�¼.�Ǽ���%Type := Null,
                strSQL = strSQL & "'" & txtEdit(0).Text & "',"
                '  ����_In     In Ʊ�����ü�¼.����%Type := Null,
                strSQL = strSQL & "'" & txtEdit(5).Text & "',"
                '  ǩ����_In   In Ʊ�����ü�¼.ǩ����%Type := Null
                strSQL = strSQL & IIf(strUserName = "", "NULL", "'" & strUserName & "'") & ","
                '  ���id_In   In Ʊ�����ü�¼.���id%Type := Null
                strSQL = strSQL & IIf(mlng���ID = 0, "NULL", mlng���ID) & ")"
            End If
        Else
            'Zl_Ʊ�����ü�¼_Check
            strSQL = "Zl_Ʊ�����ü�¼_Check("
            '  Id_In       In Ʊ�����ü�¼.ID%Type,
            strSQL = strSQL & "" & mlng����ID & ","
            '  �˶Խ��_In In Ʊ�����ü�¼.�˶Խ��%Type,
            strSQL = strSQL & "" & IIf(optResult(0).Value, 1, 0) & ","
            '  �˶���_In   In Ʊ�����ü�¼.�˶���%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ��ע_In     In Ʊ�����ü�¼.��ע%Type,
            strSQL = strSQL & "'" & txtRemarks.Text & "',"
            '  �˶�ģʽ_In In Ʊ�����ü�¼.�˶�ģʽ%Type
            strSQL = strSQL & "" & 0 & ")"
        End If
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    lng����ID = 0
End Function

Private Sub ShowSum()
'����:��ʾ������Ϣ
    Dim strTemp As String
    Dim strName  As String
    
    strName = IIf(mblnIsBIll, "����", "����")
    
    '��ʼ����:
    '��������:
    'Ʊ��������:
    '
    '�Ѿ�ʹ�õ���С����:
    '�Ѿ�ʹ�õ�������:
    strTemp = vbCrLf & "  ��ʼ" & strName & "��" & Replace(txtEdit(1).Text, "&", "&&") & txtEdit(2).Text & vbCrLf
    strTemp = strTemp & "  ��ֹ" & strName & "��" & Replace(txtEdit(3).Text, "&", "&&") & txtEdit(4).Text & vbCrLf
    If txtEdit(2).Text = "" Or txtEdit(4).Text = "" Then
        strTemp = strTemp & "  " & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "��������" & vbCrLf
    Else
        strTemp = strTemp & "  " & IIf(mblnIsBIll, "Ʊ��", "��Ƭ") & "��������" & CalcTotal & vbCrLf
    End If
    If mstr��С���� <> "" Then
        strTemp = strTemp & "  �Ѿ�ʹ�õ���С" & strName & "��" & Replace(mstr��С����, "&", "&&") & vbCrLf
        strTemp = strTemp & "  �Ѿ�ʹ�õ����" & strName & "��" & Replace(mstr������, "&", "&&") & vbCrLf
    End If
    
    lbl˵��.Caption = strTemp
End Sub

Public Function ShowMe(ByVal frmOwner As Form, bytInFun As Byte, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng����ID As Long, Optional ByVal str��� As String = "", _
    Optional intKind As gBillType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������õĲ����ش��ڽ���ͨѶ�ĳ���,�������ӽɿ��¼
    '���:bytInFun:0-�������޸�,1-�˶����õ�
    '       str���-ȱʡ��ʹ�����
    '       intKind-�����洫���Ʊ��
    '����:
    '����:�༭�ɹ�����True,����ΪFalse
    '����:���˺�
    '����:2011-05-05 16:43:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long, blnFind As Boolean, i As Long
    
    mlngModule = lngModule: mstrPrivs = strPrivs: mbytInFun = bytInFun
    mstr��� = str���
    mstr��⿪ʼ�� = "": mstr�������� = ""
    mstr������ = "": mstr��С���� = ""
    '42618
    On Error GoTo errHandle
    If intKind <> 0 Then
        mstrPreType(intKind) = mstr���
    End If
    If UserInfo.���� = "" Then
        MsgBox "��Ϊ���Լ�ָ����Ӧ��Ա��������ʹ�ñ����ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mlng����ID = lng����ID
    Call InitContext
            
    mstrƱ�ݳ��� = zlDatabase.GetPara(20, glngSys, , , "7|7|7|7|7")
    Set mrs���� = Nothing
    Set mrs�ֶ� = Nothing
    
    With cmbƱ��
        For i = 0 To .ListCount - 1
            If .ItemData(i) = intKind Then .ListIndex = i: Exit For
        Next
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    If mlng����ID = 0 Then
        '����
        mstr��С���� = ""
        mstr������ = ""
        txtEdit(0).Text = UserInfo.����
        
        On Error Resume Next
        cmb������.Text = UserInfo.����
        If Err <> 0 Then
            If InStr(mstrPrivs, "���в���Ա") = 0 Then
                MsgBox "�㲻�߱���Ӧ����Ա���ʣ�û��Ȩ������Ʊ�ݡ�", vbInformation, gstrSysName
                mblnChange = False: Unload Me: Exit Function
            End If
        End If
        If InStr(mstrPrivs, "���в���Ա") = 0 Then cmb������.Enabled = False
        On Error GoTo errHandle
    Else
        '�޸�,��˶�
        If mintƱ�� = gBillType.���ѿ� Then
            gstrSQL = _
                "Select A.�ӿڱ�� As ʹ�����,A.������,A.ǰ׺�ı�,A.��ʼ���� As ��ʼ����,A.��ֹ���� As ��ֹ����," & vbNewLine & _
                "       A.ʹ�÷�ʽ,A.�Ǽ�ʱ��,A.�Ǽ���,A.��ǰ���� As ��ǰ����,A.ʣ������,A.����," & vbNewLine & _
                "       B.��ʼ���� as ��⿪ʼ��,B.��ֹ���� as �����ֹ�� " & vbNewLine & _
                "From ���ѿ����ü�¼ A,���ѿ�����¼ B  " & vbNewLine & _
                "Where A.ID=[1] And nvl(A.����,0)=B.ID(+) and A.�ӿڱ�� =B.�ӿڱ��(+)"
        Else
            gstrSQL = _
                "Select A.ʹ�����,A.������,A.ǰ׺�ı�,A.��ʼ����,A.��ֹ����,A.ʹ�÷�ʽ," & vbNewLine & _
                "       A.�Ǽ�ʱ��,A.�Ǽ���,A.��ǰ����,A.ʣ������,A.����," & vbNewLine & _
                "       B.��ʼ���� as ��⿪ʼ��,B.��ֹ���� as �����ֹ�� " & vbNewLine & _
                "From Ʊ�����ü�¼ A,Ʊ������¼ B  " & vbNewLine & _
                "Where A.ID=[1] And nvl(A.����,0)=B.ID(+) and A.Ʊ�� =B.Ʊ��(+)"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        If rsTmp.RecordCount = 0 Then Exit Function
        
        With cbo���
            For i = 0 To .ListCount - 1
                If mintƱ�� = gBillType.Ԥ���վ� Then
                     If .ItemData(i) = Val(NVL(rsTmp!ʹ�����)) Then
                        .ListIndex = i: blnFind = True: Exit For
                     End If
                ElseIf mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ� Then
                     If .ItemData(i) = Val(NVL(rsTmp!ʹ�����)) Then
                        .ListIndex = i: blnFind = True: Exit For
                     End If
                Else
                    If .List(i) = NVL(rsTmp!ʹ�����) Then
                       .ListIndex = i: blnFind = True: Exit For
                    End If
                End If
            Next
            If blnFind = False And Not (mintƱ�� = gBillType.���￨ Or mintƱ�� = gBillType.���ѿ�) Then
                .AddItem NVL(rsTmp!ʹ�����, " ")
                .ListIndex = .NewIndex
            End If
            
            .Enabled = IIf(NVL(rsTmp!��⿪ʼ��) = "", True, False)
            lblUserType.Enabled = .Enabled
        End With
        mlng���� = zlStr.ActualLen(NVL(rsTmp!��ʼ����))
        cmbƱ��.Enabled = False
        
        txtEdit(1).Text = IIf(IsNull(rsTmp("ǰ׺�ı�")), "", rsTmp("ǰ׺�ı�"))
        txtEdit(2).Text = Mid(rsTmp("��ʼ����"), Len(txtEdit(1).Text) + 1)
        txtEdit(2).Tag = txtEdit(2).Text
        txtEdit(4).Text = Mid(rsTmp("��ֹ����"), Len(txtEdit(1).Text) + 1)
        txtEdit(4).Tag = txtEdit(4).Text
        txtEdit(5).Text = "" & rsTmp!����
        txtEdit(5).Tag = "" & rsTmp!����
        cmbʹ�÷�ʽ.ListIndex = rsTmp("ʹ�÷�ʽ") - 1
        txtEdit(0).Text = UserInfo.����
        dtpDate.Value = rsTmp("�Ǽ�ʱ��")
        
        On Error Resume Next
        cmb������.Text = rsTmp("������")
        If Err <> 0 Then
            cmb������.AddItem rsTmp("������")
            cmb������.Text = rsTmp("������")
        End If
        If InStr(mstrPrivs, "���в���Ա") = 0 Then cmb������.Enabled = False
        On Error GoTo errHandle
        If NVL(rsTmp!��⿪ʼ��) <> "" And mbytInFun = 0 Then
            Call Load�ֶ�Ʊ��(mintƱ��, Trim(NVL(rsTmp!����)), NVL(rsTmp!��⿪ʼ��), NVL(rsTmp!�����ֹ��))
        End If
        
        If mintƱ�� = gBillType.���ѿ� Then
            gstrSQL = "Select Nvl(Min(����), ' ') As ��С����, Nvl(Max(����), ' ') As ������ From ���ѿ�ʹ�ü�¼ Where ����id = [1]"
        Else
            gstrSQL = "Select Nvl(Min(����), ' ') As ��С����, Nvl(Max(����), ' ') As ������ From Ʊ��ʹ����ϸ Where ����id = [1]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        
        mstr��С���� = Trim(rsTmp("��С����"))
        mstr������ = Trim(rsTmp("������"))
        If mstr��С���� <> "" Then
            '�����Ѿ�ʹ�ã���Щ���ݾͲ��ܸ���
            txtEdit(1).Enabled = False
            txtEdit(3).Enabled = False
            Call ShowSum
        End If
    End If
    
    mblnChange = False
    Me.Caption = IIf(mbytInFun = 0, "Ʊ�����õ�", "�˶����õ�")
    If mbytInFun = 0 Then
        fraCheck.Visible = False
        lbl˵��.Width = lbl˵��.Width + (cmbʹ�÷�ʽ.Left + cmbʹ�÷�ʽ.Width - (lbl˵��.Left + lbl˵��.Width))
    Else
        fraUse.Enabled = False
    End If
    Call SetCtrlEnable
    mblnOK = False
    frmBillEdit.Show vbModal, frmOwner
    ShowMe = mblnOK
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable���Ժ�Visible����
    '����:���˺�
    '����:2010-11-18 10:59:12
    '����:33725
    '---------------------------------------------------------------------------------------------------------------------------------------------
     If mbytInFun = 0 And mlng����ID = 0 Then
        cmd����.Visible = mbln���ȷ������
     Else
        cmd����.Visible = False
        txtEdit(5).Enabled = txtEdit(5).Enabled And Not mbln���ȷ������    '����
     End If
    cmdSel.Visible = mbln���ȷ������ And mbytInFun = 0
    txtEdit(1).Enabled = txtEdit(1).Enabled And Not mbln���ȷ������        '��ʼǰ׺�ı�
    txtEdit(3).Enabled = txtEdit(3).Enabled And Not mbln���ȷ������    ''����ǰ׺�ı�
    If txtEdit(1).Enabled = False Then
         txtEdit(1).BackColor = cmdOK.BackColor
    Else
         txtEdit(1).BackColor = txtEdit(2).BackColor
    End If
    txtEdit(3).BackColor = txtEdit(1).BackColor
    cmdOK.Enabled = True
End Sub

Private Function CalcTotal() As String
'���ܣ���ȡ���ú�������
    Dim strName As String
    
    strName = IIf(mblnIsBIll, "����", "����")
    '����43366
    If InStr(1, txtEdit(4).Text, ".") > 0 Or InStr(1, txtEdit(2).Text, ".") > 0 Then
        ShowMsgbox strName & "��Χ��������С�������������룡"
        Exit Function
    End If
    
    '����:28048:
     CalcTotal = zlStr.ExpressValue(txtEdit(4).Text & "-" & txtEdit(2).Text) + 1
    'CalcTotal = CDec(txtEdit(4).Text) - CDec(txtEdit(2).Text) + 1
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Combox����
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-27 10:22:29
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str��� As String
    
    On Error GoTo errHandle
    str��� = mstrPreType(mintƱ��)
    Select Case mintƱ��
    Case gBillType.�շ��վ�, gBillType.�����վ�
        strSQL = "Select ����,����,����,ȱʡ��־ From Ʊ��ʹ����� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����)
                .ItemData(.NewIndex) = 1
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            .AddItem " "    '��������Ϊ��
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            mblnNotClick = False
            .Enabled = True: lblUserType.Enabled = True
        End With
    Case gBillType.Ԥ���վ�
        mblnNotClick = True
        With cbo���
            .Clear
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") Then
                .AddItem "����Ԥ��": .ItemData(.NewIndex) = 1
                If Val(str���) = 1 Then .ListIndex = .NewIndex
            End If
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��") Then
                .AddItem "סԺԤ��": .ItemData(.NewIndex) = 2
                If Val(str���) = 2 Then .ListIndex = .NewIndex
            End If
            '58071
            If zlStr.IsHavePrivs(mstrPrivs, "Ԥ������Ʊ��") _
                And zlStr.IsHavePrivs(mstrPrivs, "Ԥ��סԺƱ��") Then
                .AddItem " "
                .ItemData(.NewIndex) = 0
                If Val(str���) = 0 Then .ListIndex = .NewIndex
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            .Enabled = True
        End With
        mblnNotClick = False
    Case gBillType.���￨
        strSQL = _
            "Select ID, ����, ����, ȱʡ��־, ���ų���, ��������, ǰ׺�ı�, ��������" & vbNewLine & _
            "From ҽ�ƿ����" & vbNewLine & _
            "Where Nvl(�Ƿ�����, 0) >= 1" & vbNewLine & _
            "Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            Set mcllCardProperty = New Collection
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!���ų���)), CStr(NVL(rsTemp!ǰ׺�ı�)), _
                    CStr(NVL(rsTemp!��������)), Val(NVL(rsTemp!��������))), "K" & Val(NVL(rsTemp!ID))
                If Val(NVL(rsTemp!ȱʡ��־)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str���) = Val(NVL(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            mblnNotClick = False: .Enabled = True
        End With
    Case gBillType.���ѿ�
        strSQL = _
            "Select ���, ����, ���ų���, ǰ׺�ı�, �Ƿ�����, 0 As ��������" & vbNewLine & _
            "From ���ѿ����Ŀ¼" & vbNewLine & _
            "Where Nvl(����, 0) >= 1" & vbNewLine & _
            "Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        With cbo���
            .Clear
            Set mcllCardProperty = New Collection
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!���) & "-" & NVL(rsTemp!����)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!���))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!���ų���)), CStr(NVL(rsTemp!ǰ׺�ı�)), _
                    CStr(NVL(rsTemp!�Ƿ�����)), Val(NVL(rsTemp!��������))), "K" & Val(NVL(rsTemp!���))
                If Val(str���) = Val(NVL(rsTemp!���)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            mblnNotClick = False: .Enabled = True
        End With
    Case Else
        cbo���.Enabled = False: lblUserType.Enabled = False
        cbo���.ListIndex = -1
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
