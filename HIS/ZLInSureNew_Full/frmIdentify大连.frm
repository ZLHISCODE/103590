VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt���ժҪ 
      Height          =   300
      Left            =   4635
      TabIndex        =   49
      Top             =   3990
      Width           =   2500
   End
   Begin VB.TextBox Txt����޶� 
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   29
      Top             =   2910
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   6885
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3660
      Width           =   255
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5730
      TabIndex        =   33
      Top             =   3285
      Width           =   1425
   End
   Begin VB.TextBox Txt���� 
      Height          =   300
      Left            =   1080
      TabIndex        =   38
      Top             =   4035
      Width           =   2715
   End
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "����(&R)"
      Height          =   350
      Left            =   180
      TabIndex        =   41
      Top             =   4605
      Width           =   1100
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   4035
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox Txtת�ﵥ�� 
      Height          =   300
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   31
      Top             =   3270
      Width           =   3195
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -30
      TabIndex        =   47
      Top             =   630
      Width           =   8340
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2910
      Width           =   6075
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "���˿���(&G)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5820
      TabIndex        =   45
      Top             =   2490
      Width           =   1305
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������(&S)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4470
      TabIndex        =   44
      Top             =   2490
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6030
      TabIndex        =   43
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4650
      TabIndex        =   42
      Top             =   4605
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -120
      TabIndex        =   46
      Top             =   4425
      Width           =   8340
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1080
      TabIndex        =   35
      Top             =   3645
      Width           =   6060
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   6675
      TabIndex        =   52
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   18
      Left            =   6255
      TabIndex        =   51
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "����ժҪ"
      Height          =   375
      Left            =   3870
      TabIndex        =   50
      Top             =   4065
      Width           =   750
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   420
      TabIndex        =   34
      Top             =   3690
      Width           =   630
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����(&Q)"
      Enabled         =   0   'False
      Height          =   180
      Left            =   4920
      TabIndex        =   32
      Top             =   3345
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������(&M)"
      Height          =   180
      Index           =   16
      Left            =   60
      TabIndex        =   37
      Top             =   4095
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������(&Z)"
      Height          =   180
      Index           =   15
      Left            =   60
      TabIndex        =   39
      Top             =   4095
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ת�ﵥ��(&N)"
      Height          =   180
      Index           =   14
      Left            =   60
      TabIndex        =   30
      Top             =   3330
      Width           =   990
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ͨ��IC����֤��Ա��ݣ�������֤�����Ϣ��ʾ������ͬʱ�ɶԾ���������ѡ��"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   390
      Width           =   6540
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   30
      Picture         =   "frmIdentify����.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������(&J)"
      Height          =   180
      Index           =   13
      Left            =   60
      TabIndex        =   27
      Top             =   2970
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ�״̬"
      Height          =   180
      Index           =   12
      Left            =   2190
      TabIndex        =   25
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   2940
      TabIndex        =   26
      Top             =   2527
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   1080
      TabIndex        =   24
      Top             =   2527
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�α����"
      Height          =   180
      Index           =   11
      Left            =   330
      TabIndex        =   23
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   10
      Left            =   5820
      TabIndex        =   22
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���������ʻ����"
      Height          =   180
      Index           =   10
      Left            =   4350
      TabIndex        =   21
      Top             =   2190
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   9
      Left            =   2940
      TabIndex        =   20
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�½ɻ���"
      Height          =   180
      Index           =   9
      Left            =   2190
      TabIndex        =   19
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   1080
      TabIndex        =   18
      Top             =   2130
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ͳ���ۼ�"
      Height          =   180
      Index           =   8
      Left            =   330
      TabIndex        =   17
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   5820
      TabIndex        =   16
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���������ʻ����"
      Height          =   180
      Index           =   7
      Left            =   4350
      TabIndex        =   15
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   2940
      TabIndex        =   14
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   2940
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ҽ���"
      Height          =   180
      Index           =   6
      Left            =   2190
      TabIndex        =   13
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   1080
      TabIndex        =   12
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   11
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   4
      Left            =   5820
      TabIndex        =   10
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "IC����"
      Height          =   180
      Index           =   4
      Left            =   5250
      TabIndex        =   9
      Top             =   1410
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   1350
      Width           =   3195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   7
      Top             =   1410
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5400
      TabIndex        =   6
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   2
      Left            =   4980
      TabIndex        =   5
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   2550
      TabIndex        =   3
      Top             =   1020
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����޶�(&K)"
      Height          =   180
      Index           =   17
      Left            =   45
      TabIndex        =   48
      Top             =   2970
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99

Dim mblnFirst  As Boolean
Dim mstrReturn As String    '������Ϣ��
Dim mlng����ID As Long
'mbytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
Dim mbytType As Byte
Dim mblnOK As Boolean
Dim mlng���� As Long
Dim mlng��¼ID As Long
Dim mbytCallType As Byte  '(1-���ʴ�����;0-���˷��ò�ѯ���õ�)�����������Ч��byttype=4�����
Dim mint���� As Integer

Private Sub cbo�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������֤
    '--�����:
    '--������:
    '--��  ��:��֤�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    IsValid = False
    
    If LenB(StrConv(Trim(Txtת�ﵥ��.Text), vbFromUnicode)) > 6 Then
        ShowMsgbox "ת�ﵥ�ų�����,���������6���ַ�!"
        If Txtת�ﵥ��.Enabled Then Txtת�ﵥ��.SetFocus
        Exit Function
    End If
    If Me.cbo�������.ListIndex < 0 Then
        ShowMsgbox "����������ѡ��!"
        If cbo�������.Enabled Then cbo�������.SetFocus
        Exit Function
    End If
    If Me.cbo����.ListIndex < 0 Then
        ShowMsgbox "���ı���ѡ��!"
        If cbo����.Enabled Then cbo����.SetFocus
        Exit Function
    End If
    If Txtת�ﵥ��.Text <> "" And mbytType <> 0 And mbytType <> 3 Then
        If Val(txt����.Text) = 0 Then
            Dim blnYes As Boolean
            
            ShowMsgbox "����δ����,�Ƿ���Դ���?", True, blnYes
            
            If blnYes = False Then
                If txt����.Enabled Then txt����.SetFocus
                Exit Function
            End If
        End If
    End If
    Dim lng���� As Long
     lng���� = cbo�������.ItemData(cbo�������.ListIndex)
     If (lng���� = 3 Or lng���� = 4) And Trim(Txt����.Text) = "" Then
        ShowMsgbox "�󲡻�������������������!"
        Exit Function
     End If
    '����û�״̬
    '   A������B��ֹ����Cȫֹ����D����
    'mbytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    Select Case g�������_����.�ʻ�״̬
        Case "A"
        Case "B"
            If mbytType = 4 Then
            ShowMsgbox "�ò���״̬Ϊ����ֹ����״̬,ֻ��������ʹ��!"
            End If
        Case "C"
            ShowMsgbox "�ò���״̬Ϊ��ȫֹ����״̬,ֻ�����ֽ����!"
        Case "D"
            ShowMsgbox "�ò�����ҽ����������,���ܼ���!"
            Exit Function
    End Select
    
    '��鲡��״̬
    Dim lng����ID As Long
    gstrSQL = "select ����id,nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, g�������_����.���˱��)
    If mbytType <> 4 Then   '����סԺ����ʱ������֤��ǰ״̬
        If rsTemp.RecordCount > 0 Then
            If rsTemp("״̬") > 0 Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        '2005-01-04 ZHQ�޸�
        '���ܣ����ҽ������δ�Һţ��������շ�
        If rsTemp.RecordCount > 0 Then
            lng����ID = Nvl(rsTemp.Fields("����ID").Value, 0)
        Else
            lng����ID = 0
        End If
        If mbytType = 0 Then
            Dim lngRegDay As Long   '���˹Һ����������
            #If gverControl >= 4 Then
                lngRegDay = Val(zlDatabase.GetPara(21, glngSys, , "0"))
            #Else
                lngRegDay = Val(GetPara(21, glngSys, , , "0"))
            #End If
            
            If lngRegDay <> 0 Then  '=0��ʾ�������ж�
                #If gverControl >= 5 Then
                    gstrSQL = "Select No,�����,����ID From ���˹Һż�¼ " & _
                            "  Where sysdate-�Ǽ�ʱ��<=" & lngRegDay & _
                            "  And ��¼����=1 And ��¼״̬=1 And ����ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ʵ�ʹҺż�¼", lng����ID)
                #Else
                    gstrSQL = "Select No,�����,����ID From ���˹Һż�¼ " & _
                            "  Where sysdate-�Ǽ�ʱ��<=[1] And ����ID=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ʵ�ʹҺż�¼", lngRegDay, lng����ID)
                #End If
                If rsTemp.RecordCount <= 0 Then
                    MsgBox "�˲��� " & lngRegDay & " ����δ�Һţ�����ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        'סԺ����ʱ,�账���Ƿ�Ϊͬһ����
        If rsTemp.EOF Then
            ShowMsgbox "�ڱ����ʻ��в����ڵ�ǰ����!"
            Exit Function
        Else
            lng����ID = Nvl(rsTemp!����ID, 0)
            If mlng����ID <> lng����ID Then
                ShowMsgbox "��ٽ��ʵĵ�ǰ�����������֤�Ĳ��˲�һ��!"
                Exit Function
            End If
        End If
    End If
    If Txt����.Tag = "" And Txt����.Text <> "" And mbytType <> 3 Then
        ShowMsgbox "ѡ�������������,���ܼ���!"
        Exit Function
    End If
    
    
    If cbo�������.ItemData(cbo�������.ListIndex) = 3 Then
        '���������Ϊ������󲡡�ʱ�������벡��
        '20040621���˺����
        If Val(txt����.Tag) = 0 Then
            ShowMsgbox "����󲡱������벡��!"
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Function
        End If
    End If
    IsValid = True
End Function

Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mstrReturn = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strTmp1 As String
    
    If mlng����ID <> 0 And mbytType > 4 Then
        With g�������_����
             '���渽����Ϣ
             '����:
             '   ����_IN,��¼id_IN,����޶�_IN
             
             gstrSQL = "zl_���ս����¼�޶�_Update(" & _
                 mlng���� & "," & _
                 mlng��¼ID & "," & _
                 Val(Txt����޶�.Text) & ")" & _
                 ""
                 Err = 0
                 On Error Resume Next
                 gcnOracle.Execute gstrSQL, Me.Caption
                 If Err <> 0 Then
                     ShowMsgbox "���ս����¼������޶��ʧ��!"
                     Exit Sub
                 End If
         End With
         mstrReturn = mlng����ID
        Unload Me
        Exit Sub
    End If
    '��֤����
    If IsValid = False Then Exit Sub
    
    With g�������_����
        .ת�ﵥ�� = Trim(Txtת�ﵥ��)
        .������� = cbo�������.ItemData(cbo�������.ListIndex)
        
        If Txt����.Tag = "" Then
            .��ϱ��� = ""
            .������� = ""
        Else
            .��ϱ��� = Split(Txt����.Tag, "|||")(0)
            .������� = Split(Txt����.Tag, "|||")(1)
        End If
        If mbytType = 0 Then
            .���� = 0
        Else
            If .ת�ﵥ�� = "" Then
                '��ȡ����
                If Val(txt����.Text) = 0 Then
                    .���� = Get����(.ְ����ҽ���, .����, mint����)
                Else
                    .���� = Val(txt����.Text)
                End If
            Else
                .���� = Val(txt����.Text)
            End If
            
        End If
    End With
    
        
    'ȷ����ط��ش�
    
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�

    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8������ID��
    '9����;10.˳���;11��Ա���;12�ʻ����;13��ǰ״̬;14����ID;15��ְ(0,1);16����֤��;17�����;18�Ҷȼ�
    '19�ʻ������ۼ�,20�ʻ�֧���ۼ�,21����ͳ���ۼ�,22ͳ�ﱨ���ۼ�,23סԺ�����ۼ�;24�������� (1����������);25��������
    Dim int��ǰ״̬ As Integer, strUnitName As String
    int��ǰ״̬ = 0
    If mbytType = 3 Or mbytCallType = 0 Then
        '����ǹҺ�,����ȷ�����û��Ƿ����
        'mbytCallType (1-���ʴ�����;0-���˷��ò�ѯ���õ�)�����������Ч��byttype=4�����
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��.(��˳�����:2004/06/11)
        gstrSQL = "Select a.����ID,a.��ǰ״̬,b.������λ from �����ʻ� a,������Ϣ b " & _
                "  Where a.����ID=b.����ID And a.����=" & mint���� & " and  a.ҽ����='" & g�������_����.���˱�� & "'"
        Dim rsTemp As New ADODB.Recordset
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
            strUnitName = Nvl(rsTemp!������λ, "")
        End If
    End If
     
     mstrReturn = ""
    With g�������_����
        strTmp = .IC����                    '0����
        strTmp = strTmp & ";" & .���˱��   '1ҽ����
        strTmp = strTmp & ";"               '2����
        strTmp = strTmp & ";" & .����       '3����
        strTmp = strTmp & ";" & .�Ա�       '4�Ա�
        strTmp = strTmp & ";" & .��������   '5��������
        strTmp = strTmp & ";" & .���֤��   '6���֤
        strTmp = strTmp & ";" & strUnitName '7��λ����(����)
        
        strTmp1 = ""
        strTmp1 = strTmp1 & ";"    '8���Ĵ���
        strTmp1 = strTmp1 & ";" & .�������   '9˳���
        strTmp1 = strTmp1 & ";" & .ת�ﵥ��  '10��Ա���,�����ת�ﵥ��
        strTmp1 = strTmp1 & ";" & .���������ʻ����       '11�ʻ����
        strTmp1 = strTmp1 & ";" & int��ǰ״̬               '12��ǰ״̬
        strTmp1 = strTmp1 & ";" & IIf(Val(Me.txt����.Tag) = 0, "", Me.txt����.Tag)             '13����ID
        '���˺�:20040911,����������
        'ҽ������Ϊ,A��ְ��B���ݡ�L���ݡ�T����,Q ��ҵ����,E����
        strTmp1 = strTmp1 & ";" & Decode(.ְ����ҽ���, "A", 1, "B", 2, "L", 3, "T", 4, "Q", 5, "E", 6, 1) '.�������  '14��ְ(0,1)
        strTmp1 = strTmp1 & ";" & .���������ʻ���� '15����֤��,Ŀǰ�Ҵ���ǲ��������ʻ����
        strTmp1 = strTmp1 & ";" & IIf(.���� = 0, "", .����) '16�����
        strTmp1 = strTmp1 & ";" & .�������       '17�Ҷȼ�,��ľ���������
        strTmp1 = strTmp1 & ";" & .���������ʻ����         '18�ʻ������ۼ�
        strTmp1 = strTmp1 & ";0"        '19�ʻ�֧���ۼ�
        strTmp1 = strTmp1 & ";" & .ͳ���ۼ�  '20����ͳ���ۼ�
        strTmp1 = strTmp1 & ";" & .����          '21ͳ�ﱨ���ۼ�
        strTmp1 = strTmp1 & ";0"        '22סԺ�����ۼ�
    End With
    
    '--------------------------------------------------------------------------
    '2004-06-08,ȡ���Һ���Ŀ������,�豣����صĲ�����Ϣ.
    'If mlng����ID <> 0 And mbytType = 3 Then
        
    'Else
        mlng����ID = BuildPatiInfo(0, strTmp & strTmp1, mlng����ID, mint����)
   ' End If
    
    With g�������_����
        '���渽����Ϣ
        '����:
        '    ����_IN,����id_IN,�α����1_IN,�α����2_IN,�α����3_IN,�α����4_IN,�α����5_IN,
        gstrSQL = "zl_�����ʻ�����_Update(" & _
            mint���� & "," & _
            mlng����ID & "," & _
            Val(.�α����1) & "," & _
            Val(.�α����2) & "," & _
            Val(.�α����3) & "," & _
            Val(.�α����4) & "," & _
            Val(.�α����5) & ")" & _
            ""
            Err = 0
            On Error Resume Next
            gcnOracle.Execute gstrSQL, Me.Caption
            If Err <> 0 Then
                ShowMsgbox "�����ʻ�������Ϣ����ʧ��,������Щ��Ϣ��������ʹ��!"
                Exit Sub
            End If
            Dim strҽ�Ƹ��ʽ As String
            
            If InStr(1, "ABE", .ְ����ҽ���) <> 0 Then
                'A��ְ,B.����,E��������)
                strҽ�Ƹ��ʽ = "������ҽ�Ʊ���"
            End If
            
            If .ҽ������ = 1 Then
                If Val(.�α����4) = 1 Then
                    '0���������á�1��������
                    strҽ�Ƹ��ʽ = "��������"
                End If
                If Val(.�α����5) = 1 Then
                    '0���˲����á�1���˿���
                    strҽ�Ƹ��ʽ = "���˱���"
                End If
                If InStr(1, "LT", .ְ����ҽ���) <> 0 Then
                    strҽ�Ƹ��ʽ = "����ҽ��"
                End If
                If InStr(1, "LT", .ְ����ҽ���) <> 0 Then
                    'T.����,L.����
                    strҽ�Ƹ��ʽ = "����ҽ��"
                End If
                If .ְ����ҽ��� = "Q" Then
                    '��ҵ����
                    strҽ�Ƹ��ʽ = "��ҵ����"
                End If
            Else
                If mint���� = TYPE_���������� Then
                    strҽ�Ƹ��ʽ = "������ҽ�Ʊ���"
                End If
            End If
            Err = 0
            On Error GoTo errHand:
            '���²�����Ϣ��ҽ�Ƹ��ʽ
            gstrSQL = "zl_������Ϣҽ�Ƹ���_Update(" & mlng����ID & ",'" & _
                strҽ�Ƹ��ʽ & "')"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    '���ظ�ʽ:�м���벡��ID
   '--�����֤��ǿ�ƶԲ�����Ϣ���������һ�θ���,��ֹ������ֿ�
   gstrSQL = "update ������Ϣ A set A.����=(select ���� from �����ʻ� where ����id=A.����id) where ����id=" & mlng����ID
   gcnOracle.Execute gstrSQL
    
    If mlng����ID > 0 Then
        mstrReturn = strTmp & ";" & mlng����ID & strTmp1
    End If
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & mint����
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt����
    End If
    txt����.SetFocus
End Sub

Private Sub cmd�鿨_Click()
    SetCtlEn False
    mblnOK = ReadCard
    
    '�����������
    '�ܺ�ȫ  2005-02-17
    Dim dtAge As Variant
    If Len(Trim(lblEdit(3).Caption)) <> 0 Then
        Select Case Len(Trim(lblEdit(3).Caption))
        Case 15
            dtAge = "19" & Substr(Trim(lblEdit(3).Caption), 7, 6)
        Case 18
            dtAge = Substr(Trim(lblEdit(3).Caption), 7, 8)
        Case Else
            dtAge = Null
        End Select
        If IsNull(dtAge) Then
            lblEdit(13).Caption = ""
        Else
            lblEdit(13).Caption = CInt(Format(zlDatabase.Currentdate, "yyyy")) - CInt(Substr(dtAge, 1, 4))
            If Format(zlDatabase.Currentdate, "MMdd") < Substr(dtAge, 5, 4) Then
                lblEdit(13).Caption = CInt(lblEdit(13).Caption) - 1
            End If
            If CInt(lblEdit(13).Caption) < 0 Then lblEdit(13).Caption = 0
        End If
    End If
    
    If Txtת�ﵥ�� = "" And mbytType <> 0 Then
        Me.txt���� = Format(Get����(g�������_����.ְ����ҽ���, g�������_����.����, mint����), "###,###0.00;-###,###0.00; ;")
    End If
    SetCtlEn True
    If Txt����.Enabled Then
        Txt����.SetFocus
    ElseIf cmdOK.Enabled Then
        cmdOK.SetFocus
    End If
End Sub
Private Sub SetCtlEn(ByVal blnTrue As Boolean)
    cmd�鿨.Enabled = blnTrue
    cmdOK.Enabled = blnTrue And mblnOK
    txt����.Enabled = blnTrue And mbytType = 4
    lbl����.Enabled = blnTrue And mbytType = 4
    cmdCancel.Enabled = blnTrue
    Txtת�ﵥ��.Enabled = blnTrue And mbytType <> 3
    cbo�������.Enabled = blnTrue And mbytType <> 3
    txt����.Enabled = blnTrue And mbytType <> 3
    cmd����.Enabled = blnTrue And mbytType <> 3
    Txt����.Enabled = blnTrue And mbytType = 0 And mblnOK
    lbl(16).Enabled = blnTrue And mbytType = 0 And mblnOK
    cbo����.Enabled = blnTrue
    Txt���ժҪ.Enabled = blnTrue And mblnOK
    Txt���ժҪ.Locked = True
        
End Sub
Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnOK = False
    Txt����.Tag = ""
    Txt����.Text = ""
    txt����.Text = ""
    SetCtlEn True
    If cbo�������.Enabled And mlng����ID = 0 Then
        cbo�������.SetFocus
    ElseIf Txt����޶�.Visible Then
        Txt����޶�.SetFocus
    End If
End Sub
Private Function ReadCard() As Boolean
    
    ReadCard = False
   '��֤�û����
    If ��ȡ�������_����(IIf(mint���� = TYPE_����������, 2, 1), mint����) = False Then
        Exit Function
    End If
    Call SetCtlData
    ReadCard = True
End Function
Private Function SetCtlData()
    '����:���ÿؼ�����
    Dim int�Ա� As Integer
    Dim int�������� As Integer
    Dim rsTemp As New ADODB.Recordset
        
    Txt���ժҪ = ""
    
    Err = 0
    On Error Resume Next
    '������������Ϣ��ֵ
    With g�������_����
        lblEdit(0).Caption = .���˱��
        lblEdit(1).Caption = .����
        lblEdit(3).Caption = Trim(.���֤��)
        int�Ա� = Val(IIf(Len(lblEdit(3)) = 18, Mid(lblEdit(3), 17, 1), Right(lblEdit(3), 1))) Mod 2
        '�������֤ȡ����Ӧ���Ա�
        lblEdit(2).Caption = IIf(int�Ա� = 0, "Ů", "��")
        .�������� = zlCommFun.GetIDCardDate(Trim(.���֤��))
        '��������
        If IsDate(.��������) And .�������� <> "" Then
            .���� = Abs(Int((zlDatabase.Currentdate - CDate(.��������)) / 365))
        Else
            .���� = 0
        End If
        
        .�Ա� = lblEdit(2).Caption
        lblEdit(4).Caption = .IC����
        lblEdit(5).Caption = .�������
        '2004/09/11:��������
        lblEdit(6).Caption = Decode(.ְ����ҽ���, "A", "��ְ", "B", "����", "L", "����", "T", "����", "Q", "��ҵ����", "E", "����", "δ֪")
        lblEdit(7).Caption = Format(.���������ʻ����, "###,###0.00;-###,###0.00; ;")
        lblEdit(8).Caption = Format(.ͳ���ۼ�, "###,###0.00;-###,###0.00; ;")
        lblEdit(9).Caption = Format(.�½ɷѻ���, "###,###0.00;-###,###0.00; ;")
        lblEdit(10).Caption = Format(.���������ʻ����, "###,###0.00;-###,###0.00; ;")
        
        lblEdit(11).Caption = Decode(.�α����3, "0", "��", "1", "�±�", "��������Ա")
        lblEdit(12).Caption = Decode(.�ʻ�״̬, "A", "����", "B", "��ֹ��", "C", "ȫֹ��", "D", "����", "����ȷ��")
        chk����.Value = IIf(.�α����4 = 1, 1, 0)
        chk����.Value = IIf(.�α����5 = 1, 1, 0)
        
        If mbytType <> 4 Then
            int�������� = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9OutExse\", "��Ѱ���۵���", "")
            gstrSQL = "Select Distinct(����) as ���� From ������ü�¼ Where ��¼���� = 4 " & _
                      "And No=(select max(�Һŵ�) from ����ҽ����¼ where id=(select max(id) from ����ҽ����¼  " & _
                      " Where ����id=(select distinct(����id)  from �����ʻ� where ҽ����='" & .���˱�� & _
                      "') And ����ʱ��>trunc(Sysdate)-" & int�������� & "))"
        Else
            '2005-10-14 ZHQ
            '����ʱֱ����ȡ��Ժ���
            gstrSQL = "Select ������Ϣ as ���� From ������ " & _
                    "   Where �������=3 And ��ϴ���=1 And ����ID In (Select ����ID From �����ʻ� where ҽ����='" & .���˱�� & "')" & _
                    "   Order by ��ҳID Desc"
        End If
        rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
        If Not rsTemp.EOF Then
            Txt���ժҪ = Nvl(rsTemp!����, "")
        Else
            Txt���ժҪ = ""
        End If
        
    End With
End Function

Private Function LoadCobData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ؾ���������ݺ���������
    '--�����:
    '--������:
    '--��  ��:���سɹ�,����True,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Me.cbo�������.Clear
    '   bytType-����(0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����)
    With cbo�������
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
        
           '  1-1,2-3,3-5,4-"S"
            .AddItem "��ͨ����"
            .ItemData(.NewIndex) = 1
            .AddItem "��������"
            .ItemData(.NewIndex) = 2
            .AddItem "�����"
            .ItemData(.NewIndex) = 3
            .AddItem "������������"
            .ItemData(.NewIndex) = 4
        End If
        If mbytType = 1 Or mbytType = 2 Or mbytType = 4 Then
            '5-2,6-4,7-"O",8-"Q"
            
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 5
            .AddItem "��ͥ����סԺ"
            .ItemData(.NewIndex) = 6
            .AddItem "��������סԺ"
            .ItemData(.NewIndex) = 7
            .AddItem "���˱���סԺ"
            .ItemData(.NewIndex) = 8
        End If
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    
    '����ҽ����������
    strSQL = "Select * From ��������Ŀ¼ where ����=" & mint���� & " Order by ���"
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    If rsTmp.RecordCount = 0 Then
        ShowMsgbox "ҽ������δ����,���ڱ�����������������!"
        Exit Function
    End If
    With rsTmp
        cbo����.Clear
        Do While Not .EOF
            cbo����.AddItem Nvl(!����) & "-" & Nvl(!����)
            cbo����.ItemData(cbo����.NewIndex) = Nvl(!���, 0)
            If Nvl(!���, 0) = 2 And gblnKFQCom_���� Then
                cbo����.ListIndex = cbo����.NewIndex
            End If
            .MoveNext
        Loop
        If cbo����.ListCount <> 0 Then
            If cbo����.ListIndex < 0 Then
               cbo����.ListIndex = 0
            End If
        End If
    End With
    
    LoadCobData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetPatient(ByVal intinsure As Integer, ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0, _
                Optional lng���� As Long, Optional lng��¼ID As Long, Optional bytCallType As Byte = 1) As String
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ���˵������Ϣ
    '--�����:bytType-����(mbytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����)
    '         lng����ID-����ID
    '         lng����-�����¼�е�����
    '         lng��¼id-���̼�¼�е�id
    '         bytCallType(1-���ʴ�����;0-���˷��ò�ѯ���õ�)�����������Ч��byttype=4�����
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    mstrReturn = ""
    mlng����ID = lng����ID
    mint���� = intinsure
    mbytType = bytType
    
    mlng���� = lng����
    mlng��¼ID = lng��¼ID
    
    If lng����ID <> 0 And mbytType > 4 Then
        '��ȷ����ز�����Ϣ

        gstrSQL = "select b.ҽ����,a.����,a.�Ա�,a.����,a.��������,a.���֤��, " & _
                 "        b.����,b.�Ҷȼ� as �������,b.˳���,b.����֤�� as ���������ʻ����,b.����id, " & _
                 "        b.�α����1,b.�α����2,b.�α����3,b.�α����4,b.�α����5, " & _
                 "        b.��Ա��� as ת�ﵥ��,b.�ʻ���� as �����ʻ����,b.��ְ as ְ����ҽ��� " & _
                 " from ������Ϣ a,�����ʻ� b " & _
                 " where a.����id=b.����id and b.����=" & mint���� & " and a.����id=" & lng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        With rsTemp
            If Not .EOF Then
                g�������_����.IC���� = Nvl(!����)
                g�������_����.���������ʻ���� = Nvl(!���������ʻ����, 0)
                g�������_����.�����ʻ���ǰֵ = 0
                g�������_����.�����ʻ�ԭʼֵ = 0
                g�������_����.�α����1 = Nvl(!�α����1)
                g�������_����.�α����2 = Nvl(!�α����2)
                g�������_����.�α����3 = Nvl(!�α����3)
                g�������_����.�α����4 = Nvl(!�α����4)
                g�������_����.�α����5 = Nvl(!�α����5)
                g�������_����.�������� = Format(!��������, "yyyy-mm-dd")
                g�������_����.���˱�� = Nvl(!ҽ����)
                g�������_����.���������ʻ���� = Nvl(!�����ʻ����, 0)
                g�������_����.������� = Decode(Nvl(!�������), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, "S", 4, "T", 4, "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
                g�������_����.�����ʻ�״̬ = 0
                
                g�������_����.���� = Val(Nvl(!����))
                g�������_����.���� = 0
                g�������_����.���֤�� = Nvl(!���֤��)
                g�������_����.ͳ���ۼ� = 0
                g�������_����.���� = Nvl(!����)
                g�������_����.�Ա� = Nvl(!�Ա�)
                g�������_����.ҽ������ = IIf(mint���� = 82, 1, 2)
                g�������_����.�½ɷѻ��� = 0
                g�������_����.�ʻ�״̬ = ""
                g�������_����.��ϱ��� = ""
                g�������_����.������� = ""
                g�������_����.֧����� = 0
                g�������_����.ְ����ҽ��� = Nvl(!ְ����ҽ���)
                
                g�������_����.������� = 0
                g�������_����.ת�ﵥ�� = Nvl(!ת�ﵥ��)
                
                gstrSQL = "Select ����޶� from ���ս����¼ where ����=" & mlng���� & " and ��¼id=" & mlng��¼ID
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If .EOF Then
                    Txt����޶�.Text = ""
                Else
                    Txt����޶�.Text = Format(!����޶�, "####0.00;-#####0.00; ;")
                End If
                
                '��������
                Call SetCtlData
                '
                SetCtlVisible
                Me.Height = 4185
                fra(0).Top = 3255
                cmdOK.Top = fra(0).Top + fra(0).Height + 40
                cmdCancel.Top = cmdOK.Top
                Me.Caption = "��������޶�¼��"
                lblInfor.Caption = "���벡�˵�����޶"
            End If
        End With
    End If
    
    
    Me.Show 1
    GetPatient = mstrReturn
End Function
Private Sub SetCtlVisible()
    '���ÿؼ���Vizible
    cbo�������.Visible = False
    Txtת�ﵥ��.Visible = False
    txt����.Visible = False
    txt����.Visible = False
    cmd����.Visible = False
    Txt����.Visible = False
    Txt����޶�.Visible = True
    lbl(17).Visible = True
    lbl(16).Visible = False
    lbl(14).Visible = False
    lbl(13).Visible = False
    lbl����.Visible = False
    lbl����.Visible = False
    cmdOK.Enabled = False
    cmd�鿨.Visible = False
    Txt���ժҪ.Visible = False
    
End Sub
Private Sub Form_Load()
    mblnFirst = True
    
    '���ط�������ҽ������
    Call LoadCobData
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����.Text = ""
        txt����.Tag = ""
    End If
End Sub

Private Sub Txt����_Change()
    Txt����.Tag = ""
End Sub

Private Sub Txt����_GotFocus()
    zlControl.TxtSelAll Txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt����.Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = UCase(Txt����.Text)
            str�Ա� = g�������_����.�Ա�
            If str�Ա� = "��" Then
                str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
            ElseIf str�Ա� = "Ů" Then
                str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
            End If

            strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
                " And (A.���� Like '" & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str�Ա� & _
                " Order by A.���,A.����"

            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������Input", , , , , , True, _
                Txt����.Left + Me.Left, _
                Txt����.Top + Me.Top, Txt����.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt����.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                Txt����.Tag = rsTmp!���� & "|||" & rsTmp!����
                If cmdOK.Enabled Then
                    cmdOK.SetFocus
                Else
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                Call Txt����_GotFocus
                Txt����.SetFocus
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt����, KeyAscii, m�ı�ʽ
    End If
End Sub

Private Sub Txt����_LostFocus()
    '--2004-12-28   ZHQ
    'ҽ��Ҫ��������ϱ�������
    If mbytType = 0 Then
        If Len(Trim(Txt����)) = 0 Then
            Txt����.SetFocus
        End If
    End If
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt����, KeyAscii, m���ʽ
End Sub

Private Sub Txtת�ﵥ��_Change()
    txt����.Enabled = Txtת�ﵥ�� <> "" And mbytType <> 0
    lbl����.Enabled = txt����.Enabled
End Sub

Private Sub Txtת�ﵥ��_GotFocus()
   zlControl.TxtSelAll Txtת�ﵥ��
End Sub

Private Sub Txtת�ﵥ��_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
    txt����.ForeColor = &HC0&
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" Or txt����.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt����.Text
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ','��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=[1] And (A.���� like [1] || '%' or A.���� like [1] || '%' or A.���� like [1] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, strText)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(mint����, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        txt����.ForeColor = Txtת�ﵥ��.ForeColor
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Txtת�ﵥ��_KeyPress(KeyAscii As Integer)
    KeyAscii = asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt����޶�_Change()
     cmdOK.Enabled = Val(Txt����޶�.Text) <> 0
End Sub

Private Sub Txt����޶�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt����޶�_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt����޶�, KeyAscii, m���ʽ
End Sub

