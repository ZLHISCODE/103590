VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify�˳� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo��Ժ��� 
      Height          =   300
      Left            =   7110
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   4275
      Width           =   1155
   End
   Begin VB.ComboBox cboסԺ��� 
      Height          =   300
      Left            =   5145
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   4275
      Width           =   1170
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   4410
      Left            =   -300
      TabIndex        =   59
      Top             =   5715
      Visible         =   0   'False
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   7779
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cbo��Ժ��� 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   4275
      Width           =   3090
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   7995
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4665
      Width           =   255
   End
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "���¶���(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   18
      Top             =   5310
      Width           =   1305
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7110
      TabIndex        =   61
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   19
      Top             =   615
      Width           =   8475
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -540
      TabIndex        =   17
      Top             =   5055
      Width           =   9030
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   840
      TabIndex        =   56
      Top             =   4650
      Width           =   7425
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5955
      TabIndex        =   60
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "��Ժ���"
      Height          =   180
      Index           =   1
      Left            =   6360
      TabIndex        =   53
      Top             =   4335
      Width           =   720
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      Caption         =   "סԺ���"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   51
      Top             =   4335
      Width           =   720
   End
   Begin VB.Label lblInfor 
      Caption         =   "��Ժ���"
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   49
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   5130
      TabIndex        =   58
      Top             =   1005
      Width           =   3135
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   435
      TabIndex        =   55
      Top             =   4710
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������ֹ����"
      Height          =   180
      Index           =   22
      Left            =   4020
      TabIndex        =   48
      Top             =   3945
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   22
      Left            =   5130
      TabIndex        =   47
      ToolTipText     =   "���Բ���Ч�ڽ�ֹ����"
      Top             =   3885
      Width           =   3135
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   21
      Left            =   75
      TabIndex        =   46
      Top             =   3945
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   21
      Left            =   840
      TabIndex        =   45
      Top             =   3885
      Width           =   3075
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���Ա�ʶ"
      Height          =   180
      Index           =   20
      Left            =   6360
      TabIndex        =   44
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   20
      Left            =   7125
      TabIndex        =   43
      ToolTipText     =   "���Բ����߱�ʶ"
      Top             =   3495
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Index           =   19
      Left            =   4380
      TabIndex        =   42
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   19
      Left            =   5130
      TabIndex        =   41
      ToolTipText     =   "סԺ��Ϣ��������"
      Top             =   3495
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Index           =   18
      Left            =   2160
      TabIndex        =   40
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   18
      Left            =   2880
      TabIndex        =   39
      ToolTipText     =   "����סԺ����"
      Top             =   3495
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   17
      Left            =   75
      TabIndex        =   38
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   17
      Left            =   840
      TabIndex        =   37
      ToolTipText     =   "�ϴ������ҽ����"
      Top             =   3495
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����ۼ�"
      Height          =   180
      Index           =   15
      Left            =   4380
      TabIndex        =   36
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   16
      Left            =   7125
      TabIndex        =   35
      ToolTipText     =   "��������ʻ�֧���ۼ�"
      Top             =   3105
      Width           =   1140
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ��ۼ�"
      Height          =   180
      Index           =   16
      Left            =   6360
      TabIndex        =   34
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   5130
      TabIndex        =   33
      ToolTipText     =   "�������Բ������ۼ�"
      Top             =   3105
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������ۼ�"
      Height          =   180
      Index           =   14
      Left            =   1980
      TabIndex        =   32
      Top             =   3150
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   2880
      TabIndex        =   31
      ToolTipText     =   "�����ν���ۼ�"
      Top             =   3105
      Width           =   1020
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�����ۼ�"
      Height          =   180
      Index           =   13
      Left            =   75
      TabIndex        =   30
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   840
      TabIndex        =   29
      ToolTipText     =   "����ͳ�����֧���ۼ�"
      Top             =   3105
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   12
      Left            =   7110
      TabIndex        =   28
      ToolTipText     =   "�ϴγ俨����"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   5130
      TabIndex        =   27
      ToolTipText     =   "�ۼƻ�������ʻ����"
      Top             =   2693
      Width           =   1110
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�俨����"
      Height          =   180
      Index           =   12
      Left            =   6360
      TabIndex        =   26
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ۼƻ����ʻ�"
      Height          =   180
      Index           =   11
      Left            =   4020
      TabIndex        =   25
      Top             =   2745
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   24
      Top             =   2693
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   840
      TabIndex        =   23
      Top             =   2693
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�α�����"
      Height          =   180
      Index           =   10
      Left            =   2160
      TabIndex        =   22
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ձ��"
      Height          =   180
      Index           =   9
      Left            =   75
      TabIndex        =   21
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��Ա���"
      Height          =   180
      Index           =   3
      Left            =   4380
      TabIndex        =   3
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   8
      Left            =   4380
      TabIndex        =   15
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   8
      Left            =   5130
      TabIndex        =   16
      Top             =   2265
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1005
      Width           =   3075
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����˻�����Ϣ��ʾ������ͨ��[���¶���]��ť���½��ж�ȡ���˻�����Ϣ��"
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   20
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify�˳�.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   435
      TabIndex        =   0
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ᱣ�Ϻ�"
      Height          =   180
      Index           =   1
      Left            =   4230
      TabIndex        =   5
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   435
      TabIndex        =   2
      Top             =   1515
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   4
      Left            =   435
      TabIndex        =   7
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   7
      Left            =   75
      TabIndex        =   13
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   6
      Left            =   4380
      TabIndex        =   11
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   5
      Left            =   2520
      TabIndex        =   9
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   1425
      Width           =   3075
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   5130
      TabIndex        =   4
      Top             =   1425
      Width           =   3135
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   1838
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   10
      Top             =   1838
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   840
      TabIndex        =   14
      Top             =   2265
      Width           =   3075
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   5130
      TabIndex        =   12
      Top             =   1830
      Width           =   3135
   End
End
Attribute VB_Name = "frmIdentify�˳�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����,88-��סԺ��Ϣ���в�ѯ

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean

Private Sub cbo��Ժ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub


Private Sub cbo��Ժ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cboסԺ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd�鿨_Click()
    If ��ȡ�α���Ա��Ϣ = False Then
         cmdȷ��.Enabled = False
         Call ClearData
         Exit Sub
     End If
     Call LoadCtrlData
     cmdȷ��.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmdȷ��.Enabled = False
    Call LoadBase
    If ��ȡ�α���Ա��Ϣ = False Then
         Call ClearData
         Exit Sub
     End If
     Call LoadCtrlData
     Call InitCtlData
    cmdȷ��.Enabled = True
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim lng״̬ As Long
    
    IsValid = False
    If Trim(g�������_�˳�.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If cmd�鿨.Enabled Then cmd�鿨.SetFocus
        Exit Function
    End If
    
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '����鵱ǰ״̬
        Else
           '�º�����20051231�޸����ӣ�����ҽ������������ؾ�ҽ:Ҫ����ؿ������ڱ��ذ���סԺ�Ǽ�
            If mbytType = 1 Then
               If g�������_�˳�.��ؿ���־ = "1" Then
                  MsgBox "�ÿ�����ؿ��������ڱ���סԺ��", vbOKOnly + vbExclamation, gstrSysName
                  Exit Function
               End If
            End If
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�˳ɺ˹�ҵ, g�������_�˳�.��ᱣ�Ϻ�)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
         '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�˳ɺ˹�ҵ & " and  ҽ����='" & g�������_�˳�.��ᱣ�Ϻ� & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
        Else
            mlng����ID = 0
        End If
        mstrReturn = mlng����ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function
Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim StrInput  As String, strOutput As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim int��ǰ״̬ As Integer
    Dim lng״̬ As Long
    Dim str�������� As String
    
    
    If IsValid = False Then Exit Sub
    
    int��ǰ״̬ = 0
    
    
    '�º�����20050310�޸ģ����ڽӿ��ĵ�������Ҫ����YB_HHMD���ṹ����ֶ��ɡ�yyjb���޸�Ϊ��kh��
    
    gstrSQL = "Select * From YB_HHMD where kh='" & g�������_�˳�.IC���� & "'"
    
    rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�
    If rsTemp.EOF Then
        g�������_�˳�.��״̬ = "a"     '����
    Else
        Select Case Val(Nvl(rsTemp!Kzt))
        Case 0 '��ʧ
            ShowMsgbox "ע�⣺" & vbCrLf & "�ÿ��Ѿ�����ʧ!"
        Case 1 'Ƿ��
            ShowMsgbox "ע�⣺" & vbCrLf & "�ÿ��Ѿ�Ƿ��!"
        Case 2 'ͣ��
            ShowMsgbox "ע�⣺" & vbCrLf & "�ÿ��Ѿ�ͣ��!"
        Case 3 '����
            ShowMsgbox "ע�⣺" & vbCrLf & "�ÿ��Ѿ�����!"
        End Select
        g�������_�˳�.��״̬ = Val(Nvl(rsTemp!Kzt))
    End If
    If rsTemp.State = 1 Then rsTemp.Close
        
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�˳ɺ˹�ҵ & " and  ҽ����='" & g�������_�˳�.��ᱣ�Ϻ� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    
    If txt����.Tag <> "" Then
        g�������_�˳�.���ֱ��� = txt����.Tag
        str�������� = Split(txt����.Text & "]", "]")(1)
    Else
        g�������_�˳�.���ֱ��� = ""
        str�������� = ""
    End If
    
    If mbytType <> 1 And mbytType <> 4 Then
        g�������_�˳�.��Ժ��� = ""
        g�������_�˳�.סԺ��� = ""
        g�������_�˳�.��Ժ��� = ""
    Else
        g�������_�˳�.��Ժ��� = cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex)
        g�������_�˳�.סԺ��� = cboסԺ���.ItemData(cboסԺ���.ListIndex)
        g�������_�˳�.��Ժ��� = cbo��Ժ���.ItemData(cbo��Ժ���.ListIndex)
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�˳�
        strIdentify = .IC����                                '0����
        strIdentify = strIdentify & ";" & .��ᱣ�Ϻ�             '1ҽ����
        strIdentify = strIdentify & ";"                    '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & .�Ա�             '4�Ա�
        strIdentify = strIdentify & ";" & .��������                 '5��������
        strIdentify = strIdentify & ";" & .���֤��              '6���֤
        strIdentify = strIdentify & ";"          '7.��λ����(����)
        strAddition = ";0"                      '8.���Ĵ���
        strAddition = strAddition & ";" & .���ձ��                                '9.˳���
        strAddition = strAddition & ";" & .��Ժ���                                    '10��Ա���
        strAddition = strAddition & ";" & .�����ʻ����                   '11�ʻ����

        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";"             '13����ID
        strAddition = strAddition & ";1"                        '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .��Ա���           '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                         '17�Ҷȼ�
        strAddition = strAddition & ";" & .�����ʻ����                            '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�˳ɺ˹�ҵ)
    
    If mlng����ID = 0 Then Exit Sub

    If mbytType = 1 Or mbytType = 4 Or mbytType = 0 Then
        '���¸�����Ϣ(���̲�������:)
        '   ����ID_IN,����ͳ�����֧���ۼ�,�����ν���ۼ�,�α�����,ͳ����𹲸��ν���ۼ�,�������Բ������ۼ�,ͳ��֧���ۼ�,
        '   �ۼƻ�������ʻ����,�ϴγ俨����,��������ʻ�֧���ۼ�,��ǰ���,���ʻ������ͬ,
        '   �ϴ������ҽ����,����סԺ����,סԺ��Ϣ��������,���Բ����߱�ʶ,���Բ�����,���Բ���Ч�ڽ�ֹ����,���ִ���,��������
        gstrSQL = "ZL_ҽ�����˸�����Ϣ_UPDATE("
        gstrSQL = gstrSQL & "" & mlng����ID & ","
        gstrSQL = gstrSQL & "'" & g�������_�˳�.����ͳ��֧�� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�˳�.�������ۼ� & "',"
        gstrSQL = gstrSQL & "" & IIf(IsDate(g�������_�˳�.�α�����), "to_Date('" & g�������_�˳�.�α����� & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.ͳ�ﹲ�����ۼ� & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.���Բ������ۼ� & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.ͳ��֧���ۼ� & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.�ۼƻ�������ʻ� & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g�������_�˳�.�ϴγ俨����), "to_Date('" & g�������_�˳�.�ϴγ俨���� & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.�ʻ�֧���ۼ� & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.��ǰ��� & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g�������_�˳�.�ϴ���������), "to_Date('" & g�������_�˳�.�ϴ��������� & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "" & g�������_�˳�.����סԺ���� + 1 & ","
        gstrSQL = gstrSQL & "" & IIf(IsDate(g�������_�˳�.סԺ��������), "to_Date('" & g�������_�˳�.סԺ�������� & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "'" & g�������_�˳�.������ʶ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�˳�.���Բ����� & "',"
        gstrSQL = gstrSQL & "" & IIf(IsDate(g�������_�˳�.������Ч����), "to_Date('" & g�������_�˳�.������Ч���� & "','yyyy-mm-dd')", "NULL") & ","
        gstrSQL = gstrSQL & "'" & g�������_�˳�.���ֱ��� & "',"
        gstrSQL = gstrSQL & "'" & str�������� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�˳�.סԺ��� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�˳�.��Ժ��� & "')"
        ExecuteProcedure_�˳� "ҽ�����˸�����Ϣ"
    Else
    End If
    'g�������_�˳�.����ID = mlng����ID
    If mbytType = 4 Then
    
    '�º�����20050311�޸ģ�������
    'gstrSQL = "Select AF21,AF22 From ҽ�����˸�����Ϣ where ����id=mlng����ID "
        
      gstrSQL = "Select AF21,AF22 From ҽ�����˸�����Ϣ where ����id=" & mlng����ID
      
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��صĳ�Ժ��Ϣ"
        If rsTemp.EOF Then
            g�������_�˳�.���ҽԺ = ""
            g�������_�˳�.���ҽԺ���� = ""
        Else
            g�������_�˳�.���ҽԺ = Nvl(rsTemp!AF21)
            g�������_�˳�.���ҽԺ���� = Nvl(rsTemp!AF22)
        End If
    Else
        g�������_�˳�.���ҽԺ = ""
        g�������_�˳�.���ҽԺ���� = ""
    End If
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    g�������_�˳�.����ID = mlng����ID
    
    '�º�����20050402��ӣ���Ϊ���Բ����������Է����Բ����о�ҽ����
    If mbytType = 0 And g�������_�˳�.������ʶ = "1" Then
     If MsgBox("�û����Ƿ������Բ���ʽ����ҽ�����㣿", vbOKCancel, "�������") = vbOK Then
        blnmxb = True
     Else
        blnmxb = False
     End If
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    With g�������_�˳�
        lblEdit(0).Caption = .IC����
        lblEdit(1).Caption = .��ᱣ�Ϻ�
        lblEdit(2).Caption = .����
        lblEdit(3).Caption = Decode(.��Ա���, "10", "����Ա", "11", "��ҵ��Ա", "12", "��ҵ��Ա", "13", "�¸���Ա", "14", "������Ա", "15", "ͣн��ְ", "16", "�ż���Ա", "20", "������Ա", "δ֪")

        lblEdit(4).Caption = .�Ա�
        lblEdit(5).Caption = .����     '����
        lblEdit(6).Caption = Format(.�����ʻ����, "####0.00;-####0.00;;")
        lblEdit(7).Caption = .���֤��
        lblEdit(8).Caption = .��������      '��������
        lblEdit(9).Caption = .���ձ��
        lblEdit(10).Caption = .�α�����
        lblEdit(11).Caption = Format(.�ۼƻ�������ʻ�, "####0.00;-####0.00;;")
        lblEdit(12).Caption = .�ϴγ俨����

        lblEdit(13).Caption = Format(.����ͳ��֧��, "####0.00;-####0.00;;")
        lblEdit(14).Caption = Format(.�������ۼ�, "####0.00;-####0.00;;")
        lblEdit(15).Caption = Format(.���Բ������ۼ�, "####0.00;-####0.00;;")
        lblEdit(16).Caption = Format(.�ʻ�֧���ۼ�, "####0.00;-####0.00;;")
        lblEdit(17).Caption = .�ϴ���������
        lblEdit(18).Caption = .����סԺ����
        lblEdit(19).Caption = .סԺ��������
        lblEdit(20).Caption = .������ʶ
        lblEdit(21).Caption = .���Բ�����
        lblEdit(22).Caption = .������Ч����
    End With
End Sub

Private Sub Form_Load()
        mblnFirst = True
End Sub


Private Function ��ȡ�α���Ա��Ϣ() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim StrInput As String, strOutput As String
    Dim strArr As Variant, strArr1 As Variant
    
    ��ȡ�α���Ա��Ϣ = False
    Err = 0:    On Error GoTo errHand:
        
    If ҵ������_�˳�(�˳�_��ȡ�ֿ�����Ϣ, "", strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    '       IC����|������ݺ���|����|�Ա�|ҽ�Ʋα���Ա���|�����ʻ����|���ձ��|����ͳ�����֧���ۼ�|�����ν���ۼ�|��ؿ���־
    strArr = Split(strOutput, "|")
    If ҵ������_�˳�(�˳�_JbylReadIC, "", strOutput) = False Then
        Call ClearData
    End If
    '       ��ᱣ�Ϻ�|����|��Ա���|�α�����|ͳ����𹲸��ν���ۼ�|�������Բ������ۼ�|ͳ��֧���ۼ�|�ۼƻ�������ʻ����|�ϴγ俨����|��������ʻ�֧���ۼ�|��ǰ���|�ϴ������ҽ����|����סԺ����|סԺ��Ϣ��������|���Բ����߱�ʶ|���Բ�����|���Բ���Ч�ڽ�ֹ����
    strArr1 = Split(strOutput, "|")
    
    With g�������_�˳�
            .IC���� = strArr(0)
            .��ᱣ�Ϻ� = strArr1(0)
            .���֤�� = strArr(1)
            .���� = strArr(2)
            .�Ա� = Decode(strArr(3), "1", "��", "0", "Ů", "9", "δ֪", strArr(3))
            .��Ա��� = strArr(4)
            .�����ʻ���� = Val(strArr(5)) / 100
            .���ձ�� = strArr(6)
            .����ͳ��֧�� = Val(strArr(7)) / 100       '����ͳ�����֧���ۼ�
            .�������ۼ� = Val(strArr(8)) / 100         '�����ν���ۼ�
            
            '�º�����20051231�޸ģ�����ҽ�����ĵ����ӿ��ĵ���������ؾ�ҽ��
            
            .��ؿ���־ = strArr(9)
            
            .�α����� = zlCommFun.AddDate(strArr1(3))
            .ͳ�ﹲ�����ۼ� = Val(strArr1(4)) / 100   'ͳ����𹲸��ν���ۼ�
            .���Բ������ۼ� = Val(strArr1(5)) / 100   '�������Բ������ۼ�
            .ͳ��֧���ۼ� = Val(strArr1(6)) / 100 'ͳ��֧���ۼ�
            .�ۼƻ�������ʻ� = Val(strArr1(7)) / 100    '�ۼƻ�������ʻ����
            .�ϴγ俨���� = zlCommFun.AddDate(strArr1(8))
            .�ʻ�֧���ۼ� = Val(strArr1(9)) / 100     '��������ʻ�֧���ۼ�
            .��ǰ��� = Val(strArr1(10)) / 100
            .�ϴ��������� = zlCommFun.AddDate(strArr1(11))     '�ϴ������ҽ����
            .����סԺ���� = Val(strArr1(12))
            .סԺ�������� = zlCommFun.AddDate(strArr1(13))    'סԺ��Ϣ��������
            .������ʶ = strArr1(14)          '���Բ����߱�ʶ
            If InStr(1, strArr1(15), "1") = 0 Then
               .���Բ����� = strArr1(15)
            Else
             .���Բ����� = Lpad(InStr(1, strArr1(15), "1"), 9, "0") '���Բ�����
            End If
            .������Ч���� = zlCommFun.AddDate(strArr1(16))      '���Բ���Ч�ڽ�ֹ����
            If Trim(.���֤��) = "" Then
                .�������� = ""
                .���� = 0
            Else
                .�������� = zlCommFun.GetIDCardDate(.���֤��)
                .���� = Get����(.��������)
            End If
    End With
    ��ȡ�α���Ա��Ϣ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '��������Ϣ
    With g�������_�˳�
        .IC���� = ""
        .��ᱣ�Ϻ� = ""
        .���֤�� = ""
        .���� = ""
        .�Ա� = ""
        .��Ա��� = ""
        .�����ʻ���� = 0
        .���ձ�� = ""
        .����ͳ��֧�� = 0        '����ͳ�����֧���ۼ�
        .�������ۼ� = 0          '�����ν���ۼ�
        .�α����� = ""
        .ͳ�ﹲ�����ۼ� = 0    'ͳ����𹲸��ν���ۼ�
        .���Բ������ۼ� = 0    '�������Բ������ۼ�
        .ͳ��֧���ۼ� = 0 'ͳ��֧���ۼ�
        .�ۼƻ�������ʻ� = 0     '�ۼƻ�������ʻ����
        .�ϴγ俨���� = ""
        .�ʻ�֧���ۼ� = 0      '��������ʻ�֧���ۼ�
        .��ǰ��� = 0
        .�ϴ��������� = ""     '�ϴ������ҽ����
        .����סԺ���� = 0
        .סԺ�������� = ""    'סԺ��Ϣ��������
        .������ʶ = ""          '���Բ����߱�ʶ
        .���Բ����� = ""       '���Բ�����
        .������Ч���� = ""     '���Բ���Ч�ڽ�ֹ����
        .���� = 0
        .�������� = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub LoadBase()
    '��������
    Me.cbo��Ժ���.Clear
    Me.cboסԺ���.Clear
    Me.cbo��Ժ���.Clear
    With Me.cbo��Ժ���
        .AddItem "1-������Ժ"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-����ת��"
        .ItemData(.NewIndex) = 2
        .AddItem "3-����ת��"
        .ItemData(.NewIndex) = 3
        .AddItem "4-�����Բ����ص�һ��סԺ"
        .ItemData(.NewIndex) = 4
    End With
    With Me.cboסԺ���
        .AddItem "0-����סԺ"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-��������"
        .ItemData(.NewIndex) = 1
    End With
    With Me.cbo��Ժ���
        .AddItem "1-������Ժ"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "2-ת������"
        .ItemData(.NewIndex) = 2
        .AddItem "3-ת������"
        .ItemData(.NewIndex) = 3
    End With
    If mbytType = 1 Or mbytType = 4 Then
        Me.cbo��Ժ���.Enabled = mbytType = 1
        Me.cboסԺ���.Enabled = True
        Me.cbo��Ժ���.Enabled = mbytType <> 1
        If Me.cbo��Ժ���.Enabled Then
            Me.cbo��Ժ���.ListIndex = -1
        End If
    Else
        Me.cbo��Ժ���.ListIndex = -1: Me.cbo��Ժ���.Enabled = False
        Me.cboסԺ���.ListIndex = -1: Me.cboסԺ���.Enabled = False
        Me.cbo��Ժ���.ListIndex = -1: Me.cbo��Ժ���.Enabled = False
    End If
   ' Me.cbo��Ժ���.Enabled = False
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select jbdm ��������,jbmc ��������,case isnull(jblx,'0') when '0' then '��ͨ��' else '���Բ�' end  �������� " & _
            "   From YB_BZML "
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnSQLSEVER_�˳�
     
     With rsTemp
         If .EOF Then
             MsgBox "�������κβ���,�����أ�", vbInformation, gstrSysName
             Exit Sub
         End If
         
         If .RecordCount > 1 Then
             Set mshSelect.Recordset = rsTemp
             With mshSelect
                 .Cols = 3
                 .Top = txt����.Top - .Height
                 .Left = txt����.Left + txt����.Width - .Width
                 .Visible = True
                 .SetFocus
                 .ColWidth(0) = 2000
                 .ColWidth(1) = 3000
                 .ColWidth(2) = .Width - .ColWidth(1)
                 .Row = 1
                 .COL = 0
                 .ColSel = .Cols - 1
                 Exit Sub
             End With
         Else
             txt���� = "[" & Nvl(!��������) & "]" & IIf(IsNull(!��������), "", !��������)
             txt����.Tag = Nvl(!��������)
             zlCommFun.PressKey vbKeyTab
         End If
     End With

End Sub

Private Sub mshSelect_DblClick()
    With mshSelect
        If .Row > 0 And .TextMatrix(.Row, 0) <> "" Then
            mshSelect_KeyPress 13
        End If
    End With
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    OpenIme GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSQL As String
      
    If KeyCode = vbKeyReturn Then
        If Me.txt���� = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        If Trim(txt����) = "" Then Exit Sub
        If Trim(txt����.Tag) <> "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        txt���� = UCase(txt����)
         Dim rsTemp As New ADODB.Recordset
        gstrSQL = "" & _
                " Select jbdm ��������,jbmc ��������,case isnull(jblx,'0') when '0' then '��ͨ��' else '���Բ�' end  �������� " & _
                " From YB_BZML " & _
                " Where " & zlCommFun.GetLike("", "jbdm", Me.txt����) & " Or " & _
                        zlCommFun.GetLike("", "jbmc", Me.txt����)

        With rsTemp
            .CursorLocation = adUseClient
            .Open gstrSQL, gcnSQLSEVER_�˳�
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Cols = 3
                    .Top = txt����.Top - .Height
                    .Left = txt����.Left + txt����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 2000
                    .ColWidth(1) = 3000
                    .ColWidth(2) = .Width - .ColWidth(1) - 30
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt���� = "[" & Nvl(!��������) & "]" & IIf(IsNull(!��������), "", !��������)
                txt����.Tag = Nvl(!��������)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
    End If
End Sub

Private Sub txt����_LostFocus()
    OpenIme ""
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            txt����.Text = "[" & .TextMatrix(.Row, 0) & "]" & .TextMatrix(.Row, 1)
            txt����.Tag = .TextMatrix(.Row, 0)
            If cmdȷ��.Enabled Then cmdȷ��.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub
'Ѱ����ĳһ��Ԫֵ��ȵ���
Private Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .Rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

Private Sub InitCtlData()
    '��ʼ�ؼ�����
    Dim i As Integer
    Dim str��Ժ��� As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng����ID As Long
    gstrSQL = "Select ����id,��Ա��� From �����ʻ� where ҽ����='" & g�������_�˳�.��ᱣ�Ϻ� & "'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then Exit Sub
    
    str��Ժ��� = Nvl(rsTemp!��Ա���, "0")
    lng����ID = Nvl(rsTemp!����ID, 0)
    If lng����ID = 0 Then Exit Sub
    gstrSQL = "Select AF17,AF18,AF19,AF20 From ҽ�����˸�����Ϣ where ����id=" & lng����ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then Exit Sub
    
    'ȷ����Ժ���:
    If mbytType = 1 Or mbytType = 4 Then
        For i = 0 To cbo��Ժ���.ListCount - 1
            If cbo��Ժ���.ItemData(i) = str��Ժ��� Then
               cbo��Ժ���.ListIndex = i: Exit For
            End If
        Next
        For i = 0 To cboסԺ���.ListCount - 1
            If cboסԺ���.ItemData(i) = Nvl(rsTemp!AF19, "0") And cboסԺ���.Enabled Then
               cboסԺ���.ListIndex = i: Exit For
            End If
        Next
        For i = 0 To cbo��Ժ���.ListCount - 1
            If cbo��Ժ���.ItemData(i) = Nvl(rsTemp!AF20, "0") And cbo��Ժ���.Enabled Then
               cbo��Ժ���.ListIndex = i: Exit For
            End If
        Next
    End If
    
    'ȷ����ز���
    Me.txt����.Text = IIf(Nvl(rsTemp!AF17) = "", "", "[" & Nvl(rsTemp!AF17) & "]" & Nvl(rsTemp!AF18))
    Me.txt����.Tag = Nvl(rsTemp!AF17)
End Sub
