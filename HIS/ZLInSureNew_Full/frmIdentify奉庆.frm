VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk�����ʻ� 
      Caption         =   "�¸����ʻ�(&D)"
      Height          =   210
      Left            =   5100
      TabIndex        =   13
      Top             =   1305
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5190
      Visible         =   0   'False
      Width           =   2085
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   -225
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   6218
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
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "����(&R)"
      Height          =   350
      Left            =   165
      TabIndex        =   2
      Top             =   4065
      Width           =   1100
   End
   Begin VB.ComboBox cbo��� 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   4365
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3180
      Width           =   2310
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5625
      TabIndex        =   4
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4335
      TabIndex        =   3
      Top             =   4065
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   270
      Left            =   6420
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   2  'OFF
      Left            =   870
      TabIndex        =   1
      Top             =   3585
      Width           =   5820
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   34
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   33
      Top             =   3915
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ͳ������"
      Height          =   180
      Index           =   11
      Left            =   135
      TabIndex        =   28
      Top             =   3225
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4365
      TabIndex        =   21
      Top             =   2025
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4365
      TabIndex        =   8
      Top             =   855
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   870
      TabIndex        =   6
      Top             =   885
      Width           =   2310
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����˻�����Ϣ��ʾ������ͨ��������ť���½��ж�ȡIC����Ϣ��"
      Height          =   180
      Left            =   630
      TabIndex        =   35
      Top             =   360
      Width           =   5400
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify����.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   5
      Top             =   952
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��֤��"
      Height          =   180
      Index           =   1
      Left            =   3615
      TabIndex        =   7
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   9
      Top             =   1312
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   3975
      TabIndex        =   11
      Top             =   1305
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   4
      Left            =   135
      TabIndex        =   18
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   3615
      TabIndex        =   16
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Index           =   6
      Left            =   3615
      TabIndex        =   20
      Top             =   2085
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����(&T)"
      Height          =   180
      Index           =   7
      Left            =   3345
      TabIndex        =   30
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   495
      TabIndex        =   14
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   9
      Left            =   135
      TabIndex        =   22
      Top             =   2505
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   135
      TabIndex        =   26
      Top             =   2857
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   12
      Left            =   3615
      TabIndex        =   24
      Top             =   2475
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   870
      TabIndex        =   10
      Top             =   1260
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4365
      TabIndex        =   12
      Top             =   1260
      Width           =   525
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   870
      TabIndex        =   15
      Top             =   1650
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4365
      TabIndex        =   17
      Top             =   1650
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   870
      TabIndex        =   19
      Top             =   2040
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   870
      TabIndex        =   23
      Top             =   2430
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   855
      TabIndex        =   29
      Top             =   3150
      Width           =   2310
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   870
      TabIndex        =   27
      Top             =   2805
      Width           =   5805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4365
      TabIndex        =   25
      Top             =   2430
      Width           =   2310
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   225
      TabIndex        =   31
      Top             =   3645
      Width           =   630
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean
Private mstrArr             '���ͨ����������Ϣ 0 ��������,1 ���,2 ���� ....
Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk�����ʻ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd����_Click()
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "" & _
            "   Select id, ����, ����, ������,to_char(���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��" & _
            "   From ҽ������Ŀ¼"
                
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        With rsTemp
            If .EOF Then
                MsgBox "�������κβ���,�����أ�", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt����.Top - .Height
                    .Left = txt����.Left + txt����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = .Width - .ColWidth(1) - .ColWidth(2) - .ColWidth(3) - .ColWidth(4)
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt���� = "[" & Nvl(!����) & "]" & IIf(IsNull(!����), "", !����)
                txt����.Tag = Nvl(!ID)
                zlCommFun.PressKey vbKeyTab
            End If
        End With
End Sub

Private Sub cmd�鿨_Click()
    
   If ��ȡ�α���Ա��Ϣ_���� = False Then
        cmdȷ��.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmdȷ��.Enabled = True
End Sub

Private Sub Form_Activate()
    '
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If ��ȡ�α���Ա��Ϣ_���� = False Then
        Call ClearData
        cmdȷ��.Enabled = False
        Exit Sub
    End If
    Call LoadCtrlData
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
    Dim lng����ID  As Long
    
    lng����ID = Val(txt����.Tag)
    IsValid = False
    If mbytType = 4 Or mbytType = 1 Then
        '2008/04/01:��֮ԣҪ����Ժ����ѡ����
        If lng����ID = 0 Then
            ShowMsgbox "δ���벡��"
            If txt����.Enabled Then txt����.SetFocus
            Exit Function
        End If
    End If
    If Trim(g�������_����.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If cmd�鿨.Enabled Then cmd�鿨.SetFocus
        Exit Function
    End If
    
    If Trim(txt����) <> "" And Val(txt����.Tag) = 0 Then
        ShowMsgbox "����ѡ�����,������ѡ��!"
        txt����.SetFocus
        Exit Function
    End If
    If cbo���.Text = "" Then
        ShowMsgbox "֧�����δѡ��"
        Exit Function
    End If
    
    If �α��ʸ����_���� = False Then
        Exit Function
    End If
 
    'סԺ״̬�ж� 1 ��סԺ  2 ��������סԺ��  0 ��ʼ��ֵ��ͬ2��  4 תԺ
    If ��ȡסԺ״̬_����(lng״̬) = False Then
        Exit Function
    End If
    
    If lng״̬ = 1 And mbytType = 0 Then
        ShowMsgbox "��ǰ�����Ѿ���Ժ�����ܿ�����!"
        Exit Function
    End If
    If lng״̬ = 1 And mbytType = 1 Then
        ShowMsgbox "��ǰ�����Ѿ���Ժ,������Ժ�Ǽ�!"
        Exit Function
    End If
    
    If lng״̬ = 4 And InStr(1, cbo���.Text, "תԺ") = 0 Then
        ShowMsgbox "��ǰ����ΪתԺ����ѡ��ҽ�����ΪתԺ!"
        Exit Function
    End If
    
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ٲ׷���, g�������_����.ҽ��֤��)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
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
    Dim str��� As String
    Dim int��ǰ״̬ As Integer
    Dim lng״̬ As Long
    
    g�������_����.ҽ����� = Split(cbo���.Text, "-")(0)
    If cbo�������.ListIndex >= 0 Then
        g�������_����.������� = Split(cbo�������.Text, "-")(0)
    Else
        g�������_����.������� = ""
    End If
    lng����ID = Val(txt����.Tag)
    
    If IsValid = False Then Exit Sub
    
    Err = 0: On Error GoTo errHand:
    DebugTool "1.���в���ȷ��"
    If lng����ID <> 0 And txt����.Text <> "" Then
        g�������_����.���ֱ��� = Mid(txt����.Text, 2, InStr(1, txt����.Text, "]") - 2)
        g�������_����.�������� = Mid(txt����.Text, InStr(1, txt����.Text, "]") + 1)
    Else
        g�������_����.���ֱ��� = "000000"
    End If
    DebugTool "2.���в���ȷ���ɹ�"
    
    g�������_����.�¸����ʻ� = IIf(chk�����ʻ�.Value = 1, True, False)
    g�������_����.����ID = lng����ID
    
    g�������_����.ҽ����� = Mid(cbo���.Text, 1, InStr(1, cbo���.Text, "-") - 1)
    int��ǰ״̬ = 0
    'ȷ�������Ƿ�����
    If (mbytType = 0 Or mbytType = 3) And g�������_����.ҽ����� = "13" Then
        If ���ͨ����������Ϣ() = False Then Exit Sub
    End If
    
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�ٲ׷��� & " and  ҽ����='" & g�������_����.ҽ��֤�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            If mlng����ID <> Nvl(rsTemp!����ID, 0) And mlng����ID <> 0 Then
                ShowMsgbox "���ǵ�ǰ��Ҫ������û�"
                Exit Sub
            End If
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    DebugTool "3.��ʼ����������Ϣ��"
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_����
        
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .ҽ��֤��           '1ҽ����
        strIdentify = strIdentify & ";" & .ҽ�����                   '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤��           '6���֤
        strIdentify = strIdentify & ";" & .��λ���� & IIf(.��λ���� = 0, "", "(" & .��λ���� & ")")          '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";"                              '9.˳���
        strAddition = strAddition & ";"                                '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";" & IIf(lng����ID = 0, "", lng����ID)             '13����ID
        strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .������� & "|" & .��Ŀ���� & "|" & .��Ŀ����                               '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";" & .�������             '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                 '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    DebugTool "5.��ʼ���没��"
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�ٲ׷���)
    If mlng����ID = 0 Then Exit Sub
    DebugTool "����ɹ�"
    
    If mbytType = 0 Or mbytType = 3 Then
        '��ȡ�����
        Dim lng������� As Long
        Dim str������ˮ�� As String
        gstrSQL = "Select nvl(�������,0)+1 as ������� From �����ʻ� where ����ID=" & mlng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lng������� = Nvl(rsTemp!�������, 1)
        g�������_����.����� = mlng����ID & "_" & lng�������
        '���±����ʻ�
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ٲ׷��� & ",'�������','" & lng������� & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
        DebugTool "6.���еǼǴ���"
        '�Ƚ��еǼǴ���
        If ���˵ǼǴ���(str������ˮ��) = False Then
            Exit Sub
        End If
        
        DebugTool "�Ǽǳɹ�"
        '���潫������ˮ��
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ٲ׷��� & ",'˳���','''" & str������ˮ�� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���潻����ˮ��")
        DebugTool "7.���¾�����Ϣ:" & g�������_����.�ʻ����
        
        If ���¾�����Ϣ_����(0, strOutput) = False Then Exit Sub
        DebugTool "8.���¾�����Ϣ"
    Else
        gstrSQL = "Select * From �����ʻ� where ����ID=" & mlng����ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        With g�������_����
            .����� = mlng����ID & "_" & Nvl(rsTemp!�������, 0)
        End With
    End If
    g�������_����.����ID = mlng����ID
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    Unload Me
    Exit Sub
errHand:
    DebugTool "Err:" & Err.Description
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function ���ͨ����������Ϣ() As Boolean
    Dim StrInput As String, strOutput As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, strArr
    Err = 0
    On Error GoTo errHand:
    With g�������_����
'        strInput = .ҽ��֤�� & "|"
'        strInput = strInput & "|" & Split(cbo�������.Text, "-")(0)
'        If ҵ������_����(����_���ͨ����������Ϣ, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        If UBound(mstrArr) <= 1 Then
'            ShowMsgbox "û����ص�������Ϣ����"
'            Exit Function
'        End If
'
'        '��������
'        If InitTable(rsTemp, "�������|C|50||��Ŀ���|C|50||��Ŀ����|C|100") = False Then Exit Function
'        With rsTemp
'            For i = 0 To UBound(strArr) Step 3
'                .AddNew
'                !������� = strArr(i)
'                !��Ŀ���� = strArr(i + 1)
'                !��Ŀ���� = strArr(i + 2)
'                .Update
'            Next
'        End With
'        'ѡ��һ����Ч��������Ŀ
'        If frmListSel.ShowSelect(rsTemp, "�������", "��ѡ��һ����Ч��������Ŀ����", "����", False) = False Then
'            .������� = ""
'            .��Ŀ���� = ""
'            .��Ŀ���� = ""
'            Exit Function
'        End If
'        .������� = Nvl(rsTemp!�������)
'        .��Ŀ���� = Nvl(rsTemp!��Ŀ����)
'        .��Ŀ���� = Nvl(rsTemp!��Ŀ����)
    End With
    ���ͨ����������Ϣ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function InitTable(ByRef rsTemp As ADODB.Recordset, Optional ByVal strFields As String = "����|C|30||����|C|50") As Boolean
    Dim strArr, strArr1
    Dim i As Long
    Err = 0
    On Error GoTo errHand:
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        strArr = Split(strFields, "||")
        For i = 0 To UBound(strArr)
            strArr1 = Split(strArr(i), "|")
            Select Case strArr1(1)
            Case "C"
                .Fields.Append strArr1(0), adLongVarChar, Val(strArr1(2))
            Case "N"
                .Fields.Append strArr1(0), adDouble, Val(strArr1(2)), adFldIsNullable
            Case Else
            End Select
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    InitTable = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function ���˵ǼǴ���(ByRef str������ˮ�� As String) As Boolean
    '��������Ǽ�
    Dim StrInput As String, strOutput As String
    '�����ض��������ݣ�סԺ�������|ҽ��֤����|IC����|��Ժ����|��Ժ��������|������
    '                  סԺ�������|���˱��|IC����|ҽ�����|��Ժ����|��Ժ��������|����|������
    With g�������_����
        StrInput = .����� & "|"
        StrInput = StrInput & .ҽ��֤�� & "|"
        StrInput = StrInput & .���� & "|"
        StrInput = StrInput & .ҽ����� & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & "" & "|"
        StrInput = StrInput & gstrUserName
    End With
    DebugTool "���벡�˵Ǽ�"
    Err = 0
    On Error GoTo errHand:
    If ҵ������_����(����_���˵Ǽ�, StrInput, strOutput) = False Then Exit Function
    DebugTool "���˵Ǽǳɹ�"
    str������ˮ�� = Split(strOutput, "|")(2)
    
    DebugTool "���˵Ǽǳɹ�"
    ���˵ǼǴ��� = True
    Exit Function
errHand:
    DebugTool "Err:" & Err.Description
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    DebugTool "���������֤,����ʼ���������Ϣ"
    If Loadҽ����� = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    Call Load�������
    DebugTool "����ɹ�(�����֤)"
    
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
    Dim rsTemp As New ADODB.Recordset
    With g�������_����
        lblEdit(0).Caption = .����
        lblEdit(1).Caption = .ҽ��֤��
        lblEdit(2).Caption = .����
        lblEdit(3).Caption = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        
        lblEdit(4).Caption = .����
        lblEdit(5).Caption = .��������
        
        lblEdit(6).Caption = .���֤��
        lblEdit(7).Caption = .������
        lblEdit(8).Caption = .�������
        lblEdit(9).Caption = Format(.�ʻ����, "####0.00;#####0.00; ;")
        lblEdit(10).Caption = .��λ���� & IIf(.��λ���� <> "", "(" & .��λ���� & ")", "")
        lblEdit(11).Caption = .ͳ������
    End With
    
    gstrSQL = "Select ����ID,���� from �����ʻ� where ҽ����='" & g�������_����.ҽ��֤�� & "' and ����=" & TYPE_�ٲ׷���
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ��ز���"
    If rsTemp.EOF Then Exit Sub
    g�������_����.ҽ����� = Nvl(rsTemp!����)
    
    gstrSQL = "Select * From ҽ������Ŀ¼ where ID=" & Nvl(rsTemp!����ID, 0)
    If rsTemp.State = 1 Then rsTemp.Close
    Call SQLTest(App.ProductName, "��ȡ������Ϣ", gstrSQL)
    rsTemp.Open gstrSQL, gcnOracle_����
    Call SQLTest
    If rsTemp.EOF Then
        Exit Sub
    End If
    txt����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txt����.Tag = Nvl(rsTemp!ID, 0)
    Dim i As Long
    For i = 0 To cbo���.ListCount - 1
        If InStr(1, cbo���.List(i), g�������_����.ҽ����� & "-") <> 0 Then
            cbo���.ListIndex = i
            Exit For
        End If
    Next
    
End Sub
Private Sub Form_Load()
        'mblnFirst
        mblnFirst = True
End Sub

Private Sub mshSelect_Click()
    With mshSelect
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            SetColumnSort mshSelect, mintPreCol, mintsort
            Exit Sub
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

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshSelect
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .COL = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .COL = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
End Sub


'����ͷ��������
Private Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .COL = intCol
            .ColSel = intCol
            intTemp = .TextMatrix(.Row, 0)
            If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
               .Sort = flexSortStringNoCaseAscending
               intPreSort = flexSortStringNoCaseAscending
            Else
               .Sort = flexSortStringNoCaseDescending
               intPreSort = flexSortStringNoCaseDescending
            End If
            intPreCol = intCol
            .Row = FindRow(mshFilter, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .COL = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
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
            "   Select id, ����, ����, ������, to_char(���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��" & _
            "   From ҽ������Ŀ¼" & _
            "   Where " & zlCommFun.GetLike("", "����", Me.txt����) & " Or " & _
                        zlCommFun.GetLike("", "����", Me.txt����) & " Or " & _
                        zlCommFun.GetLike("", "������", Me.txt����)
                       
        
        With rsTemp
            .CursorLocation = adUseClient
            .Open gstrSQL, gcnOracle_����
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = txt����.Top - .Height
                    .Left = txt����.Left + txt����.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 2000
                    .ColWidth(3) = 1400
                    .ColWidth(4) = .Width - .ColWidth(1) - .ColWidth(2) - .ColWidth(3) - .ColWidth(4)
                    .Row = 1
                    .COL = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt���� = "[" & Nvl(!����) & "]" & IIf(IsNull(!����), "", !����)
                txt����.Tag = Nvl(!ID)
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
            txt����.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
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

Private Function Loadҽ�����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "" & _
        "   Select * From ҽ��ҽ����� " & _
        "   where " & IIf(mbytType = 0, "  nvl(��־,0)=0", "  nvl(��־,0)=1") & _
        "   Order by ����"
    Err = 0
    On Error GoTo errHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "ҽ��ҽ�����"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "����ҽ��ҽ���������HIS��Ӧ����ϵ!"
        Exit Function
    End If
    
    With rsTemp
        cbo���.Clear
        Do While Not .EOF
            cbo���.AddItem Nvl(!����) & "--" & Nvl(!����)
            .MoveNext
        Loop
    End With
    cbo���.ListIndex = 0
    Loadҽ����� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
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
    With g�������_����
        .���� = ""
        .�Ա� = ""
        .���֤�� = ""
        .�������� = ""
        .������ = ""
        .������� = ""
        .��λ���� = ""
        .��λ���� = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub

Private Sub Load�������()
        With cbo�������
            .Clear
            .AddItem "1-������"
            .ListIndex = .NewIndex
            .AddItem "2-��������"
            .AddItem "3-��������"
            .AddItem "4-�������Բ�"
            .AddItem "5-ת��"
            .AddItem "6-תԺ"
            .AddItem "7-����ҩƷ"
        End With
End Sub
