VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm������Ŀ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "frm������Ŀ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "�ύ����(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7620
      TabIndex        =   15
      Top             =   6750
      Width           =   1545
   End
   Begin VB.CommandButton cmdȫ�������� 
      Caption         =   "ȫ��������(&N)"
      Height          =   465
      Left            =   5760
      TabIndex        =   14
      Top             =   6750
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CheckBox chk��ʾ������Ŀ 
      Caption         =   "��ʾ�������Ŀ(&A)"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   7290
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1875
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   9450
      TabIndex        =   16
      Top             =   1410
      Width           =   1155
   End
   Begin VB.CommandButton cmdδ�����Ŀ��ѯ 
      Caption         =   "δ�����Ŀ��ѯ(&R)"
      Height          =   465
      Left            =   360
      TabIndex        =   11
      Top             =   6750
      Width           =   1695
   End
   Begin VB.CommandButton cmdȫ��תΪ�Է� 
      Caption         =   "ȫ��תΪ�Է�(&F)"
      Height          =   465
      Left            =   4080
      TabIndex        =   13
      Top             =   6750
      Width           =   1545
   End
   Begin VB.CommandButton cmdȫ�����ͨ�� 
      Caption         =   "ȫ�����ͨ��(&V)"
      Height          =   465
      Left            =   2400
      TabIndex        =   12
      Top             =   6750
      Width           =   1545
   End
   Begin TabDlg.SSTab TabList 
      Height          =   6375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "���շ���Ŀ(&1)"
      TabPicture(0)   =   "frm������Ŀ����.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Bill(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ѪҺ�׵���(&2)"
      TabPicture(1)   =   "frm������Ŀ����.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5265
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   9287
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5265
         Index           =   1
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   9287
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���˻�����Ϣ"
      Height          =   735
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6780
         TabIndex        =   6
         Top             =   270
         Width           =   2085
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt�������� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lblҽ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6150
         TabIndex        =   5
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblסԺ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3270
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm������Ŀ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�����޸��嵥��zl9Insure.vbp��frm������Ŀ������frmҽ���ʻ���mdl����
Private mintInsure As Integer
Private mlng����ID As Long, mlng��ҳID As Long
Private Enum ColDefine
    Col_����ID
    Col_��Ŀ��Ϣ
    Col_���
    Col_����
    Col_���
    Col_������ˮ��
    Col_��˱�־
    Col_Count
End Enum
Private Enum BillDefine '��ͬ��tabList.tab
    ���շ���Ŀ
    ѪҺ�׵���
End Enum
Private Enum Marker
    ������
    ���ͨ��
    תΪ�Է�
End Enum

Private Const gstr�������� As String = "21"
'���:סԺ��|������ϸ��ˮ��|�������
'����:��1������ʧ�ܣ�1�������ɹ���
'������ǣ�0����ʾ�ô�����ϸ��ˮ�Ŷ�Ӧ�Ĵ�������Ŀδͨ�����������Է���Ŀ����1����ʾ�ô�������Ŀͨ�����������ð�Ŀ¼�й涨�ı������뱾�ν���ҽ���ѡ�
Private Const gstrδ�����Ŀ��ѯ As String = "22"
'���:סԺ��|��������|����
'����:δ������Ŀ����|������ϸ��ˮ��1|������ϸ��ˮ��2|��|������ϸ��ˮ��n����෵��60������������ʵ�ʵ�����
'�������ͣ�1�����շ�������2ѪҺ�׵�������
'���ڣ���ʽΪYYYYMMDD

Public Sub ShowMe(ByVal intInsure As Integer)
    mintInsure = intInsure
    Me.Show 1
End Sub

Private Sub Bill_DblClick(Index As Integer, Cancel As Boolean)
    With Bill(Index)
        If Trim(.TextMatrix(.Row, Col_������ˮ��)) = "" Then Exit Sub
        
        If .TextMatrix(.Row, Col_��˱�־) = "" Then
            .TextMatrix(.Row, Col_��˱�־) = "��"
        ElseIf .TextMatrix(.Row, Col_��˱�־) = "��" Then
            .TextMatrix(.Row, Col_��˱�־) = "��"
        Else
            .TextMatrix(.Row, Col_��˱�־) = "��"      'ֻ����������״̬���л�����Ҫ���û�����ȡ���ļ���
        End If
    End With
End Sub

Private Sub chk��ʾ������Ŀ_Click()
    '��ȡ������Ŀ����������ˮ��˳����ʾ
    
    If txtҽ����.Text = "" Then
        MsgBox "����ȷ�����ˣ�", vbInformation, gstrSysName
        txtסԺ��.SetFocus
        Exit Sub
    End If
    
    Call ShowData
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdȫ��������_Click()
    Call BillMarker(������)
End Sub

Private Sub cmdȫ�����ͨ��_Click()
    Call BillMarker(���ͨ��)
End Sub

Private Sub cmdȫ��תΪ�Է�_Click()
    Call BillMarker(תΪ�Է�)
End Sub

Private Sub cmdȷ��_Click()
    Dim objTarget As BillEdit
    Dim str������ˮ�� As String
    Dim intVerify As Integer, strInput As String, OutputData
    Dim intTab As Integer, lngRow As Long, lngRows As Long
    On Error GoTo errHand
    
    If MsgBox("��ȷ�����������������ϴ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '����HIS���ݿ⣬ͬ������ҽ��
    For intTab = ���շ���Ŀ To ѪҺ�׵���
        Set objTarget = Bill(intTab)
        lngRows = objTarget.Rows - 1
        For lngRow = 1 To lngRows
            str������ˮ�� = objTarget.TextMatrix(lngRow, Col_������ˮ��)
            If objTarget.TextMatrix(lngRow, Col_��˱�־) = "��" Then
                intVerify = 1
            ElseIf objTarget.TextMatrix(lngRow, Col_��˱�־) = "��" Then
                intVerify = 0
            Else
                intVerify = -1
            End If
            
            '��������
            If intVerify <> -1 And str������ˮ�� <> "" Then
                If mintInsure = TYPE_������ Then
                    strInput = "21|" & GetIdentify(mlng����ID, mlng��ҳID) & "|" & str������ˮ�� & "|" & intVerify
                    If HandleBusiness(strInput, OutputData) Then
                        gstrSQL = "zlYB_�����Ŀ��_UPDATE(" & intTab + 1 & "," & mlng����ID & "," & mlng��ҳID & ",'" & str������ˮ�� & "'," & intVerify + 1 & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ŀ��")
                    End If
                Else
                    gstrSQL = "Update �м��_������ϸ Set ���շ��������='" & IIf(intVerify = 1, "11", "00") & "' Where ������ˮ��='" & str������ˮ�� & "'"
                    gcn����������.Execute gstrSQL
                    
                    gstrSQL = "zlYB_�����Ŀ��_UPDATE(" & intTab + 1 & "," & mlng����ID & "," & mlng��ҳID & ",'" & str������ˮ�� & "'," & intVerify + 1 & ")"
                    gcn����������.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        Next
    Next
    
    Call ShowData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdδ�����Ŀ��ѯ_Click()
    Dim strInput As String
    Dim lngCount As Long, lngMin As Long, LNGMAX As Long
    Dim OutputData
    
    If Me.txtסԺ��.Text = "" Then
        MsgBox "����ȷ���������!", vbInformation, gstrSysName
        Me.txtסԺ��.SetFocus
        Exit Sub
    End If
    
    strInput = gstrδ�����Ŀ��ѯ & "|" & GetIdentify(mlng����ID, mlng��ҳID) & "|" & TabList.Tab + 1 & "|" & Format(zlDatabase.Currentdate(), "yyyyMMdd")
    If HandleBusiness(strInput, OutputData) Then
        '���β������ݿ�
        lngCount = Val(OutputData(1))
        LNGMAX = IIf(lngCount > 60, 60, lngCount)
        For lngMin = 1 To LNGMAX
            gstrSQL = "zlYB_�����Ŀ��_UPDATE(" & TabList.Tab + 1 & "," & mlng����ID & "," & mlng��ҳID & ",'" & OutputData(lngMin + 1) & "',0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����������Ŀ��")
        Next
        
        Call ShowData
        If lngCount > 60 Then
            MsgBox "���ι���" & lngCount & "����ϸ��Ҫ���,���ڶ���ӿ�һ����෵��60����ϸ,�������굱ǰ���ݺ����»�ȡ���µĴ������ϸ����������", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitBill(Bill(���շ���Ŀ))
    Call InitBill(Bill(ѪҺ�׵���))
    
    cmdδ�����Ŀ��ѯ.Visible = (mintInsure = TYPE_������)
    If mintInsure = TYPE_���������� Then Me.Caption = "������Ŀ����(�����û�ע�⣬����������ϸ�ϴ��������޸�����״̬)"
    Call gclsInsure.InitInsure(gcnOracle, mintInsure)
End Sub

Private Sub BillMarker(ByVal intState As Integer)
    Dim strState As String
    Dim objTarget As BillEdit
    Dim lngRow As Long, lngRows As Long
    
    strState = IIf(intState = 1, "��", IIf(intState = 2, "��", ""))
    Set objTarget = Bill(TabList.Tab)
    lngRows = objTarget.Rows - 1
    
    For lngRow = 1 To lngRows
        objTarget.TextMatrix(lngRow, Col_��˱�־) = strState
    Next
End Sub

Private Sub InitBill(ByVal objTarget As BillEdit)
    With objTarget
        .ClearBill
        .Rows = 2
        .Cols = Col_Count
        
        .TextMatrix(0, Col_����ID) = "ID"
        .TextMatrix(0, Col_��Ŀ��Ϣ) = "��Ŀ��Ϣ"
        .TextMatrix(0, Col_���) = "���"
        .TextMatrix(0, Col_����) = "����"
        .TextMatrix(0, Col_���) = "���"
        .TextMatrix(0, Col_������ˮ��) = "������ˮ��"
        .TextMatrix(0, Col_��˱�־) = "��˱�־"
        .ColWidth(Col_����ID) = 0
        .ColWidth(Col_��Ŀ��Ϣ) = 2200
        .ColWidth(Col_���) = 1500
        .ColWidth(Col_����) = 1200
        .ColWidth(Col_���) = 1200
        .ColWidth(Col_������ˮ��) = 1300
        .ColWidth(Col_��˱�־) = 800
        .ColData(Col_����ID) = 5
        .ColData(Col_��Ŀ��Ϣ) = 5
        .ColData(Col_���) = 5
        .ColData(Col_����) = 5
        .ColData(Col_���) = 5
        .ColData(Col_������ˮ��) = 5
        .ColData(Col_��˱�־) = 0
        
        .PrimaryCol = Col_����ID
        .LocateCol = Col_��˱�־
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub ShowData()
    Dim objTarget As BillEdit
    Dim rsTemp As New ADODB.Recordset
    
    Set objTarget = Bill(TabList.Tab)
    Call InitBill(objTarget)
    
    '��ȡ������Ŀ
    If mintInsure = TYPE_������ Then
        gstrSQL = " Select A.ID,'['||C.����||']'||C.���� AS ��Ŀ��Ϣ,C.���,A.����*A.���� AS ����,A.ʵ�ս�� AS ���,B.������ˮ��,B.��˱�־" & _
                  " From סԺ���ü�¼ A,�����Ŀ�� B,�շ�ϸĿ C" & _
                  " Where Substr(A.ժҪ||'|',1,Instr(A.ժҪ||'|','|',1,1)-1)=B.������ˮ�� And A.����ID=B.����ID And A.��ҳID=B.��ҳID " & _
                  " And A.�շ�ϸĿID=C.ID And Nvl(A.ʵ�ս��,0)<>0" & _
                  " And B.����ID=" & mlng����ID & " And B.��ҳID=" & mlng��ҳID & " And B.����=" & TabList.Tab + 1 & _
                  IIf(chk��ʾ������Ŀ.Value = 1, "", " And Nvl(B.��˱�־,0)=0") & _
                  " Order by B.������ˮ��"
        Call OpenRecordset(rsTemp, "��ȡ������Ŀ")
    Else
        gstrSQL = " Select 0 AS ID,'['||C.����||']'||C.���� AS ��Ŀ��Ϣ,C.���,A.����,A.���,A.������ˮ��,B.��˱�־" & _
                  " From �м��_������ϸ A,�����Ŀ�� B,ZLHIS.�շ�ϸĿ C" & _
                  " Where A.������ˮ��=B.������ˮ�� And A.��Ŀ����=C.����" & _
                  " And B.����ID=" & mlng����ID & " And B.��ҳID=" & mlng��ҳID & " And B.����=" & TabList.Tab + 1 & _
                  IIf(chk��ʾ������Ŀ.Value = 1, "", " And Nvl(B.��˱�־,0)=0") & _
                  " Order by B.������ˮ��"
        Call OpenRecordset(rsTemp, "��ȡ������Ŀ", gstrSQL, gcn����������)
    End If
    
    With rsTemp
        Do While Not .EOF
            objTarget.TextMatrix(.AbsolutePosition, Col_����ID) = !ID
            objTarget.TextMatrix(.AbsolutePosition, Col_��Ŀ��Ϣ) = !��Ŀ��Ϣ
            objTarget.TextMatrix(.AbsolutePosition, Col_���) = Nvl(!���)
            objTarget.TextMatrix(.AbsolutePosition, Col_����) = !����
            objTarget.TextMatrix(.AbsolutePosition, Col_���) = !���
            objTarget.TextMatrix(.AbsolutePosition, Col_������ˮ��) = !������ˮ��
            objTarget.TextMatrix(.AbsolutePosition, Col_��˱�־) = IIf(!��˱�־ = 1, "��", IIf(!��˱�־ = 2, "��", ""))
            
            .MoveNext
            objTarget.Rows = objTarget.Rows + 1
        Loop
    End With
End Sub

Private Sub TabList_Click(PreviousTab As Integer)
    Call ShowData
End Sub

Private Sub txtסԺ��_KeyDown(KeyCode As Integer, Shift As Integer)
    '��ȡ�ò��˵Ļ�����Ϣ����ҽ������ֱ���˳�
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    '��ȡ���˻�����Ϣ
    gstrSQL = " Select A.����ID,A.סԺ���� AS ��ҳID,A.����,A.סԺ��,B.ҽ���� " & _
              " From ������Ϣ A,�����ʻ� B" & _
              " Where A.����ID=B.����ID And B.����=" & mintInsure & " And A.סԺ��='" & txtסԺ��.Text & "'"
    Call OpenRecordset(rsTemp, "��ȡ���˻�����Ϣ")
    If rsTemp.RecordCount = 0 Then
        txt��������.Text = ""
        txtҽ����.Text = ""
        MsgBox "�ò��˲�����ҽ�����ˣ�����¼���סԺ�Ų����ڣ�", vbInformation, gstrSysName
        txtסԺ��.SetFocus
        Exit Sub
    End If
    
    '��ʾ���˵Ļ�����Ϣ
    Me.txt��������.Text = rsTemp!����
    Me.txtҽ����.Text = rsTemp!ҽ����
    mlng����ID = rsTemp!����ID
    mlng��ҳID = rsTemp!��ҳID
    
    Call ShowData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
