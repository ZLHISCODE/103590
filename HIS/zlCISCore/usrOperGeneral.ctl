VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.UserControl usrOperGeneral 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   8040
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   2
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   660
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   10
      Left            =   915
      MaxLength       =   20
      TabIndex        =   7
      Top             =   990
      Width           =   2145
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   15
      TabIndex        =   10
      Top             =   1755
      Width           =   690
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   3120
      Width           =   675
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2370
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   915
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1320
      Width           =   1875
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      ItemData        =   "usrOperGeneral.ctx":0000
      Left            =   915
      List            =   "usrOperGeneral.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2715
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   3
      Left            =   945
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   1
      Left            =   945
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2055
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   2
      Left            =   945
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   4
      Left            =   945
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3450
      Width           =   1845
   End
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   1380
      Index           =   0
      Left            =   3960
      TabIndex        =   31
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
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
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   0
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   915
      TabIndex        =   3
      Top             =   330
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   2
      Left            =   915
      TabIndex        =   13
      Top             =   1680
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   3
      Left            =   915
      TabIndex        =   16
      Top             =   2025
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   4
      Left            =   915
      TabIndex        =   24
      Top             =   3060
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   5
      Left            =   915
      TabIndex        =   27
      Top             =   3405
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   1380
      Index           =   1
      Left            =   3960
      TabIndex        =   33
      Top             =   1365
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
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
   Begin ZL9BillEdit.BillEdit bill2 
      Height          =   1380
      Index           =   0
      Left            =   3960
      TabIndex        =   35
      Top             =   2730
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
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
   Begin ZL9BillEdit.BillEdit bill2 
      Height          =   1800
      Index           =   1
      Left            =   3960
      TabIndex        =   37
      Top             =   4095
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   3175
      Appearance      =   0
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
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   2145
      Index           =   2
      Left            =   900
      TabIndex        =   29
      Top             =   3750
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3784
      Appearance      =   0
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ģ"
      Height          =   180
      Index           =   17
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� ��"
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   5
      Left            =   690
      TabIndex        =   2
      Top             =   375
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   10
      Left            =   3180
      TabIndex        =   30
      Top             =   30
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   9
      Left            =   3180
      TabIndex        =   32
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   6
      Left            =   675
      TabIndex        =   14
      Top             =   2070
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   7
      Left            =   690
      TabIndex        =   11
      Top             =   1755
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʽ"
      Height          =   180
      Index           =   8
      Left            =   105
      TabIndex        =   17
      Top             =   2415
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   19
      Top             =   2775
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   12
      Left            =   675
      TabIndex        =   25
      Top             =   3465
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Index           =   11
      Left            =   675
      TabIndex        =   22
      Top             =   3120
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Һ����                      ML"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1380
      Width           =   2880
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ա"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   3750
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Index           =   2
      Left            =   3180
      TabIndex        =   36
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ���"
      Height          =   180
      Index           =   3
      Left            =   3180
      TabIndex        =   34
      Top             =   2730
      Width           =   720
   End
End
Attribute VB_Name = "usrOperGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private i As Long
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private mlng����id As Long                      '��紫��
Private mlngҽ��id As Long                      '��紫��
Private mstr�Ա� As String

Private mlng������¼id As Long

Private mlng������id As Long                    '�ݴ����
Private mlngOrderID As Long                     '�ݴ����

Private Const STR_COMPART = "|';"
Private Const LAWLChar = "';`|,"""

Private mblnMode As Boolean 'Ϊ���Ǳ�ʾ���û����еı༭����ʱ�Ÿ�ֵ

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String

Private mblnLoaded As Boolean

'-------------------------------------------------------------------------------------------------------------------
'��������������
Public Property Get DispMode() As Boolean
    '�Ƿ�Ϊ��ʾģʽ
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    ShowOperGeneral mlng����id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        
        dtp(0).Enabled = False
        dtp(1).Enabled = False
        dtp(2).Enabled = False
        dtp(3).Enabled = False
        dtp(4).Enabled = False
        dtp(5).Enabled = False
        
        cbo(0).Locked = True
        cbo(1).Locked = True
        cbo(2).Locked = True
        
        txt(0).Locked = True
        txt(10).Locked = True
        
        bill(0).Active = False
        bill(1).Active = False
        bill(2).Active = False
        
        bill2(0).Active = False
        bill2(1).Active = False
        
        chk(0).Enabled = False
        chk(1).Enabled = False
        
    End If
    
End Property

Public Property Get ID���˲���() As Long
    '���ز��˲���ID
    
    ID���˲��� = mlng����id
End Property

Public Property Let ID���˲���(ByVal New_ID���˲��� As Long)
    '���ò��˲���ID,�����ò����ǲ��Ǵ���
    
    mlng����id = New_ID���˲���
    ShowOperGeneral mlng����id, Not mDispMode
    
End Property

Public Property Let Getҽ��id(ByVal New_ҽ��ID As Long)
        
    mlngҽ��id = New_ҽ��ID
        
End Property

Public Property Get Getҽ��id() As Long
        
    Getҽ��id = mlngҽ��id
        
End Property

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '���ô��������������
    '���lngErrNum=-1 ��ʾ �ؼ��Լ�����Ĵ���
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
    '�������һ�εĴ����
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
    '�������һ�δ��������ַ���
    ReturnErrDescription = mReturnErrDescription
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function CheckStrValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByRef strError As String) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
        
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        strError = "���������ݺ��зǷ��ַ���"
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            strError = "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��"
            Exit Function
        End If
    End If
    
    CheckStrValid = True
End Function

Private Function PopSelect(ByVal objBill As BillEdit, Optional ByVal str�Ա� As String = "0") As Boolean
    '----------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim sglX As Single
    Dim sglY As Single
    Dim strNote As String
    Dim strLvw As String
           
    On Error GoTo errHand
    
    CalcPosition sglX, sglY, objBill
    
    Select Case objBill.Col
    Case 1
        gstrSql = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "���� " & _
                "FROM ������Ϸ��� " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "B.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.���� " & _
                "FROM �������Ŀ¼ A,����������� B " & _
                "WHERE A.ID=B.���ID "
                    
        strNote = "��ѡ��һ�����������Ŀ"
        strLvw = "����,1200,0,1;����,2400,0,2"
    Case 0
        gstrSql = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "NULL AS ����," & _
                        "����," & _
                        "NULL AS ���� " & _
                "FROM ����������� " & _
                "WHERE ���='D' " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.���� " & _
                "FROM ��������Ŀ¼ A " & _
                "WHERE ���='D' " & _
                    "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (" & str�Ա� & ") "
                    
        strNote = "��ѡ��һ������������Ŀ"
        strLvw = "����,1200,0,1;����,2400,0,2;����,810,0,0"
    End Select
    
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ"
    
    If rs.BOF Then Exit Function
    
    If frmSelectTree.ShowSelect(Screen, _
                                rs, _
                                sglX, sglY, 5400, 2400, _
                                objBill.MsfObj.CellHeight, _
                                "�����ʾ_2", _
                                strLvw, _
                                strNote) Then
        
        PopSelect = True
        
        objBill.Text = zlCommFun.Nvl(rs("����").Value)
        
        Select Case objBill.Col
        Case 0
            objBill.TextMatrix(objBill.Row, 0) = objBill.Text
            objBill.TextMatrix(objBill.Row, 4) = zlCommFun.Nvl(rs("ID").Value)
        Case 1
            objBill.TextMatrix(objBill.Row, 1) = objBill.Text
            objBill.TextMatrix(objBill.Row, 3) = zlCommFun.Nvl(rs("ID").Value)
        End Select
        objBill.TextMatrix(objBill.Row, 2) = zlCommFun.Nvl(rs("����").Value)
        
        objBill.RowData(objBill.Row) = "1"
        
        '������Ӧ�ļ������Ŀ¼�򼲲�����
        MatchDiagnoses Val(objBill.TextMatrix(objBill.Row, 4)), Val(objBill.TextMatrix(objBill.Row, 3)), objBill
                
    End If
   
    Exit Function
   
errHand:
   If ErrCenter = 1 Then Resume
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As BillEdit)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.MsfObj.hwnd, objPoint)
    
    x = objPoint.x * 15 + objBill.MsfObj.CellLeft - 45
    y = objPoint.y * 15 + objBill.MsfObj.CellTop + objBill.MsfObj.CellHeight - 30
End Sub

Private Sub MatchDiagnoses(ByVal lngCodeKey As Long, ByVal lngListKey As Long, objMsf As BillEdit, Optional ByVal strCaption As String)
    Dim rs As New ADODB.Recordset
    '----------------------------------------------------------------------
    '1.֪���������룬���Ӧ�ļ������Ŀ¼
    '2.֪���������Ŀ¼�����Ӧ�ļ�������
    '----------------------------------------------------------------------
    gstrSql = "SELECT A.����ID,A.���ID,B.���� AS ��������,C.���� AS ������� " & _
                "FROM ������϶��� A,��������Ŀ¼ B,�������Ŀ¼ C " & _
                "WHERE A.����ID=B.ID AND A.���ID=C.ID AND (A.����ID=" & lngCodeKey & " OR A.���ID=" & lngListKey & ")"
                
    zlDatabase.OpenRecordset rs, gstrSql, strCaption
    If rs.BOF = False Then
        If rs.RecordCount > 0 Then
            objMsf.TextMatrix(objMsf.Row, 0) = zlCommFun.Nvl(rs("�������"))
            objMsf.TextMatrix(objMsf.Row, 1) = zlCommFun.Nvl(rs("��������"))
            objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("���ID"))
            objMsf.TextMatrix(objMsf.Row, 3) = zlCommFun.Nvl(rs("����ID"))
        End If
    End If
End Sub

Private Sub ReDimArray(ByRef LngCount As Long, ByRef strArray() As String)
    
    '���ܣ����¶�������
    LngCount = LngCount + 1
    ReDim Preserve strArray(1 To LngCount)
        
End Sub

Private Function ShowDownList2(ByVal frmMain As Object, _
                            ByVal bytMode As Byte, _
                            objMsf As BillEdit, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            Optional ByVal blnWhere As Boolean = False) As Boolean
    '----------------------------------------------------------------------
    '����:��ʾ������ʾ�Ի���
    '����:
    '����:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim strInput As String
    Dim sglWidth As Single
    Dim sglHeight As Single
    Dim strPath As String
        
    On Error GoTo errHand
    
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "�������зǷ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                
    Select Case bytMode
    Case 1
        gstrSql = "SELECT   ����," & _
                           "����," & _
                           "����," & _
                           "����," & _
                           "ID " & _
                    "FROM ��������Ŀ¼ " & _
                    "WHERE ���='D' " & _
                        "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (" & mstr�Ա� & ") " & _
                        "AND (���� LIKE " & strInput & " OR ���� LIKE " & strInput & " OR ���� LIKE " & strInput & ")"
        sglWidth = 5100
        sglHeight = 2400
        strLvw = "����,1200,0,0;����,2400,0,0;����,900,0,0;����,900,0,0"
        strPath = "�������_����"
    Case 2
        gstrSql = "SELECT A.����," & _
                           "A.����," & _
                           "A.ID " & _
                    "FROM �������Ŀ¼ A " & _
                    "Where A.��� = 1 " & _
                          "AND (���� LIKE " & strInput & " OR ���� LIKE " & strInput & " " & _
                          "OR A.id IN (SELECT B.���id " & _
                                        "FROM ������ϱ��� B " & _
                                        "WHERE 1=1 " & _
                                            "AND (���� LIKE " & strInput & " OR ���� LIKE " & strInput & ")))"
        sglWidth = 4500
        sglHeight = 2400
        strLvw = "����,1200,0,0;����,3000,0,0"
        strPath = "�������_Ŀ¼"
    End Select
            
    Call zlDatabase.OpenRecordset(rs, gstrSql, "������Ҫ")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownList2 = True
        GoTo FillPoint
        Exit Function
    End If
        
    If frmSelectList.ShowSelect(Screen, rs, strLvw, x, y, sglWidth, sglHeight, "������Ҫ\" & strPath, "�������ѡ��һ����Ŀ") Then
        ShowDownList2 = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:

    objMsf.Text = zlCommFun.Nvl(rs("����"))
    Select Case bytMode
    Case 1
        objMsf.TextMatrix(objMsf.Row, 0) = objMsf.Text
        objMsf.TextMatrix(objMsf.Row, 4) = zlCommFun.Nvl(rs("ID"))
    Case 2
        objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
        objMsf.TextMatrix(objMsf.Row, 3) = zlCommFun.Nvl(rs("ID"))
    End Select
    objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("����"))
    objMsf.RowData(objMsf.Row) = "1"
    
    '������Ӧ�ļ������Ŀ¼�򼲲�����
    Call MatchDiagnoses(Val(objMsf.TextMatrix(objMsf.Row, 4)), Val(objMsf.TextMatrix(objMsf.Row, 3)), objMsf, "������Ҫ��")
            
    Exit Function
errHand:
    objMsf.Text = ""
End Function

Private Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'����: װ��������ָ�������������������е���������
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Private Function PopOperateSelect(ByVal objBill As BillEdit, ByVal bytMode As Byte) As Boolean
    '----------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim sglX As Single
    Dim sglY As Single
    Dim strNote As String
    Dim strLvw As String
    Dim str�Ա� As String
           
    On Error GoTo errHand
    
    
    '��ѯ���Ա�
    str�Ա� = "0,1,2"
    If mlng����id > 0 Then
        gstrSql = "SELECT B.�Ա� FROM ���˲�����¼ A,������Ϣ B WHERE A.����ID=B.����ID  and A.ID=" & mlng����id
    Else
        gstrSql = "SELECT B.�Ա� FROM ����ҽ����¼ A,������Ϣ B WHERE A.����ID=B.����ID  and A.ID=" & mlngҽ��id
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫר��ֽ"
    If rs.BOF = False Then
        If zlCommFun.Nvl(rs("�Ա�").Value, "") Like "*��*" Then str�Ա� = "0,1"
        If zlCommFun.Nvl(rs("�Ա�").Value, "") Like "*Ů*" Then str�Ա� = "0,2"
    End If
    
    CalcPosition sglX, sglY, objBill
    
    Select Case bytMode
    Case 1          '������Ŀ
        gstrSql = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "����," & _
                        "����," & _
                        "NULL AS ��λ " & _
                "FROM ���Ʒ���Ŀ¼ " & _
                "WHERE ����=5 " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.���㵥λ AS ��λ " & _
                "FROM ������ĿĿ¼ A "
        gstrSql = gstrSql & _
                "WHERE (����ʱ�� = TO_DATE('30000101', 'YYYYMMDD') OR ����ʱ�� IS NULL) " & _
                    "AND ������� IN (2, 3) " & _
                    "AND ��� = 'F' " & _
                    "AND NVL(�����Ա�,0) IN (" & str�Ա� & ")"
                    
        strNote = "��ѡ��һ������������Ŀ"
        strLvw = "����,1200,0,1;����,2400,0,2;��λ,900,0,0"
    Case 2
        gstrSql = "SELECT ID," & _
                        "�ϼ�ID," & _
                        "0 AS ĩ��," & _
                        "NULL AS ����," & _
                        "����," & _
                        "NULL AS ���� " & _
                "FROM ����������� " & _
                "WHERE ���='D' " & _
                "START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.����id AS �ϼ�ID, " & _
                        "1 AS ĩ��, " & _
                        "A.����, " & _
                        "A.����, " & _
                        "A.���� " & _
                "FROM ��������Ŀ¼ A " & _
                "WHERE ���='S' " & _
                    "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (" & str�Ա� & ") "
                    
        strNote = "��ѡ��һ������������Ŀ"
        strLvw = "����,1200,0,1;����,2400,0,2;����,810,0,0"
    End Select
    
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ"
    
    If rs.BOF Then Exit Function
    
    If frmSelectTree.ShowSelect(Screen, _
                                rs, _
                                sglX, _
                                sglY, _
                                9000, _
                                3000, _
                                objBill.MsfObj.CellHeight, _
                                "������ʾ_2", _
                                strLvw, _
                                strNote) Then
        
        If CheckHave(objBill, objBill.Row, zlCommFun.Nvl(rs("ID").Value)) Then
            MsgBox "���б����Ѿ����ڴ���Ŀ[" & zlCommFun.Nvl(rs("����").Value) & "]��", vbInformation, gstrSysName
            Exit Function
        End If
        
        PopOperateSelect = True
        
        objBill.Text = zlCommFun.Nvl(rs("����").Value)
        objBill.TextMatrix(objBill.Row, 1) = objBill.Text
        objBill.RowData(objBill.Row) = zlCommFun.Nvl(rs("ID").Value)

    End If
   
    Exit Function
   
errHand:
   If ErrCenter = 1 Then Resume
End Function

Private Function ShowDownListOperate(ByVal frmMain As Object, ByVal bytMode As Byte, objMsf As BillEdit, Optional ByVal blnWhere As Boolean = False, Optional ByVal str�Ա� As String = "0") As Boolean
    '----------------------------------------------------------------------
    '����:��ʾ������ʾ�Ի���
    '����:
    '����:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim strInput As String
    Dim sglWidth As Single
    Dim sglHeight As Single
    Dim strPath As String
    Dim sglY As Single
    Dim sglX As Single
        
    On Error GoTo errHand
    
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "�������зǷ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                
    Select Case bytMode
    Case 2
        gstrSql = "SELECT   ����," & _
                           "����," & _
                           "����," & _
                           "����," & _
                           "ID " & _
                    "FROM ��������Ŀ¼ " & _
                    "WHERE ���='S' " & _
                        "AND DECODE(�Ա�����,'��',1,'Ů',2,0) IN (" & str�Ա� & ") " & _
                        "AND (���� LIKE " & strInput & " OR ���� LIKE " & strInput & " OR ���� LIKE " & strInput & ")"
        sglWidth = 5100
        sglHeight = 2400
        strLvw = "����,1200,0,0;����,2400,0,0;����,900,0,0;����,900,0,0"
        strPath = "����_����"
    Case 1
        gstrSql = "SELECT A.����," & _
                           "A.����," & _
                           "A.ID " & _
                    "FROM ������ĿĿ¼ A " & _
                    "Where A.��� = 'F' " & _
                          "AND A.������� IN (2, 3) " & _
                          "AND NVL(�����Ա�,0) IN (" & str�Ա� & ") " & _
                          "AND (���� LIKE " & strInput & " OR ���� LIKE " & strInput & " " & _
                          "OR A.id IN (SELECT B.������ĿID " & _
                                        "FROM ������Ŀ���� B " & _
                                        "WHERE (���� LIKE " & strInput & " OR ���� LIKE " & strInput & ")))"
        sglWidth = 9000
        sglHeight = 3000
        strLvw = "����,1200,0,0;����,3000,0,0"
        strPath = "����_����"
    End Select
            
    Call zlDatabase.OpenRecordset(rs, gstrSql, "������Ҫ")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownListOperate = True
        GoTo FillPoint
        Exit Function
    End If
    
    CalcPosition sglX, sglY, objMsf
    
    If frmSelectList.ShowSelect(Screen, rs, strLvw, sglX, sglY, sglWidth, sglHeight, "������Ҫ\" & strPath, "�������ѡ��һ����Ŀ") Then
        ShowDownListOperate = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:
    If CheckHave(objMsf, objMsf.Row, zlCommFun.Nvl(rs("ID").Value)) Then
        MsgBox "���б����Ѿ����ڴ���Ŀ[" & zlCommFun.Nvl(rs("����").Value) & "]��", vbInformation, gstrSysName
        Exit Function
    End If
    
    objMsf.Text = zlCommFun.Nvl(rs("����"))
    objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
    objMsf.RowData(objMsf.Row) = zlCommFun.Nvl(rs("ID"))
    
    Exit Function
errHand:
    objMsf.Text = ""
End Function

Public Function ShowDownListPerson(ByVal frmMain As Object, _
                                objMsf As BillEdit, _
                                ByVal lngDeptKey As Long, _
                                ByVal x As Single, _
                                ByVal y As Single, _
                                Optional ByVal blnWhere As Boolean = False, _
                                Optional ByVal blnFlag As Boolean = False, _
                                Optional ByVal lngKey As Long = 0) As Boolean
    '----------------------------------------------------------------------
    '����:��ʾ������ʾ�Ի���
    '����:
    '����:
    '----------------------------------------------------------------------
    Dim strInput As String
    Dim strSelected As String
    Dim lngLoop As Long
    Dim strClass As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "�������зǷ��ַ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                        
    strSelected = "0"
    For lngLoop = 1 To objMsf.Rows - 1
        If lngLoop <> objMsf.Row Then
            strSelected = strSelected & "," & objMsf.RowData(lngLoop)
        End If
    Next
    
    strClass = "ҽ��"
    If InStr("ϴ�ֻ�ʿ;Ѳ�ػ�ʿ", Trim(objMsf.TextMatrix(objMsf.Row, 0))) > 0 Then strClass = "��ʿ"
        
    gstrSql = "SELECT   A.���," & _
                       "A.����," & _
                       "A.����," & _
                       "C.���� AS ����," & _
                       "DECODE(C.ID," & lngDeptKey & ",1,2) AS ���," & _
                       "A.ID " & _
                "FROM ��Ա�� A,��Ա����˵�� B,���ű� C,������Ա D " & _
                "WHERE A.ID=B.��Աid AND C.ID=D.����id AND D.��Աid=A.ID AND D.ȱʡ=1 " & _
                    "AND A.ID NOT IN (" & strSelected & ") " & _
                    "AND B.��Ա����='" & strClass & "' " & _
                    "AND (A.��� LIKE " & strInput & " OR A.���� LIKE " & strInput & " OR A.���� LIKE " & strInput & ") " & _
                "ORDER BY ���"
    
    Call zlDatabase.OpenRecordset(rs, gstrSql, "������Ҫ")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownListPerson = True
        GoTo FillPoint
    End If
    
    If frmSelectList.ShowSelect(Screen, rs, "���,1200,0,0;����,2400,0,0;����,900,0,0;����,900,0,0", x, y, 8100, 3000, "������Ҫ" & "����_��Ա", "����±���ѡ��һ����Ա") Then
        ShowDownListPerson = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:
    objMsf.Text = zlCommFun.Nvl(rs("����"))
    objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
    objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("���"))
    objMsf.RowData(objMsf.Row) = zlCommFun.Nvl(rs("ID"))
    
    Exit Function
errHand:
    objMsf.Text = ""
End Function
'------------------------------------------------------------------------------------------------------------

Private Sub ShowOperGeneral(lng����ID As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------
    '���ܣ��ⲿ������ʾ������Ҫ�Ĺ���
    '------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    
    mlng����id = lng����ID
    mDispMode = Not blnEditMode
    
    mstr�Ա� = "0,1,2"
    
    '���߼�Ӧ�ȳ�ʼ�ؼ�
    InitData
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub

    '����Ƿݲ����ǲ��Ǵ���
    strSQL = _
        "SELECT a.ID" & vbCrLf & _
        "  FROM ���˲������� A" & vbCrLf & _
        " WHERE a.Ԫ������ = 4 and " & vbCrLf & _
        "      a.Ԫ�ر��� IN" & vbCrLf & _
        "      (SELECT ����" & vbCrLf & _
        "         FROM ����Ԫ��Ŀ¼" & vbCrLf & _
        "        WHERE ���� = 4 AND ���� = '������Ҫ��¼��')" & vbCrLf & _
        " AND A.id=" & mlng����id
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������Ҫ��¼��")
    
    If rsTmp.RecordCount = 0 And mlngҽ��id = 0 Then
        SetErr -1, "�ò����������޵���������Ҫ��¼����"
'        Exit Sub
    End If
    
    Call ReadData
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LocalCheck�Ƿ�Ƿ�(txt As Control, ByVal strLawlChar As String) As Boolean
'����:����ǲ��ǰ���strLawlChar����ַ���,����оͷ���Ϊ�����ͷ��ط�
On Error GoTo ErrHandle
    Dim strSour As String
    
    If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
        If TypeOf txt Is ComboBox Then
            If txt.Style <> 0 Then
                '����ComboBoxΪѡ��������ֻ����������
                LocalCheck�Ƿ�Ƿ� = True
                Exit Function
            End If
        End If
        strSour = txt.Text
        If Len(strSour) > 0 Then
            For i = 1 To Len(strLawlChar)
                If InStr(strSour, Mid(strLawlChar, i, 1)) > 0 Then
                    txt.SelStart = InStr(strSour, Mid(strLawlChar, i, 1))
                    txt.SelLength = 1
                    MsgBox "�ı�������зǷ��ַ���", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                    Exit Function
                End If
            Next
            If VarType(txt.Tag) = vbLong Or VarType(txt.Tag) = vbInteger Then
                If zlCommFun.ActualLen(strSour) > txt.Tag And txt.Tag > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            ElseIf VarType(txt.Tag) = vbString And IsNumeric(txt.Tag) Then
                If zlCommFun.ActualLen(strSour) > CLng(txt.Tag) And CLng(txt.Tag) > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetgcnOracle()
    '-------------------------------------------------------------------------------------------------
    '�ӿ�
    '-------------------------------------------------------------------------------------------------
    
    Call InitCommon(gcnOracle)
End Sub

Private Sub InitData()
    '��ʼ������
    
    Dim strTmp As String
    
    On Error GoTo ErrHandle
        
    If Not gcnOracle Is Nothing Then
        If Not gcnOracle.State <> adStateOpen Then
            If Ambient.UserMode = True Then
                strSQL = "select * FROM ������ĿĿ¼ where ��� ='G'"
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������Ҫ��¼��")
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    For i = 0 To rsTmp.RecordCount - 1
                        cbo(0).AddItem rsTmp("����") & "-" & rsTmp("����") & Space(200) & zlCommFun.Nvl(rsTmp("��������"))
                        rsTmp.MoveNext
                    Next
                    cbo(0).ListIndex = 0
                End If
            End If
        End If
    
    End If
    With cbo(1)
        .Clear
        .AddItem "1-��"
        .AddItem "2-��"
        .AddItem "3-��"
        .AddItem "4-Σ(��)"
        .ListIndex = 0
    End With
    
    strTmp = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59"
    
    dtp(0).Value = strTmp
    dtp(1).Value = strTmp
    dtp(2).Value = strTmp
    dtp(3).Value = strTmp
    dtp(4).Value = strTmp
    dtp(5).Value = strTmp
    
    mblnLoaded = True
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------
    '���ܣ��������ݿ��������
    '------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
        
    On Error GoTo ErrHandle
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Function
    
    gstrSql = "SELECT ����,0 FROM ����������ģ"
    Call zlDatabase.OpenRecordset(rs, gstrSql, "������Ҫ")
    If rs.RecordCount = 0 Then
        MsgBox "ϵͳ���ݲ�������û������������ģ�����ֵ䣡", vbInformation, gstrSysName
        Exit Function
    End If
    If rs.BOF = False Then Call AddComboData(cbo(2), rs)
    
    
    mlng������¼id = 0
        
    If mlngҽ��id = 0 Then
        strTmp = "SELECT A.ҽ��id FROM ���˲�����¼ A,���˲������� B WHERE B.������¼id=A.ID AND B.ID=" & mlng����id
        zlDatabase.OpenRecordset rs, strTmp, "������Ҫ��"
        If rs.BOF = False Then mlngҽ��id = zlCommFun.Nvl(rs("ҽ��id").Value, 0)
    End If
        
    If mlngҽ��id > 0 Then
        '�����Ǵ���������ϵͳ����
        
        'ͨ��ҽ��id���Ҳ�������id
        strTmp = "SELECT ID FROM ����������¼ WHERE ҽ��id=" & mlngҽ��id
        zlDatabase.OpenRecordset rs, strTmp, "������Ҫ��"
        If rs.BOF = False Then mlng������¼id = zlCommFun.Nvl(rs("ID").Value, 0)
                    
    Else
        '�����ǴӲ�������վ���ã���Ҫȷ�������������޸�
        
        strTmp = "SELECT ID FROM ����������¼ WHERE ����id=" & mlng����id
        zlDatabase.OpenRecordset rs, strTmp, "������Ҫ��"
        If rs.BOF = False Then mlng������¼id = zlCommFun.Nvl(rs("ID").Value, 0)
        
    End If
    
    strTmp = "SELECT ҽ��id,������,������id FROM ����������¼ WHERE ID=" & mlng������¼id
    zlDatabase.OpenRecordset rs, strTmp, "������Ҫ��"
    If rs.BOF = False Then
        mlng������id = zlCommFun.Nvl(rs("������id").Value, 0)
        mlngOrderID = zlCommFun.Nvl(rs("ҽ��id").Value, 0)
    End If
    
    '1.��ȡ������������
    gstrSql = "SELECT A.* FROM ����������¼ A WHERE A.ID=" & mlng������¼id
    
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        dtp(0).Value = Format(zlCommFun.Nvl(rs("������ʼʱ��")), "YYYY-MM-DD HH:MM")
        dtp(1).Value = Format(zlCommFun.Nvl(rs("��������ʱ��")), "YYYY-MM-DD HH:MM")
        If IsNull(rs("����ʼʱ��")) = False Then
            chk(0).Value = 1
            dtp(2).Value = Format(zlCommFun.Nvl(rs("����ʼʱ��")), "YYYY-MM-DD HH:MM")
            dtp(3).Value = Format(zlCommFun.Nvl(rs("�������ʱ��")), "YYYY-MM-DD HH:MM")
        Else
            dtp(2).Value = Format(zlCommFun.Nvl(rs("������ʼʱ��")), "YYYY-MM-DD HH:MM")
            dtp(3).Value = Format(zlCommFun.Nvl(rs("��������ʱ��")), "YYYY-MM-DD HH:MM")
        End If
        If IsNull(rs("������ʼʱ��")) = False Then
            chk(1).Value = 1
            dtp(4).Value = Format(zlCommFun.Nvl(rs("������ʼʱ��")), "YYYY-MM-DD HH:MM")
            dtp(5).Value = Format(zlCommFun.Nvl(rs("��������ʱ��")), "YYYY-MM-DD HH:MM")
        Else
            dtp(4).Value = Format(zlCommFun.Nvl(rs("������ʼʱ��")), "YYYY-MM-DD HH:MM")
            dtp(5).Value = Format(zlCommFun.Nvl(rs("��������ʱ��")), "YYYY-MM-DD HH:MM")
        End If
                        
                        
        zlControl.CboLocate cbo(0), zlCommFun.Nvl(rs("����ʽ"))
        zlControl.CboLocate cbo(1), zlCommFun.Nvl(rs("��������"))
        zlControl.CboLocate cbo(2), zlCommFun.Nvl(rs("������ģ"))
        txt(10).Text = zlCommFun.Nvl(rs("������"))
        txt(0).Text = zlCommFun.Nvl(rs("��Һ����"))
    End If
    
    bill(0).Rows = 2
    bill(1).Rows = 2
    bill(2).Rows = 2
    bill2(0).Rows = 2
    bill2(1).Rows = 2
    ClearSpecRowCol bill(0), 1, Array()
    ClearSpecRowCol bill(1), 1, Array()
    ClearSpecRowCol bill(2), 1, Array()
    ClearSpecRowCol bill2(0), 1, Array()
    ClearSpecRowCol bill2(1), 1, Array()
    
    '2.��ȡ����������¼
    gstrSql = "SELECT DECODE(A.������ĿID,null,'2-����','1-����') AS ������Դ," & _
                    "A.��������," & _
                    "A.ȱʡ," & _
                    "DECODE(A.������Ŀid,NULL,A.��������ID,A.������Ŀid) AS ID " & _
                "FROM ����������� A,����������¼ B " & _
                "WHERE A.��¼id=B.ID " & _
                        "AND A.����=1 " & _
                        "AND B.ID=" & mlng������¼id
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(0).RowData(1) > 0 Then bill(0).Rows = bill(0).Rows + 1
            
            bill(0).RowData(bill(0).Rows - 1) = zlCommFun.Nvl(rs("ID").Value, 0)
            bill(0).TextMatrix(bill(0).Rows - 1, 0) = zlCommFun.Nvl(rs("������Դ").Value)
            bill(0).TextMatrix(bill(0).Rows - 1, 1) = zlCommFun.Nvl(rs("��������").Value)
            bill(0).TextMatrix(bill(0).Rows - 1, 2) = IIf(zlCommFun.Nvl(rs("ȱʡ").Value) = 1, "��", "")
            
            rs.MoveNext
        Loop
    End If
            
    
    '3.��ȡ����������¼
    gstrSql = "SELECT DECODE(A.������ĿID,null,'2-����','1-����') AS ������Դ," & _
                    "A.��������," & _
                    "A.ȱʡ," & _
                    "DECODE(A.������Ŀid,NULL,A.��������ID,A.������Ŀid) AS ID " & _
                "FROM ����������� A,����������¼ B " & _
                "WHERE A.��¼id=B.ID " & _
                        "AND A.����=2 " & _
                        "AND B.ID=" & mlng������¼id
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(1).RowData(1) > 0 Then bill(1).Rows = bill(1).Rows + 1
            
            bill(1).RowData(bill(1).Rows - 1) = zlCommFun.Nvl(rs("ID").Value, 0)
            bill(1).TextMatrix(bill(1).Rows - 1, 0) = zlCommFun.Nvl(rs("������Դ").Value)
            bill(1).TextMatrix(bill(1).Rows - 1, 1) = zlCommFun.Nvl(rs("��������").Value)
            bill(1).TextMatrix(bill(1).Rows - 1, 2) = IIf(zlCommFun.Nvl(rs("ȱʡ").Value) = 1, "��", "")
            
            rs.MoveNext
        Loop
    End If
    If bill(1).RowData(1) = 0 Then CopyMsfGrid bill(0), bill(1)
        
    '3.��ȡ��ǰ��ϼ�¼
    If mlngҽ��id > 0 Then
        gstrSql = "select ���ID," & _
                          "����ID," & _
                          "(select ���� FROM �������Ŀ¼ where id = ���ID) AS ��ϱ���," & _
                          "(select ���� FROM ��������Ŀ¼ where id = ����ID) AS ��������," & _
                          "������� " & _
                     "From ������ϼ�¼ " & _
                    "where ҽ��id = " & mlngҽ��id & " and ������� = 8"
    Else
        gstrSql = "select ���ID," & _
                          "����ID," & _
                          "(select ���� FROM �������Ŀ¼ where id = ���ID) AS ��ϱ���," & _
                          "(select ���� FROM ��������Ŀ¼ where id = ����ID) AS ��������," & _
                          "������� " & _
                     "From ������ϼ�¼ " & _
                    "where ����id = " & mlng����id & " and ������� = 8"
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill2(0).RowData(1) > 0 Then bill2(0).Rows = bill2(0).Rows + 1
            
            bill2(0).RowData(bill2(0).Rows - 1) = "1"
            bill2(0).TextMatrix(bill2(0).Rows - 1, 1) = zlCommFun.Nvl(rs("��ϱ���").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 0) = zlCommFun.Nvl(rs("��������").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 2) = zlCommFun.Nvl(rs("�������").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 3) = zlCommFun.Nvl(rs("���ID").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 4) = zlCommFun.Nvl(rs("����ID").Value)
            
            rs.MoveNext
        Loop
    End If
    
     '3.��ȡ������ϼ�¼
    If mlngҽ��id > 0 Then
        gstrSql = "select ���ID," & _
                          "����ID," & _
                          "(select ���� FROM �������Ŀ¼ where id = ���ID) AS ��ϱ���," & _
                          "(select ���� FROM ��������Ŀ¼ where id = ����ID) AS ��������," & _
                          "������� " & _
                     "From ������ϼ�¼ " & _
                    "where ҽ��id = " & mlngҽ��id & " and ������� = 9"
    Else
        gstrSql = "select ���ID," & _
                          "����ID," & _
                          "(select ���� FROM �������Ŀ¼ where id = ���ID) AS ��ϱ���," & _
                          "(select ���� FROM ��������Ŀ¼ where id = ����ID) AS ��������," & _
                          "������� " & _
                     "From ������ϼ�¼ " & _
                    "where ����id = " & mlng����id & " and ������� = 9"
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill2(1).RowData(1) > 0 Then bill2(1).Rows = bill2(1).Rows + 1
            
            bill2(1).RowData(bill2(1).Rows - 1) = "1"
            bill2(1).TextMatrix(bill2(1).Rows - 1, 1) = zlCommFun.Nvl(rs("��ϱ���").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 0) = zlCommFun.Nvl(rs("��������").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 2) = zlCommFun.Nvl(rs("�������").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 3) = zlCommFun.Nvl(rs("���ID").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 4) = zlCommFun.Nvl(rs("����ID").Value)
            
            rs.MoveNext
        Loop
    End If
    If bill2(1).RowData(1) = 0 Then CopyMsfGrid bill2(0), bill2(1)
    
    '3.��ȡ������Ա��¼
    gstrSql = "SELECT A.��Աid," & _
                    "DECODE(A.��λ,'����ҽ��','1-����ҽ��','����ҽ��',2,'����ҽ��',3,'ϴ�ֻ�ʿ',4,5) AS ���," & _
                    "D.���� AS ��λ," & _
                    "A.���� " & _
                "FROM ����������Ա A,������λ D " & _
                "WHERE  D.����=A.��λ " & _
                        "AND A.��¼ID=" & mlng������¼id & " " & _
                "ORDER BY ���"
                
    zlDatabase.OpenRecordset rs, gstrSql, "������Ҫ��"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(2).RowData(1) > 0 Then bill(2).Rows = bill(2).Rows + 1
            
            bill(2).RowData(bill(2).Rows - 1) = zlCommFun.Nvl(rs("��Աid").Value, 0)
            bill(2).TextMatrix(bill(2).Rows - 1, 0) = zlCommFun.Nvl(rs("��λ").Value)
            bill(2).TextMatrix(bill(2).Rows - 1, 1) = zlCommFun.Nvl(rs("����").Value)
            
            rs.MoveNext
        Loop
    End If
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
'����: ���ָ�������ָ����ָ���е�����
'����: obj=Ҫ����������ؼ�
'      intRow=Ҫ������к�
'      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Private Sub CopyMsfGrid(ByVal objFrom As Object, ByRef objTo As Object)
    Dim lngRow As Long
    Dim lngCol As Long
    
    objTo.Rows = objFrom.Rows
    objTo.Cols = objFrom.Cols
    
    For lngRow = 1 To objFrom.Rows - 1
        objTo.RowData(lngRow) = objFrom.RowData(lngRow)
        For lngCol = 0 To objFrom.Cols - 1
            objTo.TextMatrix(lngRow, lngCol) = objFrom.TextMatrix(lngRow, lngCol)
        Next
    Next
End Sub

Private Function CheckDataValid(ByRef strError As String) As Boolean
    '----------------------------------------------------------------------
    '���ܣ����������޸ĵ����ݽ��кϷ���У��
    '���أ�У��Ϸ�����True�����򷵻�False
    '----------------------------------------------------------------------
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    CheckDataValid = False
        
    strError = ""
    
    If mDispMode Then
        strError = "��ǰΪ��ʾģʽ���ܱ������ݣ�"
        SetErr -1, "��ǰΪ��ʾģʽ���ܱ������ݣ�"
        Exit Function
    End If
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Function
    
    If CheckStrValid(txt(0).Text, txt(0).MaxLength, strError) = False Then
        zlControl.TxtSelAll txt(0)
        txt(0).SetFocus
        Exit Function
    End If
    
    If CheckStrValid(txt(10).Text, txt(10).MaxLength, strError) = False Then
        zlControl.TxtSelAll txt(10)
        txt(10).SetFocus
        Exit Function
    End If
    
    If dtp(0).Value > dtp(1).Value Then
        strError = "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        dtp(0).SetFocus
        GoTo errHand
    End If
    
    If Abs(DateDiff("h", CDate(Format(dtp(0).Value, "YYYY-MM-DD HH:MM")), CDate(Format(dtp(1).Value, "YYYY-MM-DD HH:MM")))) > 12 Then
        strError = "������ʼʱ�����������ʱ��֮�䲻�ܴ���12Сʱ��"
        dtp(0).SetFocus
        GoTo errHand
    End If
    
    
    If dtp(2).Value > dtp(3).Value And chk(0).Value = 1 Then
        strError = "����ʼʱ�䲻�ܴ����������ʱ�䣡"
        dtp(2).SetFocus
        GoTo errHand
    End If
    
    If chk(0).Value = 1 And cbo(0).ListIndex = -1 Then
        strError = "����ָ������ʽ��"
        cbo(0).SetFocus
        GoTo errHand
    End If
    
    If chk(0).Value = 1 And cbo(1).ListIndex = -1 Then
        strError = "����ָ������������"
        cbo(1).SetFocus
        GoTo errHand
    End If
    
    If dtp(4).Value > dtp(5).Value And chk(1).Value = 1 Then
        strError = "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        dtp(4).SetFocus
        GoTo errHand
    End If
    
    If CheckAllNumber(txt(0).Text) = False Then
        strError = "��Һ��������Ϊȫ���֣�"
        
        zlControl.TxtSelAll txt(0)
        txt(0).SetFocus
        GoTo errHand
    End If
    
    Dim LngCount As Long
    
    LngCount = 0
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 And InStr(bill(2).TextMatrix(lngLoop, 0), "����ҽ��") > 0 Then
            LngCount = LngCount + 1
            If LngCount > 1 Then
                strError = "����ҽ��ֻ��һ����"
                bill(2).SetFocus
                GoTo errHand
            End If
            If LngCount > 2 Then Exit For
        End If
    Next
    If LngCount < 1 Then
        strError = " ����ָ������������ҽ����"
        bill(2).SetFocus
        GoTo errHand
    End If
        
    '������������Ƿ��зǷ��ַ�����������������
    For lngIndex = 0 To 1
        For lngLoop = 1 To bill(lngIndex).Rows - 1
            If bill(lngIndex).RowData(lngLoop) > 0 Then
                Exit For
            End If
        Next
            
        If lngLoop = bill(lngIndex).Rows Then
            If lngIndex = 0 Then
                strError = "������һ������������"
            Else
                strError = "������һ������������"
            End If
                        
            bill(lngIndex).SetFocus
            GoTo errHand
        End If
        
        For lngLoop = 1 To bill(lngIndex).Rows - 1
            If bill(lngIndex).RowData(lngLoop) > 0 Then
                If CheckStrValid(bill(lngIndex).TextMatrix(lngLoop, 1), 50, strError) = False Then
                    bill(lngIndex).Col = 1
                    bill(lngIndex).Row = lngLoop
                    bill(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
    
    '�����������Ƿ��зǷ��ַ�������
    For lngIndex = 0 To 1
        For lngLoop = 1 To bill2(lngIndex).Rows - 1
            If bill2(lngIndex).RowData(lngLoop) > 0 Then
                If CheckStrValid(bill2(lngIndex).TextMatrix(lngLoop, 2), 100, strError) = False Then
                    bill2(lngIndex).Col = 2
                    bill2(lngIndex).Row = lngLoop
                    bill2(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
               
    '�����Ա���롢��Ա�����Ƿ��зǷ��ַ�������
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            If CheckStrValid(bill(2).TextMatrix(lngLoop, 1), 20, strError) = False Then
                bill(2).Col = 1
                bill(2).Row = lngLoop
                bill(2).SetFocus
                Exit Function
            End If
            
            If CheckStrValid(bill(2).TextMatrix(lngLoop, 2), 10, strError) = False Then
                bill(2).Col = 1
                bill(2).Row = lngLoop
                bill(2).SetFocus
                Exit Function
            End If
            
        End If
    Next
    
    CheckDataValid = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, ReturnStrSQL As String, strError As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '���ܣ��������ݣ�����ֻ���γ�SQL��䣬��Ҫ�ڵ��ô�ר�õ��Ĵ�����ִ��
    '������
    '---------------------------------------------------------------------------------------------------------
    Dim str����ʽ As String
    Dim str�������� As String
    Dim strTmp As String
    Dim LngCount As Long
    Dim rs As New ADODB.Recordset
    
    Dim strSQL() As String
    Dim lngLoop As Long
    
    On Error GoTo ErrHandle
    
    '����������ݵ���Ч��
    If CheckDataValid(strError) = False Then Exit Function
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_�����������_DELETE(" & mlng������¼id & ",1)"
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_�����������_DELETE(" & mlng������¼id & ",2)"
    
    If mlngҽ��id > 0 Then
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_������ϼ�¼_DELETE2(" & mlngҽ��id & ",8)"
    
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_������ϼ�¼_DELETE2(" & mlngҽ��id & ",9)"
        
        mlngOrderID = mlngҽ��id
    Else
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_������ϼ�¼_DELETE(" & lng����ID & "," & lng��ҳID & ",1," & lng����ID & ",'8,9')"
    End If
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_����������Ա_CANCELPERSON(" & mlng������¼id & ")"
    
    
    'ReDimArray lngCount, strSQL
    'strSQL(lngCount) = "ZL_����������¼_DELETE(" & mlng������¼id & ")"
    
    'ʼ������,��Ϊ�ڵ��ô˽�Ʒǰ������ɾ��
    
    If cbo(0).Text <> "" Then
        str����ʽ = Trim(Mid(cbo(0).Text, 1, InStr(cbo(0).Text, Space(200)) - 1))
        str�������� = Trim(Mid(cbo(0).Text, InStr(cbo(0).Text, Space(200)) + 200))
    End If
    
    If mlng������¼id = 0 Then
        mlng������¼id = zlDatabase.GetNextId("����������¼")
        
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_����������¼_INSERT(" & mlng������¼id & "," & _
                                    lng����ID & "," & _
                                    IIf(lng��ҳID = 0, "NULL", lng��ҳID) & "," & _
                                    IIf(mlngOrderID > 0, mlngOrderID, "Null") & "," & _
                                    lng����ID & "," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(1).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(2).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(3).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(str����ʽ) & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & str�������� & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(cbo(1).Text) & "'", "NULL") & ",'" & _
                                    txt(0).Text & "'," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                    txt(10).Text & "'," & _
                                    mlng������id & ",'" & _
                                    cbo(2).Text & "')"
    Else
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_����������¼_UPDATE(" & mlng������¼id & "," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(1).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(2).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(3).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(str����ʽ) & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & str�������� & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(cbo(1).Text) & "'", "NULL") & ",'" & _
                                    txt(0).Text & "'," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                    txt(10).Text & "'," & _
                                    mlng������id & ",'" & _
                                    cbo(2).Text & "'," & lng����ID & ")"
    End If
    
    '��д��������
    For lngLoop = 1 To bill(0).Rows - 1
        If bill(0).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_�����������_INSERT(" & mlng������¼id & ",1," & IIf(bill(0).TextMatrix(lngLoop, 2) = "��", 1, 0) & ",'" & bill(0).TextMatrix(lngLoop, 1) & "'," & IIf(Val(Mid(bill(0).TextMatrix(lngLoop, 0), 1, 1)) = 1, "NULL," & bill(0).RowData(lngLoop), bill(0).RowData(lngLoop) & ",NULL") & ")"
        End If
    Next
    '��д��������
    For lngLoop = 1 To bill(1).Rows - 1
        If bill(1).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_�����������_INSERT(" & mlng������¼id & ",2," & IIf(bill(1).TextMatrix(lngLoop, 2) = "��", 1, 0) & ",'" & bill(1).TextMatrix(lngLoop, 1) & "'," & IIf(Val(Mid(bill(1).TextMatrix(lngLoop, 0), 1, 1)) = 1, "NULL," & bill(1).RowData(lngLoop), bill(1).RowData(lngLoop) & ",NULL") & ")"
        End If
    Next
    
    '��д��ǰ���
    For lngLoop = 1 To bill2(0).Rows - 1
        If bill2(0).RowData(lngLoop) > 0 And (Val(bill2(0).TextMatrix(lngLoop, 3)) > 0 Or Val(bill2(0).TextMatrix(lngLoop, 4)) > 0) Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_������ϼ�¼_INSERT(" & lng����ID & "," & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ",1," & lng����ID & ",8," & Val(bill2(0).TextMatrix(lngLoop, 4)) & "," & Val(bill2(0).TextMatrix(lngLoop, 3)) & ",NULL,'" & bill2(0).TextMatrix(lngLoop, 2) & "',NULL,NULL,NULL,SYSDATE," & IIf(mlngOrderID = 0, "NULL", mlngOrderID) & ")"
        End If
    Next
    
    '��д�������
    For lngLoop = 1 To bill2(1).Rows - 1
        If bill2(1).RowData(lngLoop) > 0 And (Val(bill2(1).TextMatrix(lngLoop, 3)) > 0 Or Val(bill2(1).TextMatrix(lngLoop, 4)) > 0) Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_������ϼ�¼_INSERT(" & lng����ID & "," & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ",1," & lng����ID & ",9," & Val(bill2(1).TextMatrix(lngLoop, 4)) & "," & Val(bill2(1).TextMatrix(lngLoop, 3)) & ",NULL,'" & bill2(1).TextMatrix(lngLoop, 2) & "',NULL,NULL,NULL,SYSDATE," & IIf(mlngOrderID = 0, "NULL", mlngOrderID) & ")"
        End If
    Next
    
    '��д������Ա
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_����������¼_PERSON(" & mlng������¼id & ",'" & bill(2).TextMatrix(lngLoop, 0) & "'," & bill(2).RowData(lngLoop) & ",'" & bill(2).TextMatrix(lngLoop, 2) & "','" & bill(2).TextMatrix(lngLoop, 1) & "')"
        End If
    Next
    
    strTmp = ""
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    '����SQL���
    ReturnStrSQL = strTmp
        
    SaveData = True
    
    Exit Function
    
ErrHandle:
    
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    strError = Err.Description
    Call SaveErrLog
    
End Function

Private Sub SetDefault(ByVal objBill As BillEdit, ByVal intCol As Integer)
    '
    '���ܣ�
    '
    Dim lngLoop As Long
    
    For lngLoop = 1 To objBill.Rows - 1
        If objBill.RowData(lngLoop) > 0 Then
            If objBill.TextMatrix(lngLoop, intCol) = "��" Then
                Exit For
            End If
        End If
    Next
    
    If lngLoop = objBill.Rows And objBill.RowData(1) > 0 Then
        objBill.TextMatrix(1, intCol) = "��"
    End If
    
End Sub

Private Sub bill_AfterDeleteRow(Index As Integer)
    If Index <> 2 Then SetDefault bill(Index), 2
End Sub

Private Sub bill_CellCheck(Index As Integer, Row As Long, Col As Long)
    Dim lngLoop As Long
    
    If bill(Index).TextMatrix(Row, Col) = "" Then
        SetDefault bill(Index), 2
    Else
        For lngLoop = 1 To bill(Index).Rows - 1
            If lngLoop <> Row Then bill(Index).TextMatrix(lngLoop, Col) = ""
        Next
    End If
End Sub

Private Sub bill_CommandClick(Index As Integer)
    Dim sglX As Single
    Dim sglY As Single
    
    If bill(Index).TextMatrix(bill(Index).Row, 0) <> "" Then
        Select Case Index
        Case 0, 1
            If PopOperateSelect(bill(Index), Val(Mid(bill(Index).TextMatrix(bill(Index).Row, 0), 1, 1))) Then
                SetDefault bill(Index), 2
            End If
        Case 2
            CalcPosition sglX, sglY, bill(Index)
            
            Call ShowDownListPerson(Screen, bill(Index), 0, sglX, sglY, False)
            
        End Select
    End If
    
End Sub

Private Sub bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    If Index = 0 Or Index = 1 Then SetDefault bill(Index), 2
End Sub

Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim sglX As Single
    Dim sglY As Single
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If bill(Index).Col = 1 And Trim(bill(Index).TextMatrix(bill(Index).Row, 1)) = "" And bill(Index).TxtVisible = False Then
        zlCommFun.PressKey vbKeyTab
    End If
    
    Select Case bill(Index).Col
    Case 0
        If bill(Index).TextMatrix(bill(Index).Row, 0) <> bill(Index).List(bill(Index).ListIndex) Then
            
            Call ClearSpecRowCol(bill(Index), bill(Index).Row, Array())
            
            bill(Index).TextMatrix(bill(Index).Row, 0) = bill(Index).List(bill(Index).ListIndex)
                        
        End If
    End Select
    
    If bill(Index).TxtVisible = False Then Exit Sub
            
    If Trim(bill(Index).Text) <> "" Then
        Select Case Index
        Case 0, 1
            Cancel = Not ShowDownListOperate(Screen, Val(Mid(bill(Index).TextMatrix(bill(Index).Row, 0), 1, 1)), bill(Index), True, mstr�Ա�)
            If Cancel = False Then Call SetDefault(bill(Index), 2)
        Case 2
            Call CalcPosition(sglX, sglY, bill(Index))
            Cancel = Not ShowDownListPerson(Screen, bill(Index), 0, sglX, sglY, True)
            If Cancel = False Then
                
            End If
        End Select
    Else
        If bill(Index).Col = 1 And bill(Index).RowData(bill(Index).Row) = 0 Then zlCommFun.PressKey vbKeyTab
    End If
    
End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub bill_LostFocus(Index As Integer)
    bill(Index).CmdVisible = False
'    bill(Index).CboVisible = False
End Sub


Private Sub bill2_CommandClick(Index As Integer)
    If PopSelect(bill2(Index), mstr�Ա�) Then
        
    End If
End Sub

Private Sub bill2_EditKeyPress(Index As Integer, KeyAscii As Integer)
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub bill2_EnterCell(Index As Integer, Row As Long, Col As Long)
    If bill2(Index).TextMatrix(Row, Col) = "" And Col <> 2 Then bill2(Index).TextMatrix(Row, Col) = " "
End Sub

Private Sub bill2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim sglX As Single
    Dim sglY As Single
        
    If KeyCode <> 13 Then Exit Sub
    
    If bill2(Index).TxtVisible = False Then Exit Sub
            
    Call CalcPosition(sglX, sglY, bill2(Index))

    If Trim(bill2(Index).Text) <> "" And bill2(Index).Col = 0 Or bill2(Index).Col = 1 Then
        Cancel = Not ShowDownList2(Screen, bill2(Index).Col + 1, bill2(Index), sglX, sglY, True)
        If Cancel = False Then
            
        End If
    End If
    
End Sub

Private Sub bill2_LostFocus(Index As Integer)
    bill2(Index).CmdVisible = False
    bill2(Index).CboVisible = False
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 0 Then
        dtp(2).Enabled = chk(Index).Value
        dtp(3).Enabled = chk(Index).Value
        
        cbo(0).Enabled = dtp(2).Enabled
        cbo(1).Enabled = dtp(2).Enabled
        
        If cbo(1).Enabled = False Then
            cbo(1).ListIndex = -1
        ElseIf cbo(1).ListIndex = -1 Then
            cbo(1).ListIndex = 0
        End If
        
        txt(3).Visible = Not dtp(2).Enabled
        txt(1).Visible = Not dtp(3).Enabled
        
    Else
        dtp(4).Enabled = chk(Index).Value
        dtp(5).Enabled = chk(Index).Value
        txt(2).Visible = Not dtp(4).Enabled
        txt(4).Visible = Not dtp(5).Enabled
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub UserControl_Initialize()

    '��ʼ���ؼ�����
    Dim lngLoop As Long
    
    On Error GoTo ErrHandle
    
    For lngLoop = 0 To 1
        With bill(lngLoop)
            .Cols = 3
            .TextMatrix(0, 0) = "���뷽ʽ"
            .TextMatrix(0, 1) = "��������"
            .TextMatrix(0, 2) = "ȱʡ"
            .ColWidth(0) = 855
            .ColWidth(1) = 2220
            .ColWidth(2) = 600
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 4
            .ColData(0) = 3
            .ColData(1) = 1
            .ColData(2) = -1
            .AddItem "1-����"
            .AddItem "2-����"
            .MsfObj.GridLinesFixed = flexGridFlat
            .MsfObj.GridColor = &H8000000C
            .MsfObj.GridColorFixed = &H8000000C
            .MsfObj.BackColorFixed = &HFFFFFF
            .Active = True
        End With
    Next
    
    For lngLoop = 0 To 1
        With bill2(lngLoop)
            .Cols = 5
            .TextMatrix(0, 0) = "��������"
            .TextMatrix(0, 1) = "��ϱ���"
            .TextMatrix(0, 2) = "�������"
            .TextMatrix(0, 3) = "���ID"
            .TextMatrix(0, 4) = "����ID"
            .ColWidth(0) = 900
            .ColWidth(1) = 900
            .ColWidth(2) = 1890
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            
            .ColData(0) = 1
            .ColData(1) = 1
            .ColData(2) = 4
            .ColData(3) = 5
            .ColData(4) = 5
            .PrimaryCol = 2
            .MsfObj.GridLinesFixed = flexGridFlat
            .MsfObj.GridColor = &H8000000C
            .MsfObj.GridColorFixed = &H8000000C
            .MsfObj.BackColorFixed = &HFFFFFF
            .Active = True
        End With
    Next
    
    With bill(2)
        .Cols = 3
        
        .TextMatrix(0, 0) = "��λ"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "���"
        
        .ColWidth(0) = 900
        .ColWidth(1) = 900
        .ColWidth(2) = 0
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColData(0) = 3
        .ColData(1) = 1
        .ColData(2) = 5
        .MsfObj.GridLinesFixed = flexGridFlat
        .MsfObj.GridColor = &H8000000C
        .MsfObj.GridColorFixed = &H8000000C
        .MsfObj.BackColorFixed = &HFFFFFF
        bill(2).Active = True
    End With
    
    Exit Sub
    
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
    
    '���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub UserControl_InitProperties()
    '��ʼ���˲���Ϊ0
    mlng����id = 0
    mDispMode = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
End Sub

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Resize()
    UserControl.Width = 8130
    UserControl.Height = 5985
End Sub

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
End Sub

Private Sub UserControl_Show()
    Dim objCtl As Control
    Dim rs As New ADODB.Recordset
    
    
    '��д������λ�б�
    gstrSql = "SELECT ����,0 FROM ������λ"
    Call zlDatabase.OpenRecordset(rs, gstrSql, "������Ҫ")
    If rs.RecordCount = 0 Then
        MsgBox "ϵͳ���ݲ�������û��������λ���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If rs.BOF = False Then Call AddComboData(bill(2), rs, False)
    
    'ֻ������ʱ��ʾ
    If Ambient.UserMode = True And InDesign = False Then
        If mDispMode Then
            For Each objCtl In Controls
                If UCase(TypeName(objCtl)) <> UCase("ImageList") Then
                    objCtl.Enabled = False
                End If
            Next
        End If
    End If
    
    If mblnLoaded = False Then
        InitData
        Call ReadData
    End If
    
    mblnLoaded = True
End Sub

Public Property Get Text() As String
    'Ϊÿһ���ؼ������ı�ת������
    Dim lngLoop As Long
    Dim strTmp As String
    
    'ͨ���û���������ݵõ�ת���ı�
    strTmp = "��ǰ��ϣ�" & vbCrLf
    For lngLoop = 1 To bill2(0).Rows - 1
        If bill2(0).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill2(0).TextMatrix(lngLoop, 2) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "������ϣ�" & vbCrLf
    For lngLoop = 1 To bill2(1).Rows - 1
        If bill2(1).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill2(1).TextMatrix(lngLoop, 2) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "����������" & vbCrLf
    For lngLoop = 1 To bill(0).Rows - 1
        If bill(0).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill(0).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "����������" & vbCrLf
    For lngLoop = 1 To bill(1).Rows - 1
        If bill(1).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill(1).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next

    strTmp = strTmp & "������ʼ��" & Format(dtp(0).Value, "YYYY��MM��DD�� HHʱMM��") & "   "
    strTmp = strTmp & "������ֹ��" & Format(dtp(1).Value, "YYYY��MM��DD�� HHʱMM��") & vbCrLf

    If chk(0).Value <> 0 Then
        strTmp = strTmp & "����ʼ��" & Format(dtp(2).Value, "YYYY��MM��DD�� HHʱMM��") & "   "
        strTmp = strTmp & "������ֹ��" & Format(dtp(3).Value, "YYYY��MM��DD�� HHʱMM��") & vbCrLf
        strTmp = strTmp & "����ʽ��" & cbo(0).Text & "   "
        strTmp = strTmp & "����������" & cbo(1).Text & vbCrLf
    End If
    
    If chk(1).Value <> 0 Then
        strTmp = strTmp & "������ʼ��" & Format(dtp(4).Value, "YYYY��MM��DD�� HHʱMM��") & "   "
        strTmp = strTmp & "������ֹ��" & Format(dtp(5).Value, "YYYY��MM��DD�� HHʱMM��") & vbCrLf
    End If
    
    strTmp = strTmp & "������Ա��" & vbCrLf
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            strTmp = strTmp & bill(2).TextMatrix(lngLoop, 0) & "          " & bill(2).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next

    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub

Private Function CheckAllNumber(ByVal strKey As String) As Boolean
    
    Dim lngLoop As Long
    
    For lngLoop = 1 To Len(strKey)
        If Mid(strKey, lngLoop, 1) < "0" Or Mid(strKey, lngLoop, 1) > "9" Then
            Exit Function
        End If
    Next
    
    CheckAllNumber = True
End Function

Private Function CheckHave(ByVal objBill As Object, ByVal intRow As Integer, ByVal lngKey As Long) As Boolean
    Dim lngLoop As Long
    
    For lngLoop = 1 To objBill.Rows - 1
        If objBill.RowData(lngLoop) = lngKey And lngLoop <> intRow Then
            CheckHave = True
            Exit Function
        End If
    Next
    
    CheckHave = False
End Function

