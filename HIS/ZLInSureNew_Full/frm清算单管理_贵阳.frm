VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm���㵥����_���� 
   Caption         =   "���㵥����_����"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   10575
   Icon            =   "frm���㵥����_����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10575
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabShow 
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   720
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "ҽ�Ʊ���"
      TabPicture(0)   =   "frm���㵥����_����.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "frm���㵥����_����.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "���˱���"
      TabPicture(2)   =   "frm���㵥����_����.frx":07A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSComctlLib.ImageList imgBlack 
      Left            =   2820
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":07BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":0E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   2250
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":1026
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":1240
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���㵥����_����.frx":1674
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10575
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrTool"
      MinHeight1      =   645
      Width1          =   915
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgBlack"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�걨"
               Key             =   "Add"
               Object.ToolTipText     =   "�걨���㵥"
               Object.Tag             =   "�걨"
               ImageIndex      =   1
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add1"
                     Text            =   "ҽ�Ʊ���"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add2"
                     Text            =   "��������"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add3"
                     Text            =   "���˱���"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Object.ToolTipText     =   "�������㵥"
               Object.Tag             =   "����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���㵥����_����.frx":188E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   5385
      Left            =   30
      TabIndex        =   4
      Top             =   1050
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9499
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   13275520
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd1 
         Caption         =   "ҽ�Ʊ����걨����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAdd2 
         Caption         =   "���������걨����"
      End
      Begin VB.Menu mnuEditAdd3 
         Caption         =   "���˱����걨����"
      End
      Begin VB.Menu muuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "�������㵥(&B)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu muuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "�����걨��(&V)"
      End
      Begin VB.Menu mnuEditGet 
         Caption         =   "��ѯ�������(&G)"
      End
      Begin VB.Menu muuEditSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFilter 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu muuEditSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frm���㵥����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrFilter As String
Private Enum ҽ�Ʊ���
    ID
    �ں�
    ����������
    �������
    ����Ա
    ����
    �����˴�
    ��������ʻ�
    ����ҽ�Ʋ���
    ���������˴�
    ������������ʻ�
    �����������ͳ��
    ���������ͳ��
    ��������ҽ�Ʋ���
    ������סԺ�˴�
    ������סԺ�����ʻ�
    ������סԺ����ͳ��
    ������סԺ���ͳ��
    ������סԺҽ�Ʋ���
    ��֢סԺ�˴�
    ��֢סԺ�����ʻ�
    ��֢סԺ����ͳ��
    ��֢סԺ���ͳ��
    ��֢סԺҽ�Ʋ���
    �հ���סԺ�˴�
    �հ���סԺ����
    �հ���סԺ�����ʻ�
    �հ���סԺҽ�Ʋ���
    ���ɽ����˴�
    ���ɽ�������ʻ�
    ���ɽ������ͳ��
    ���ɽ�����ͳ��
    ���ɽ���ҽ�Ʋ���
    ������ˮ��
    �������
    ����
End Enum

Private Enum ��������
    ID
    �ں�
    ����������
    �������
    ����Ա
    ����
    ��������˴�
    ������ɷ����ܶ�
    �������ͳ��֧��
    ����ǰ����˴�
    ����ǰ��ɷ����ܶ�
    ����ǰ���ͳ��֧��
    �����˴�
    ���������ܶ�
    ����ͳ��֧��
    ������ˮ��
    �������
    ����
End Enum

Private Enum ���˱���
    ID
    �ں�
    ����������
    �������
    ����Ա
    ����
    �����˴�
    ����ͳ��֧��
    סԺ�˴�
    סԺͳ��֧��
    ������ˮ��
    �������
    ����
End Enum

Private Enum ҳ��
    ҽ�Ʊ���    '������
    ��������
    ���˱���
End Enum

'�걨�������˵��
'1��ҽ�Ʊ����걨�嵥�У������˴���ָ����ͨ������˴Σ����ɽ�������˴���ָ��ͨ������ѡ��Ľ��㷽ʽΪ�����ְ��ɵĲ�������
'   a�������ߣ�����=1������֢������=2�������հ��ɣ�����=4�������ɣ�����=6��
'   b����ͨ����ѡ���˵����ֵľ����������
'   c�������걨�嵥��
'2�������걨�����У�
'   a������סԺ���ɣ��������Ϊ��������Ժ��ʽ���Ǽƻ������ģ�����=5��
'   b����������Ժ��ʽΪ�ƻ�������
'   c���ǰ��ɣ�����������������-�������-������

Public Sub ShowME(ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim strMonth As String
    
    'ȱʡֻ��ȡ�������ڵ�����
    strMonth = Format(DateAdd("m", -1, zlDatabase.Currentdate()), "yyyyMM")
    mstrFilter = " And A.�ں�>='" & strMonth & "'"
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With mshDetail
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
End Sub

Private Sub mnuEditAdd1_Click()
    If Not frmҽ�Ʊ����걨��.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditAdd2_Click()
    If Not frm���������걨��.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditAdd3_Click()
    If Not frm���˱����걨��.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditDel_Click()
    Dim lngID As Long
    Dim int���������� As Integer
    Dim str��ˮ�� As String
    On Error GoTo errHand
    
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ҽ�Ʊ���.������ˮ��)
        int���������� = Val(mshDetail.TextMatrix(mshDetail.Row, ҽ�Ʊ���.����������))
    ElseIf tabShow.Tab = ҳ��.�������� Then
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ��������.������ˮ��)
        int���������� = Val(mshDetail.TextMatrix(mshDetail.Row, ��������.����������))
    Else
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ���˱���.������ˮ��)
        int���������� = Val(mshDetail.TextMatrix(mshDetail.Row, ���˱���.����������))
    End If
    
    If MsgBox("��ȷ��Ҫ��������㵥��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    If Not InitXML Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "APPNO", str��ˮ��)
    If tabShow.Tab <> ҳ��.���˱��� Then Call InsertChild(mdomInput.documentElement, "INSURETYPE", int����������)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    '���ýӿ�
    If CommRecServer(IIf(tabShow.Tab = ҳ��.ҽ�Ʊ���, "DELRECM", IIf(tabShow.Tab = ҳ��.��������, "DELRECB", "DELRECG"))) = False Then Exit Sub
    
    gstrSQL = "ZL_���㵥_DELETE(" & lngID & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditFilter_Click()
    Dim strReturn As String
    strReturn = frm���㵥_����.ShowCondition
    If strReturn = "" Then Exit Sub
    
    mstrFilter = strReturn
    Call RefreshData
End Sub

Private Sub mnuEditGet_Click()
    Dim lngID As Long
    Dim str��ˮ�� As String, str������� As String
    On Error GoTo errHand
    
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ҽ�Ʊ���.������ˮ��)
    ElseIf tabShow.Tab = ҳ��.�������� Then
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ��������.������ˮ��)
    Else
        str��ˮ�� = mshDetail.TextMatrix(mshDetail.Row, ���˱���.������ˮ��)
    End If
    
    If Val(mshDetail.TextMatrix(mshDetail.Row, 0)) = 0 Then Exit Sub
    If str��ˮ�� = "" Then Exit Sub
    
    If Not InitXML Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "APPNO", str��ˮ��)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    '���ýӿ�
    If CommRecServer("QUERYREC") = False Then Exit Sub
    str������� = GetElemnetValue("STATUS")
    gstrSQL = "ZL_���㵥_UPDATE(" & lngID & ",'" & str������� & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditRefresh_Click()
    Call RefreshData
End Sub

Private Sub mnuEditView_Click()
    Dim lngID As Long
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    
    If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
        Call frmҽ�Ʊ����걨��.ShowME(lngID)
    ElseIf tabShow.Tab = ҳ��.�������� Then
        Call frm���������걨��.ShowME(lngID)
    Else
        Call frm���˱����걨��.ShowME(lngID)
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mshDetail_DblClick()
    Call mnuEditView_Click
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call mnuEditView_Click
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call RefreshData
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Add"
        If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
            Call mnuEditAdd1_Click
        ElseIf tabShow.Tab = ҳ��.�������� Then
            Call mnuEditAdd2_Click
        Else
            Call mnuEditAdd3_Click
        End If
    Case "Del"
        Call mnuEditDel_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Filter"
        Call mnuEditFilter_Click
    End Select
End Sub

Private Sub RefreshData()
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Call InitBill
    
    '�����������
    If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.�ں�, A.�������,A.�����������, A.����Ա, A.���� ,B.�����˴�, B.��������ʻ�, B.����ҽ�Ʋ���, B.���������˴�, B.������������ʻ�, B.�����������ͳ��, B.����������ͳ��,  " & _
                 "        B.��������ҽ�Ʋ���, B.������סԺ�˴�, B.������סԺ�����ʻ�, B.������סԺ����ͳ��, B.������סԺ���ͳ��, B.������סԺҽ�Ʋ���,  " & _
                 "        B.��֢סԺ�˴�, B.��֢סԺ�����ʻ�, B.��֢סԺ����ͳ��, B.��֢סԺ���ͳ��, B.��֢סԺҽ�Ʋ���, B.�հ���סԺ�˴�, B.�հ���סԺ����,  " & _
                 "        B.�հ���סԺ�����ʻ�, �հ���סԺҽ�Ʋ���, B.���ɽ����˴�, B.���ɽ�������ʻ�, B.���ɽ������ͳ��, B.���ɽ�����ͳ��, B.���ɽ���ҽ�Ʋ���, A.������ˮ��, A.������� " & _
                 " FROM ���㵥 A, ����ҽ��������ϸ B " & _
                 " WHERE A.ID=B.���㵥ID " & mstrFilter & " AND A.����=" & tabShow.Tab & _
                 " Order by A.�ں� Desc,A.�����������"
    ElseIf tabShow.Tab = ҳ��.�������� Then
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.�ں�, A.�������,A.�����������, A.����Ա, A.���� , B.��������˴�, B.������ɷ����ܶ�, B.�������ͳ��֧��, B.����ǰ����˴�,  " & _
                 "        B.����ǰ��ɷ����ܶ�, B.����ǰ���ͳ��֧��, B.�����˴�, B.���������ܶ�, B.����ͳ��֧��, A.������ˮ��, A.������� " & _
                 " FROM ���㵥 A, ����������ϸ B" & _
                 " WHERE A.ID=B.���㵥ID " & mstrFilter & " And A.����=" & tabShow.Tab & _
                 " Order by A.�ں� Desc,A.�����������"
    Else
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.�ں�, A.�������,A.�����������, A.����Ա, A.���� , B.�����˴�, B.����ͳ��֧��,  B.סԺ�˴�, B.סԺͳ��֧��, A.������ˮ��, A.������� " & _
                 " FROM ���㵥 A, ����������ϸ B" & _
                 " WHERE A.ID=B.���㵥ID " & mstrFilter & " And A.����=" & tabShow.Tab & _
                 " Order by A.�ں� Desc,A.�����������"
    End If
    Call OpenRecordset_OtherBase(rsTemp, "�����������", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then Set mshDetail.DataSource = rsTemp
    Call InitBill(False)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitBill(Optional ByVal blnInit As Boolean = True)
    With mshDetail
        If tabShow.Tab = ҳ��.ҽ�Ʊ��� Then
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = ҽ�Ʊ���.����
                
                .TextMatrix(0, ҽ�Ʊ���.ID) = "ID"
                .TextMatrix(0, ҽ�Ʊ���.�ں�) = "�ں�"
                .TextMatrix(0, ҽ�Ʊ���.����������) = "����������"
                .TextMatrix(0, ҽ�Ʊ���.�������) = "�������"
                .TextMatrix(0, ҽ�Ʊ���.����Ա) = "����Ա"
                .TextMatrix(0, ҽ�Ʊ���.����) = "����"
                .TextMatrix(0, ҽ�Ʊ���.�����˴�) = "�����˴�"
                .TextMatrix(0, ҽ�Ʊ���.��������ʻ�) = "��������ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.����ҽ�Ʋ���) = "����ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.���������˴�) = "���������˴�"
                .TextMatrix(0, ҽ�Ʊ���.������������ʻ�) = "�����ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.�����������ͳ��) = "����ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.���������ͳ��) = "��ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.��������ҽ�Ʋ���) = "ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.������סԺ�˴�) = "�������˴�"
                .TextMatrix(0, ҽ�Ʊ���.������סԺ�����ʻ�) = "�����ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.������סԺ����ͳ��) = "����ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.������סԺ���ͳ��) = "���ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.������סԺҽ�Ʋ���) = "ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.��֢סԺ�˴�) = "��֢סԺ�˴�"
                .TextMatrix(0, ҽ�Ʊ���.��֢סԺ�����ʻ�) = "�����ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.��֢סԺ����ͳ��) = "����ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.��֢סԺ���ͳ��) = "���ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.��֢סԺҽ�Ʋ���) = "ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.�հ���סԺ�˴�) = "�հ���סԺ�˴�"
                .TextMatrix(0, ҽ�Ʊ���.�հ���סԺ����) = "סԺ����"
                .TextMatrix(0, ҽ�Ʊ���.�հ���סԺ�����ʻ�) = "�����ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.�հ���סԺҽ�Ʋ���) = "ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.���ɽ����˴�) = "���ɽ����˴�"
                .TextMatrix(0, ҽ�Ʊ���.���ɽ�������ʻ�) = "�����ʻ�"
                .TextMatrix(0, ҽ�Ʊ���.���ɽ������ͳ��) = "����ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.���ɽ�����ͳ��) = "���ͳ��"
                .TextMatrix(0, ҽ�Ʊ���.���ɽ���ҽ�Ʋ���) = "ҽ�Ʋ���"
                .TextMatrix(0, ҽ�Ʊ���.������ˮ��) = "������ˮ��"
                .TextMatrix(0, ҽ�Ʊ���.�������) = "�������"
            End If
            .ColWidth(ҽ�Ʊ���.ID) = 0
            .ColWidth(ҽ�Ʊ���.�ں�) = 800
            .ColWidth(ҽ�Ʊ���.����������) = 0
            .ColWidth(ҽ�Ʊ���.�������) = 1200
            .ColWidth(ҽ�Ʊ���.����Ա) = 1000
            .ColWidth(ҽ�Ʊ���.����) = 1000
            .ColWidth(ҽ�Ʊ���.�����˴�) = 1000
            .ColWidth(ҽ�Ʊ���.��������ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.����ҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.���������˴�) = 1400
            .ColWidth(ҽ�Ʊ���.������������ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.�����������ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.���������ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.��������ҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.������סԺ�˴�) = 1400
            .ColWidth(ҽ�Ʊ���.������סԺ�����ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.������סԺ����ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.������סԺ���ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.������סԺҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.��֢סԺ�˴�) = 1400
            .ColWidth(ҽ�Ʊ���.��֢סԺ�����ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.��֢סԺ����ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.��֢סԺ���ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.��֢סԺҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.�հ���סԺ�˴�) = 1600
            .ColWidth(ҽ�Ʊ���.�հ���סԺ����) = 1000
            .ColWidth(ҽ�Ʊ���.�հ���סԺ�����ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.�հ���סԺҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.���ɽ����˴�) = 1400
            .ColWidth(ҽ�Ʊ���.���ɽ�������ʻ�) = 1000
            .ColWidth(ҽ�Ʊ���.���ɽ������ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.���ɽ�����ͳ��) = 1000
            .ColWidth(ҽ�Ʊ���.���ɽ���ҽ�Ʋ���) = 1000
            .ColWidth(ҽ�Ʊ���.������ˮ��) = 2000
            .ColWidth(ҽ�Ʊ���.�������) = 2500
        ElseIf tabShow.Tab = ҳ��.�������� Then
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = ��������.����
                
                .TextMatrix(0, ��������.ID) = "ID"
                .TextMatrix(0, ��������.�ں�) = "�ں�"
                .TextMatrix(0, ��������.����������) = "����������"
                .TextMatrix(0, ��������.�������) = "�������"
                .TextMatrix(0, ��������.����Ա) = "����Ա"
                .TextMatrix(0, ��������.����) = "����"
                .TextMatrix(0, ��������.��������˴�) = "��������˴�"
                .TextMatrix(0, ��������.������ɷ����ܶ�) = "�����ܶ�"
                .TextMatrix(0, ��������.�������ͳ��֧��) = "ͳ��֧��"
                .TextMatrix(0, ��������.����ǰ����˴�) = "����ǰ����˴�"
                .TextMatrix(0, ��������.����ǰ��ɷ����ܶ�) = "�����ܶ�"
                .TextMatrix(0, ��������.����ǰ���ͳ��֧��) = "ͳ��֧��"
                .TextMatrix(0, ��������.�����˴�) = "�����˴�"
                .TextMatrix(0, ��������.���������ܶ�) = "�����ܶ�"
                .TextMatrix(0, ��������.����ͳ��֧��) = "ͳ��֧��"
                .TextMatrix(0, ��������.������ˮ��) = "������ˮ��"
                .TextMatrix(0, ��������.�������) = "�������"
            End If
            .ColWidth(��������.ID) = 0
            .ColWidth(��������.�ں�) = 800
            .ColWidth(��������.����������) = 0
            .ColWidth(��������.�������) = 1200
            .ColWidth(��������.����Ա) = 1000
            .ColWidth(��������.����) = 1000
            .ColWidth(��������.��������˴�) = 1400
            .ColWidth(��������.������ɷ����ܶ�) = 1000
            .ColWidth(��������.�������ͳ��֧��) = 1000
            .ColWidth(��������.����ǰ����˴�) = 1600
            .ColWidth(��������.����ǰ��ɷ����ܶ�) = 1000
            .ColWidth(��������.����ǰ���ͳ��֧��) = 1000
            .ColWidth(��������.�����˴�) = 1000
            .ColWidth(��������.���������ܶ�) = 1000
            .ColWidth(��������.����ͳ��֧��) = 1000
            .ColWidth(��������.������ˮ��) = 2000
            .ColWidth(��������.�������) = 2500
        Else
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = ���˱���.����
                
                .TextMatrix(0, ���˱���.ID) = "ID"
                .TextMatrix(0, ���˱���.�ں�) = "�ں�"
                .TextMatrix(0, ���˱���.����������) = "����������"
                .TextMatrix(0, ���˱���.�������) = "�������"
                .TextMatrix(0, ���˱���.����Ա) = "����Ա"
                .TextMatrix(0, ���˱���.����) = "����"
                .TextMatrix(0, ���˱���.�����˴�) = "�����˴�"
                .TextMatrix(0, ���˱���.����ͳ��֧��) = "ͳ��֧��"
                .TextMatrix(0, ���˱���.סԺ�˴�) = "סԺ�˴�"
                .TextMatrix(0, ���˱���.סԺͳ��֧��) = "ͳ��֧��"
                .TextMatrix(0, ���˱���.������ˮ��) = "������ˮ��"
                .TextMatrix(0, ���˱���.�������) = "�������"
            End If
            .ColWidth(���˱���.ID) = 0
            .ColWidth(���˱���.�ں�) = 800
            .ColWidth(���˱���.����������) = 0
            .ColWidth(���˱���.�������) = 0
            .ColWidth(���˱���.����Ա) = 1000
            .ColWidth(���˱���.����) = 1000
            .ColWidth(���˱���.�����˴�) = 1400
            .ColWidth(���˱���.����ͳ��֧��) = 1000
            .ColWidth(���˱���.סԺ�˴�) = 1000
            .ColWidth(���˱���.סԺͳ��֧��) = 1000
            .ColWidth(���˱���.������ˮ��) = 2000
            .ColWidth(���˱���.�������) = 2500
        End If
    End With
End Sub

Private Sub tbrTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Add1"
        Call mnuEditAdd1_Click
    Case "Add2"
        Call mnuEditAdd2_Click
    Case "Add3"
        Call mnuEditAdd3_Click
    End Select
End Sub
