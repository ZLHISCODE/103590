VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCollectMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmCollectMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picHLTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   0
      Width           =   4125
      Begin ZL9BillEdit.BillEdit billHLTime 
         Height          =   4215
         Left            =   -30
         TabIndex        =   6
         Top             =   -30
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   7435
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
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   4185
      Left            =   990
      ScaleHeight     =   4155
      ScaleWidth      =   4095
      TabIndex        =   3
      Top             =   600
      Width           =   4125
      Begin ZL9BillEdit.BillEdit billTime 
         Height          =   4215
         Left            =   -30
         TabIndex        =   4
         Top             =   -30
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   7435
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
   Begin VB.PictureBox picNodes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   240
      ScaleHeight     =   4215
      ScaleWidth      =   4155
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   750
      Width           =   4185
      Begin MSComctlLib.TreeView tvwNodes 
         Height          =   4215
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   7435
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   450
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgIcon"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   1350
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCollectMan.frx":6852
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl CollectPage 
      Height          =   4365
      Left            =   -120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5460
      _Version        =   589884
      _ExtentX        =   9631
      _ExtentY        =   7699
      _StockProps     =   64
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   510
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCollectMan.frx":6DEC
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   30
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCollectMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnDown As Boolean
Private mblnDrop As Boolean
Private mblnRoot As Boolean
Private mblnEdit As Boolean
Private mnodSelect As Node

Private Const conMenu_����� = 1
Private Const conMenu_���� = 2
Private Const conMenu_�ָ� = 3
Private Const conMenu_���� = 4
Private Const conMenu_�˳� = 5

Private Sub billHLTime_BeforeDeleteRow(ROW As Long, Cancel As Boolean)
    billHLTime.TextMatrix(ROW, 0) = ""
    billHLTime.TextMatrix(ROW, 1) = ""
    billHLTime.TextMatrix(ROW, 2) = ""
    Cancel = True
    mblnEdit = True
End Sub

Private Sub billHLTime_cboClick(ListIndex As Long)
    mblnEdit = True
    billHLTime.TextMatrix(billHLTime.ROW, billHLTime.COL) = billHLTime.CboText
    billHLTime.RowData(billHLTime.ROW) = billHLTime.ItemData(billHLTime.ListIndex)
End Sub

Private Sub billHLTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If billHLTime.TxtVisible Then
        If billHLTime.Text = "" Then billHLTime.Text = " "
    End If
    
    mblnEdit = True
End Sub

Private Sub billTime_BeforeDeleteRow(ROW As Long, Cancel As Boolean)
    With billTime
        .TextMatrix(ROW, 1) = ""
        .TextMatrix(ROW, 2) = ""
        .TextMatrix(ROW, 3) = ""
    End With
    mblnEdit = True
    Cancel = True
End Sub

Private Sub billTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If billTime.TxtVisible Then
        If billTime.Text = "" Then billTime.Text = " "
    End If
    
    mblnEdit = True
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_�����
        Dim ParentNod As Node
        tvwNodes.Nodes.Add , , "KNEW", "NEW", 1, 1
        
        Set ParentNod = tvwNodes.SelectedItem
        Do While Not ParentNod.Child Is Nothing
            Set ParentNod.Child.Parent = tvwNodes.Nodes("KNEW")
        Loop
        tvwNodes.Nodes.Remove ParentNod.Key
        tvwNodes.Nodes("KNEW").Text = ParentNod.Text
        tvwNodes.Nodes("KNEW").Key = ParentNod.Key
        tvwNodes.Nodes(ParentNod.Key).Selected = True
        tvwNodes.SelectedItem.Selected = True
        tvwNodes.SelectedItem.Expanded = True
        
        mblnEdit = True
        mblnRoot = False
    Case conMenu_����
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
        mblnEdit = False
        mblnRoot = False
    Case conMenu_�ָ�
        Call LoadData
    Case conMenu_����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With CollectPage
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_�����
        Control.Enabled = mblnRoot And Me.CollectPage.Selected.Index = 0
    Case conMenu_����
        Control.Enabled = mblnEdit
    Case conMenu_�ָ�
        Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub CollectPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    On Error Resume Next
    Select Case Item.Index
    Case 0
        tvwNodes.SetFocus
    Case 1
        billTime.SetFocus
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Function CheckData() As Boolean
    Dim blnUsual As Boolean
    Dim lngRow As Long, lngCount As Long
    Dim strStart As String, strEnd As String
    'ֻҪ��д�����Ƶ�,��Ӧ����д������ʱ��
    
    lngCount = billTime.Rows - 1
    For lngRow = 1 To lngCount
        If billTime.TextMatrix(lngRow, 1) <> "" Then
            strStart = billTime.TextMatrix(lngRow, 2)
            strEnd = billTime.TextMatrix(lngRow, 3)
            
            If Not CheckTime(1, lngRow, 2, strStart) Then Exit Function
            If Not CheckTime(1, lngRow, 3, strEnd) Then Exit Function
            '���½��Ϸ���ʱ�丳�ر��
            billTime.TextMatrix(lngRow, 2) = strStart
            billTime.TextMatrix(lngRow, 3) = strEnd
        End If
    Next
    
    lngCount = billHLTime.Rows - 1
    For lngRow = 1 To lngCount
        If billHLTime.TextMatrix(lngRow, 1) <> "" Then
            If Not blnUsual Then blnUsual = (billHLTime.TextMatrix(lngRow, 0) = "ͨ��")
            strStart = billHLTime.TextMatrix(lngRow, 2)
            strEnd = billHLTime.TextMatrix(lngRow, 3)
            
            If Not CheckTime(2, lngRow, 2, strStart) Then Exit Function
            If Not CheckTime(2, lngRow, 3, strEnd) Then Exit Function
            '���½��Ϸ���ʱ�丳�ر��
            billHLTime.TextMatrix(lngRow, 2) = strStart
            billHLTime.TextMatrix(lngRow, 3) = strEnd
        End If
    Next
    If Not blnUsual Then
        MsgBox "�����û����¼��С���ͨ��ʱ����Թ���������ʹ�ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    CheckData = True
End Function

Private Function CheckTime(ByVal intBill As Integer, ByVal lngRow As Long, ByVal lngCOL As Long, strTime As String) As Boolean
    Dim strTitle As String
    Dim lngHour As Long, lngMin As Long
    On Error Resume Next
    '���ʱ���ʽ�Ϸ���
    
    If strTime = " " Then
        CheckTime = True
        Exit Function
    End If
    
    strTitle = "��" & lngRow & "�е�" & lngCOL & "�е�"
    Err = 0
    
    '1��ȡСʱ
    If InStr(1, strTime, ":") = 0 Then
        lngHour = strTime
    Else
        lngHour = Split(strTime, ":")(0)
    End If
    If Err <> 0 Then
        MsgBox strTitle & "ʱ���к��зǷ��ַ���" & vbCrLf & _
               "ʱ���ʽΪHH:mm,��05:00", vbInformation, gstrSysName
        CollectPage.Item(intBill).Selected = True
        Exit Function
    End If
    '1.1����С��0����23
    If lngHour < 0 Or lngHour > 23 Then
        MsgBox strTitle & "Сʱ���ܴ���23��С��0��", vbInformation, gstrSysName
        CollectPage.Item(intBill).Selected = True
        Exit Function
    End If
    
    '2��ȡ��
    If InStr(1, strTime, ":") = 0 Then
        lngMin = "00"
    Else
        lngMin = Split(strTime, ":")(1)
    End If
    If Err <> 0 Then
        MsgBox strTitle & "ʱ���к��зǷ��ַ���" & vbCrLf & _
               "ʱ���ʽΪHH:mm,��05:00", vbInformation, gstrSysName
        CollectPage.Item(intBill).Selected = True
        Exit Function
    End If
    '2.1����С��0����23
    If lngMin < 0 Or lngMin > 59 Then
        MsgBox strTitle & "���Ӳ��ܴ���59��С��0��", vbInformation, gstrSysName
        CollectPage.Item(intBill).Selected = True
        Exit Function
    End If
    
    '3������
    If intBill = 2 Then
        If billHLTime.TextMatrix(lngRow, 0) = "" Then
            MsgBox "��¼��С������Ʋ���Ϊ�գ�", vbInformation, gstrSysName
            CollectPage.Item(intBill).Selected = True
            Exit Function
        End If
        If LenB(StrConv(billHLTime.TextMatrix(lngRow, 0), vbFromUnicode)) > 20 Then
            MsgBox "��¼��С������Ʋ��ܴ���10�����ֻ�20���ַ���", vbInformation, gstrSysName
            CollectPage.Item(intBill).Selected = True
            Exit Function
        End If
    End If
    
    '������֯ʱ��
    strTime = String(2 - Len(CStr(lngHour)), "0") & CStr(lngHour) & ":" & String(2 - Len(CStr(lngMin)), "0") & CStr(lngMin)
    CheckTime = True
End Function

Private Function SaveData() As Boolean
    Dim objNode As Node
    Dim lngStart As Long, lngCount As Long
    Dim strParent As String, strItems As String, strTimes As String, strHLTimes As String
    On Error GoTo Errhand
    
    '�Ȳ���������Ŀ
    For Each objNode In tvwNodes.Nodes
        If objNode.Parent Is Nothing Then
            strParent = ""
        Else
            strParent = Mid(objNode.Parent.Key, 2)
        End If
        strItems = strItems & "|" & Mid(objNode.Key, 2) & "," & strParent
    Next
    strItems = Mid(strItems, 2)
    
    '�������µ�����
    lngCount = billTime.Rows - 1
    For lngStart = 1 To lngCount
        strTimes = strTimes & "|" & billTime.TextMatrix(lngStart, 1) & "," & billTime.TextMatrix(lngStart, 2) & "," & billTime.TextMatrix(lngStart, 3) & "," & lngStart
    Next
    strTimes = Mid(strTimes, 2)
    '������¼��С��
    lngCount = billHLTime.Rows - 1
    For lngStart = 1 To lngCount
        strHLTimes = strHLTimes & "|" & billHLTime.RowData(lngStart) & "," & billHLTime.TextMatrix(lngStart, 1) & "," & billHLTime.TextMatrix(lngStart, 2) & "," & billHLTime.TextMatrix(lngStart, 3) & "," & lngStart
    Next
    strHLTimes = Mid(strHLTimes, 2)
    Call zldatabase.ExecuteProcedure("ZL_�����������_UPDATE('" & strItems & "','" & strTimes & "','" & strHLTimes & "')", "��������")
    
    SaveData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    Dim strSQL As String
    Dim strIDs As String            '�ѽ��й���Ļ�����Ŀ
    Dim introw As Integer, lng����ID As Long, int��� As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    On Error GoTo Errhand
    
    mblnRoot = False
    mblnEdit = False
    tvwNodes.Nodes.Clear
    
    '��ȡ���л�����Ŀ
    strSQL = "" & _
            " select A.��Ŀ���,A.��Ŀ����" & _
            " from �����¼��Ŀ A" & _
            " Where A.��Ŀ��ʾ=4 " & _
            " Order By A.��Ŀ���"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ���л�����Ŀ")
    If rsTemp.RecordCount = 0 Then
        MsgBox "�������û�����Ŀ����ʹ�ñ�ģ�飡", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '��ȡ�����õ����¼����
    strSQL = " Select ���,�����" & _
             " From ���������Ŀ" & _
             " Start With ����� Is NULL Connect By Prior ���=�����"
    Set rsItem = zldatabase.OpenSQLRecord(strSQL, "��ȡ�����õ����¼����")
    '��ӽ��
    If rsItem.RecordCount <> 0 Then
        Do While Not rsItem.EOF
            rsTemp.Filter = "��Ŀ���=" & rsItem!���
            If rsTemp.RecordCount <> 0 Then
                strIDs = strIDs & "," & rsItem!���
                If IsNull(rsItem!�����) Then
                    tvwNodes.Nodes.Add , , "K" & rsItem!���, rsTemp!��Ŀ����, 1, 1
                Else
                    tvwNodes.Nodes.Add "K" & rsItem!�����, 4, "K" & rsItem!���, rsTemp!��Ŀ����, 1, 1
                End If
            End If
            rsItem.MoveNext
        Loop
        
        rsTemp.Filter = 0
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If InStr(1, "," & strIDs & ",", "," & rsTemp!��Ŀ��� & ",") = 0 Then
                tvwNodes.Nodes.Add , , "K" & rsTemp!��Ŀ���, rsTemp!��Ŀ����, 1, 1
            End If
            rsTemp.MoveNext
        Loop
    Else
        Do While Not rsTemp.EOF
            tvwNodes.Nodes.Add , , "K" & rsTemp!��Ŀ���, rsTemp!��Ŀ����, 1, 1
            rsTemp.MoveNext
        Loop
    End If
    
    '��ʼ���༭�ؼ�
    With billTime
        .ClearBill
        .Rows = 4
        .Cols = 4
        .TextMatrix(0, 0) = "���"
        .TextMatrix(0, 1) = "ʱ������"
        .TextMatrix(0, 2) = "��ʼʱ��"
        .TextMatrix(0, 3) = "����ʱ��"
        .ColData(0) = 5
        .ColData(1) = 4
        .ColData(2) = 4
        .ColData(3) = 4
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = False
        .Active = True
        
        .TextMatrix(1, 0) = "�������"
        .TextMatrix(2, 0) = "ҹ�����"
        .TextMatrix(3, 0) = "ȫ�����"
    End With
    With billHLTime
        .ClearBill
        .Rows = 2
        .Cols = 4
        .TextMatrix(0, 0) = "���"
        .TextMatrix(0, 1) = "ʱ������"
        .TextMatrix(0, 2) = "��ʼʱ��"
        .TextMatrix(0, 3) = "����ʱ��"
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
        .ColData(3) = 4
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = True
        .Active = True
    End With
    
    '��ȡ����Ԫ
    strSQL = " Select A.ID,A.���� From ���ű� A,��������˵�� B Where A.ID=B.����ID And B.������� IN (2,3) And B.��������='����'"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ����Ԫ")
    billHLTime.AddItem "ͨ��"
    Do While Not rsTemp.EOF
        billHLTime.AddItem rsTemp!����
        billHLTime.ItemData(billHLTime.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    
    '��ȡ���µ���������
    strSQL = " Select A.����ID,DECODE(B.����,NULL,'ͨ��',B.����) AS ����,A.����,A.���,A.����,A.��ʼ,A.���� From �������ʱ�� A,���ű� B " & _
             " Where A.����ID=B.ID(+) " & _
             " Order by NVL(����ID,0),����,���"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ������")
    With rsTemp
        .Filter = "����=1"
        lng����ID = 0: int��� = 1: introw = 1
        Do While Not .EOF
            If NVL(!����ID, 0) <> lng����ID Or int��� <> !��� Then
                introw = introw + 1
'                billTime.Rows = billTime.Rows + 1
                lng����ID = NVL(!����ID, 0)
                int��� = !���
            End If
            billTime.TextMatrix(introw, 1) = !����
            billTime.TextMatrix(introw, 2) = !��ʼ
            billTime.TextMatrix(introw, 3) = !����
            .MoveNext
        Loop
        
        .Filter = "����=2"
        lng����ID = 0: int��� = 1: introw = 1
        Do While Not .EOF
            If NVL(!����ID, 0) <> lng����ID Or int��� <> !��� Then
                introw = introw + 1
                billHLTime.Rows = billHLTime.Rows + 1
                lng����ID = NVL(!����ID, 0)
                int��� = !���
            End If
            billHLTime.RowData(introw) = NVL(!����ID, 0)
            billHLTime.TextMatrix(introw, 0) = !����
            billHLTime.TextMatrix(introw, 1) = !����
            billHLTime.TextMatrix(introw, 2) = !��ʼ
            billHLTime.TextMatrix(introw, 3) = !����
            .MoveNext
        Loop
        .Filter = 0
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim lngHandel As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgPublic.Icons
    
    '����������
    '-----------------------------------------------------
    cbsMain.DeleteAll
    Set objBar = cbsMain.Add("������", xtpBarTop)      '����
    objBar.EnableDocking xtpFlagStretched
    objBar.Closeable = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_�����, "�����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "����Ϊ�����"
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�ָ�, "�ָ�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ȡ������"
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�˳�, "�˳�"): objControl.Style = xtpButtonIconAndCaption
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyR, conMenu_�����           '�����
        .Add FCONTROL, vbKeyS, conMenu_����             '����
    End With
    
    '�ٴ����ҳ�ؼ�
    With CollectPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "��Ŀ����", picNodes.hwnd, 0).Tag = "��Ŀ����"
        .InsertItem(1, "���µ�����", picTime.hwnd, 0).Tag = "���µ�����"
        .InsertItem(2, "��¼��С��", picHLTime.hwnd, 0).Tag = "��¼������"
    End With
End Sub

Private Sub picHLTime_Resize()
    On Error Resume Next
    
    billHLTime.Width = picHLTime.Width
    billHLTime.Height = picHLTime.Height

End Sub

Private Sub tvwNodes_DragDrop(Source As Control, x As Single, y As Single)
    Dim nod As Node
    Set nod = tvwNodes.SelectedItem
    
    Set tvwNodes.DragIcon = Nothing
    tvwNodes.Drag 0
    mblnDrop = False
    mblnDown = False
    
    Call MoveNodes(nod, mnodSelect)
End Sub

Private Sub tvwNodes_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Dim nod As Node
    
    If Not mblnDrop Then Exit Sub
    Set nod = tvwNodes.HitTest(x, y)
    If nod Is Nothing Then Exit Sub
    Set tvwNodes.SelectedItem = nod
End Sub

Private Sub tvwNodes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnDown = (Button = 1)
End Sub

Private Sub tvwNodes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mblnDown Then Exit Sub
    
    mblnDrop = True
    Set tvwNodes.DragIcon = imgIcon.ListImages(1).Picture
    tvwNodes.Drag 1
    Set mnodSelect = tvwNodes.SelectedItem
End Sub

Private Sub tvwNodes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnDown = False
End Sub

Private Sub MoveNodes(ByVal ParentNod As Node, ByVal ChildNod As Node)
    On Error GoTo Errhand
    
    '��ָ����㼰�ӽ��ȫ���Ƶ��������
    If ParentNod Is Nothing Then Exit Sub
    If ParentNod.Key = mnodSelect.Key Then Exit Sub
    
    Set mnodSelect.Parent = ParentNod
    ParentNod.Expanded = True
    
    mblnEdit = True
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub tvwNodes_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnRoot = False
    If mblnDrop Then Exit Sub
    mblnRoot = Not (Node.Parent Is Nothing)
End Sub

Private Sub picNodes_Resize()
    On Error Resume Next
    
    tvwNodes.Width = picNodes.Width
    tvwNodes.Height = picNodes.Height
End Sub

Private Sub picTime_Resize()
    On Error Resume Next
    
    billTime.Width = picTime.Width
    billTime.Height = picTime.Height
End Sub
