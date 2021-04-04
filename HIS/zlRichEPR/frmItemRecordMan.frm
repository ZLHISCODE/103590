VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmItemRecordMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼Ƶ������"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   Icon            =   "frmItemRecordMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin ZL9BillEdit.BillEdit billTime 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5054
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
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   510
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItemRecordMan.frx":6852
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
Attribute VB_Name = "frmItemRecordMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnEdit As Boolean

Private Const conMenu_���� = 2
Private Const conMenu_�ָ� = 3
Private Const conMenu_���� = 4
Private Const conMenu_�˳� = 5

Private Sub billTime_BeforeDeleteRow(ROW As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub billTime_EnterCell(ROW As Long, COL As Long)
    If COL < 2 Then Exit Sub
    If COL < Val(billTime.TextMatrix(ROW, 0)) + 3 Then
        billTime.ColData(COL) = 4
    Else
        billTime.ColData(COL) = 0
    End If
End Sub

Private Sub billTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If billTime.TxtVisible Then
        If billTime.Text = "" Then billTime.Text = " "
    End If
    
    mblnEdit = True
End Sub

Private Sub billTime_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_����
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
        mblnEdit = False
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
    With billTime
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_����
        Control.Enabled = mblnEdit
    Case conMenu_�ָ�
        Control.Enabled = mblnEdit
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
    Dim lngRow As Long, lngCount As Long
    'ֻҪ��д�����Ƶ�,��Ӧ����д������ʱ��
    
    lngCount = billTime.Rows - 1
    For lngRow = 1 To lngCount
        If billTime.TextMatrix(lngRow, 1) <> "" Then
            If Not CheckTime(lngRow, 2) Then Exit Function
            If Not CheckTime(lngRow, 3) Then Exit Function
            If Not CheckTime(lngRow, 4) Then Exit Function
            If Not CheckTime(lngRow, 5) Then Exit Function
            If Not CheckTime(lngRow, 6) Then Exit Function
            If Not CheckTime(lngRow, 7) Then Exit Function
            If Not CheckTime(lngRow, 8) Then Exit Function
        End If
    Next
    CheckData = True
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lngCOL As Long) As Boolean
    Dim strTitle As String
    Dim strTime As String
    Dim lngHour As Long, lngMin As Long
    On Error Resume Next
    '���ʱ���ʽ�Ϸ���
    
    strTime = billTime.TextMatrix(lngRow, lngCOL)
    If strTime = "" Then
        If lngCOL <= Val(billTime.TextMatrix(lngRow, 0)) + 2 Then
            MsgBox "��" & lngRow & "���в�������δ¼������ʱ�㣡", vbInformation, gstrSysName
            CheckTime = False
            Exit Function
        Else
            CheckTime = True
            Exit Function
        End If
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
        Exit Function
    End If
    '1.1����С��0����23
    If lngHour < 0 Or lngHour > 23 Then
        MsgBox strTitle & "Сʱ���ܴ���23��С��0��", vbInformation, gstrSysName
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
        Exit Function
    End If
    '2.1����С��0����23
    If lngMin < 0 Or lngMin > 59 Then
        MsgBox strTitle & "���Ӳ��ܴ���59��С��0��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������֯ʱ��
    strTime = String(2 - Len(CStr(lngHour)), "0") & CStr(lngHour) & ":" & String(2 - Len(CStr(lngMin)), "0") & CStr(lngMin)
    billTime.TextMatrix(lngRow, lngCOL) = strTime
    
    CheckTime = True
End Function

Private Function SaveData() As Boolean
    Dim strIn As String
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim strBegin As String, strEnd As String
    Dim lngStart As Long, lngCount As Long, lngCOL As Long
    On Error GoTo errHand
    ReDim Preserve strSQL(1 To 1)
    
    gstrSQL = "ZL_������ĿƵ��_DELETE"
    strSQL(ReDimArray(strSQL)) = gstrSQL
    
    lngCount = billTime.Rows - 1
    For lngStart = 1 To lngCount
        strBegin = ""
        strEnd = ""
        gstrSQL = "ZL_������ĿƵ��_UPDATE("
        For lngCOL = 0 To Val(billTime.TextMatrix(lngStart, 0)) - 1
            If strBegin = "" Then
                strBegin = billTime.TextMatrix(lngStart, 2 + lngCOL)
                strEnd = billTime.TextMatrix(lngStart, 3 + lngCOL)
            Else
                strBegin = Format(DateAdd("s", 60, "2010-01-01 " & strEnd & ":00"), "HH:mm")
                strEnd = billTime.TextMatrix(lngStart, 3 + lngCOL)
            End If
            
            strIn = Val(billTime.TextMatrix(lngStart, 0)) & "," & lngCOL + 1 & ",'" & strBegin & "','" & strEnd & "'," & Val(billTime.TextMatrix(lngStart, 1)) & ")"
            strIn = gstrSQL & strIn
            strSQL(ReDimArray(strSQL)) = strIn
        Next
    Next
    
    'ѭ��ִ��SQL��������
    gcnOracle.BeginTrans
    blnTrans = True
    lngCount = UBound(strSQL)
    For lngStart = 1 To lngCount
        If strSQL(lngStart) <> "" Then
            Debug.Print strSQL(lngStart)
            Call zlDatabase.ExecuteProcedure(strSQL(lngStart), "���滤����ĿƵ��")
        End If
    Next
    SaveData = True
    gcnOracle.CommitTrans
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    Dim intDo As Integer, intRow As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    mblnEdit = False
    '��ʼ���༭�ؼ�
    With billTime
        .ClearBill
        .Rows = 6
        .Cols = 9
        .TextMatrix(0, 0) = "Ƶ��"
        .TextMatrix(0, 1) = "ȡ������"
        .TextMatrix(0, 2) = "��ʼʱ��"
        .TextMatrix(0, 3) = "�ֶ�1"
        .TextMatrix(0, 4) = "�ֶ�2"
        .TextMatrix(0, 5) = "�ֶ�3"
        .TextMatrix(0, 6) = "�ֶ�4"
        .TextMatrix(0, 7) = "�ֶ�5"
        .TextMatrix(0, 8) = "����ʱ��"
        .ColData(0) = 5
        .ColData(1) = 3
        .ColData(2) = 4
        .ColData(3) = 4
        .ColData(4) = 4
        .ColData(5) = 4
        .ColData(6) = 4
        .ColData(7) = 4
        .ColData(8) = 4
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 900
        .ColWidth(3) = 600
        .ColWidth(4) = 600
        .ColWidth(5) = 600
        .ColWidth(6) = 600
        .ColWidth(7) = 600
        .ColWidth(8) = 900
        .PrimaryCol = 1
        .LocateCol = 1
        .ColAlignment(1) = 1
        .AllowAddRow = False
        .Active = True
        
        .AddItem "1-ȡ��һ������"
        .AddItem "2-ȡ�м�ʱ�������"
        .AddItem "3-ȡ���һ������"
        .cboStyle = DropOlnyDown
        .ListIndex = 0
        
        .TextMatrix(1, 0) = "1"
        .TextMatrix(2, 0) = "2"
        .TextMatrix(3, 0) = "3"
        .TextMatrix(4, 0) = "4"
        .TextMatrix(5, 0) = "6"
    End With
    
    '��ȡ����ʱ������
    strSQL = " Select Ƶ��,���,DECODE(���,1,'1-ȡ��һ������',2,'2-ȡ�����е������','3-ȡ���һ������') AS ���,��ʼ,���� " & vbNewLine & _
             " From ������ĿƵ�� " & vbNewLine & _
             " Order by Ƶ��,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ������")
    intRow = 0
    With rsTemp
        Do While Not .EOF
            If intRow <> IIf(!Ƶ�� = 6, 5, !Ƶ��) Then
                intDo = 1
                intRow = intRow + 1
                billTime.TextMatrix(intRow, 1) = !���
                billTime.TextMatrix(intRow, 2) = NVL(!��ʼ)
            End If
            
            billTime.TextMatrix(intRow, 2 + intDo) = NVL(!����)
            
            intDo = intDo + 1
            .MoveNext
        Loop
    End With

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
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�ָ�, "�ָ�"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ȡ������"
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�˳�, "�˳�"): objControl.Style = xtpButtonIconAndCaption
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_����             '����
    End With
End Sub
