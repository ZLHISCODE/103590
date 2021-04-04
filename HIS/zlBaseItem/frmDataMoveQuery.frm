VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDataMoveQuery 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   Icon            =   "frmDataMoveQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8745
   Begin MSComDlg.CommonDialog cdgSave 
      Left            =   810
      Top             =   1335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   30
      TabIndex        =   8
      Top             =   5580
      Width           =   8685
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7185
         TabIndex        =   5
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   3900
         TabIndex        =   3
         ToolTipText     =   "���ң�F3"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2535
         TabIndex        =   2
         Top             =   150
         Width           =   1320
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   1230
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "�����&Excel"
         Height          =   350
         Left            =   5895
         TabIndex        =   4
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2145
         TabIndex        =   9
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.Frame fraNote 
      Height          =   645
      Left            =   15
      TabIndex        =   6
      Top             =   -45
      Width           =   8700
      Begin VB.Label lblNote 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   360
         Left            =   195
         TabIndex        =   7
         Top             =   180
         Width           =   8400
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsData 
      Height          =   4935
      Left            =   15
      TabIndex        =   0
      Top             =   630
      Width           =   8700
      _cx             =   15346
      _cy             =   8705
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmDataMoveQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private mrsData As ADODB.Recordset
Private mintType As Integer
Private mdatBegin As Date
Private mdatEnd As Date
Private mstrTitle As String
Private mstrNote As String

Private mblnExcel As Boolean
Private mlngBegin As Long

Public Sub ShowMe(ByVal intType As Integer, ByVal datBegin As Date, ByVal datEnd As Date, ByVal strTitle As String, ByVal strNote As String, FrmParent As Object)
    mintType = intType
    mdatBegin = datBegin
    mdatEnd = datEnd
    mstrTitle = strTitle
    mstrNote = strNote
    
    On Error Resume Next
    Me.Show 1, FrmParent
End Sub

Private Sub cboFind_Click()
    mlngBegin = 0
    txtFind.Text = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function HaveExcel() As Boolean
'���ܣ��ж�ϵͳ�Ƿ�װ��Excel
'˵����ͬʱ��ʼ��Excel����
    Dim objExcel As Object
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    HaveExcel = Err.Number = 0
    Set objExcel = Nothing
End Function

Private Sub cmdExcel_Click()
    Dim strFile As String
    Dim lngBack As Long, lngFore As Long
    
    strFile = Me.Caption & "(" & Format(mdatBegin, "yyyyMMdd") & "-" & Format(mdatEnd, "yyyyMMdd") & ").xls"
    On Error GoTo errH
    cdgSave.DialogTitle = "����Excel���"
    cdgSave.Filter = "Microsoft Office Excel�ļ�(*.xls)|*.xls"
    cdgSave.flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
    cdgSave.FileName = strFile
    cdgSave.CancelError = True
    cdgSave.ShowSave
    On Error GoTo 0
    strFile = cdgSave.FileName
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName, "ExcelPath", Left(strFile, Len(strFile) - Len(cdgSave.FileTitle))
    
    vsData.redraw = flexRDNone
    lngBack = vsData.BackColorSel
    lngFore = vsData.ForeColorSel
    vsData.BackColorSel = vsData.BackColor
    vsData.ForeColorSel = vsData.ForeColor
    vsData.SaveGrid strFile, flexFileExcel, flexXLSaveFixedCells
    vsData.BackColorSel = lngBack
    vsData.ForeColorSel = lngFore
    vsData.redraw = flexRDDirect
    
    If mblnExcel Then
        ShellExecute Me.hwnd, "open", strFile, "", "", SW_SHOWNORMAL
    Else
        MsgBox "�Ѿ�������ļ�""" & strFile & """�С�", vbInformation, gstrSysName
    End If
errH:
End Sub

Private Sub cmdFind_Click()
    Dim lngRow As Long, blnFull As Boolean
    
    If txtFind.Text = "" Then
        MsgBox "������Ҫ���ҵ����ݡ�", vbInformation, gstrSysName
        txtFind.SetFocus: Exit Sub
    End If
    
    lngRow = vsData.FindRow(txtFind.Text, mlngBegin + 1, cboFind.ItemData(cboFind.ListIndex), False, InStr("����,���ݺ�,�Һŵ�", cboFind.Text) = 0)
    If lngRow <> -1 Then
        mlngBegin = lngRow
        vsData.Row = lngRow
        Call vsData.ShowCell(vsData.Row, 0)
    Else
        mlngBegin = 0
        MsgBox "�Ѿ��ҵ����β����δ���ַ����������С��´ν����´ӱ�ͷ��ʼ���ҡ�", vbInformation, gstrSysName
    End If
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFind.Enabled Then cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName, mintType)
    Me.Caption = mstrTitle
    
    lblNote.Caption = "��ѯʱ�䣺" & Format(mdatBegin, "yyyy-MM-dd") & " �� " & Format(mdatEnd, "yyyy-MM-dd")
    lblNote.Caption = lblNote.Caption & vbCrLf & mstrNote
    
    If Not LoadData Then Unload Me: Exit Sub
    mblnExcel = HaveExcel
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraNote.Left = 0
    fraNote.Top = -45
    fraNote.Width = Me.ScaleWidth
    lblNote.Width = fraNote.Width - lblNote.Left * 2
    
    vsData.Left = 0
    vsData.Top = fraNote.Top + fraNote.Height
    vsData.Width = Me.ScaleWidth
    vsData.Height = Me.ScaleHeight - vsData.Top - fraFunc.Height
    
    fraFunc.Left = 0
    fraFunc.Top = vsData.Top + vsData.Height
    fraFunc.Width = Me.ScaleWidth
    
    If fraFunc.Width - cmdCancel.Width - 500 >= 6500 Then
        cmdCancel.Left = fraFunc.Width - cmdCancel.Width - 500
    Else
        cmdCancel.Left = 6500
    End If
    cmdExcel.Left = cmdCancel.Left - cmdExcel.Width
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsData.State = 1 Then mrsData.Close
    Set mrsData = Nothing
    
    Call SaveWinState(Me, App.ProductName, mintType)
End Sub

Private Function LoadData() As Boolean
    Dim strBegin As String, strEnd As String
    Dim i As Long
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    Set mrsData = New ADODB.Recordset
    
'    strBegin = "To_Date('" & Format(mdatBegin, "yyyy-MM-dd") & "','YYYY-MM-DD')"
'    strEnd = "To_Date('" & Format(mdatEnd, "yyyy-MM-dd") & "','YYYY-MM-DD')+1"
    strBegin = Format(mdatBegin, "yyyy-MM-dd")
    strEnd = Format(DateAdd("d", 1, mdatEnd), "yyyy-MM-dd")
    Select Case mintType
    Case 0
        '�շѣ��Һŵ���
        gstrSQL = _
        " Select Distinct  '�����շ�' As ��������, Decode(d.�շ����,'4','����δ�ڴ�֮ǰ����','ҩƷδ�ڴ�֮ǰ��ҩ') as �޷�ת��ԭ��," & _
        "       d.No As ���ݺ�,d.��ʶ��,d.����,d.�Ա�,d.����,To_Char(d.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��" & _
        " From ҩƷ�շ���¼ l," & _
        "     ( Select d.Id,d.����id,d.NO,d.��ʶ��,d.����,d.�Ա�,d.����,d.�Ǽ�ʱ��,d.�շ����" & _
        "       From ������ü�¼ d " & _
        "       Where d.�Ǽ�ʱ��>=[1] And d.�Ǽ�ʱ��<[2] And d.����ID Is Not Null" & _
        "             And d.��¼���� = 1 And d.�շ���� In ('4', '5', '6', '7')) d" & _
        " Where l.No = d.No And l.����id = d.Id And Nvl(��ҩ��ʽ, 0) <> -1" & _
        "       And (l.������� >=[2] Or l.������� Is Null) And l.���� In (8, 24)"
                          
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(c.��¼����,1,'�����շ�',4,'����Һ�') As ��������,'����ʱʹ�õ�Ԥ����δ����' as �޷�ת��ԭ��," & _
        "       c.No As ���ݺ�,c.��ʶ��,c.����,c.�Ա�,c.����,To_Char(c.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��" & _
        " From ������ü�¼ c,����Ԥ����¼ d," & _
        "      (    Select d.No " & _
        "           From ����Ԥ����¼ d," & _
        "               (Select ����id From ������ü�¼  Where �Ǽ�ʱ��>=[1] And �Ǽ�ʱ��<[2] And ��¼���� In (1, 4) And Nvl(���ʷ���,0)=0 ) l" & _
        "           Where d.����id = l.����id And d.��¼���� In (1, 11)" & _
        "           Group By d.No" & _
        "           Having d.No Is Not Null And Sum(d.���) - Sum(d.��Ԥ��) <> 0) n" & _
        " Where d.No = n.No And d.��¼���� In (1, 11)" & _
        " And c.����ID=d.����ID And c.��¼���� IN(1, 4) And Nvl(c.���ʷ���,0)=0" & _
        " Order By ��������,���ݺ� Desc"
    Case 1
        '���ʵ���
        gstrSQL = _
        " Select Distinct Decode(d.�����־,2,'סԺ����','�������') As ��������, Decode(d.�շ����,'4','����δ�ڴ�֮ǰ����','ҩƷδ�ڴ�֮ǰ��ҩ') as �޷�ת��ԭ��," & _
        "       d.No As ���ݺ�,d.��ʶ��,d.����,d.�Ա�,d.����,To_Char(d.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From ҩƷ�շ���¼ l," & _
        "     ( Select d.Id,d.����id,d.NO,d.��ʶ��,d.����,d.�Ա�,d.����,d.�Ǽ�ʱ��,d.�շ����,d.�����־" & _
        "       From סԺ���ü�¼ d" & _
        "       Where d.�Ǽ�ʱ��>=[1] And �Ǽ�ʱ��<[2] And d.����ID Is Not Null" & _
        "             And d.��¼����=2 And d.�շ���� In ('4', '5', '6', '7') " & _
        "       Union ALL " & _
        "       Select d.Id,d.����id,d.NO,d.��ʶ��,d.����,d.�Ա�,d.����,d.�Ǽ�ʱ��,d.�շ����,d.�����־" & _
        "       From ������ü�¼ d" & _
        "       Where d.�Ǽ�ʱ��>=[1] And �Ǽ�ʱ��<[2] And d.����ID Is Not Null" & _
        "             And d.��¼����=2 And d.�շ���� In ('4', '5', '6', '7') " & _
        "       ) d" & _
        " Where l.No = d.No And l.����id = d.Id And Nvl(��ҩ��ʽ, 0) <> -1" & _
        "       And (l.������� >= [2] Or l.������� Is Null)  And l.���� In (9, 10, 25, 26)"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(n.��¼����,2,Decode(d.�����־,2,'סԺ����','�������'),3,'�Զ�����',5,'���￨����') As ��������,'ͬʱ����ĵ�����δ�������' as �޷�ת��ԭ��," & _
        "       d.No As ���ݺ�,d.��ʶ��,d.����,d.�Ա�,d.����,To_Char(d.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From ������ü�¼ d," & _
        "     (   Select d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) As ��¼����" & _
        "         From ������ü�¼ d, ���˽��ʼ�¼ l" & _
        "         Where d.����id = l.Id And l.�շ�ʱ��>=[1] And l.�շ�ʱ��<[2] " & _
        "               And d.��¼���� In (2, 12, 3, 13, 5, 15) And d.���ʷ��� = 1" & _
        "         Group By d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����)" & _
        "         Having d.No Is Not Null And d.��� Is Not Null And Nvl(Sum(d.ʵ�ս��),0) - Nvl(Sum(d.���ʽ��),0) <> 0 " & _
        "       ) n" & _
        " Where d.No = n.No And d.��� = n.��� And Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) = n.��¼����"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(n.��¼����,2,Decode(d.�����־,2,'סԺ����','�������'),3,'�Զ�����',5,'���￨����') As ��������,'ͬʱ����ĵ�����δ�������' as �޷�ת��ԭ��," & _
        "       d.No As ���ݺ�,d.��ʶ��,d.����,d.�Ա�,d.����,To_Char(d.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From סԺ���ü�¼ d," & _
        "     (   Select d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) As ��¼����" & _
        "         From סԺ���ü�¼ d, ���˽��ʼ�¼ l" & _
        "         Where d.����id = l.Id And l.�շ�ʱ��>=[1] And l.�շ�ʱ��<[2] " & _
        "               And d.��¼���� In (2, 12, 3, 13, 5, 15) And d.���ʷ��� = 1" & _
        "           Group By d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����)" & _
        "           Having d.No Is Not Null And d.��� Is Not Null And Nvl(Sum(d.ʵ�ս��),0) - Nvl(Sum(d.���ʽ��),0) <> 0 " & _
        "       ) n" & _
        " Where d.No = n.No And d.��� = n.��� And Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) = n.��¼����"
        
        '���������
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(Mod(c.��¼����,10),2,Decode(c.�����־,2,'סԺ����','�������'),3,'�Զ�����',5,'���￨����') As ��������, '����ʱʹ�õ�Ԥ����δ����' as �޷�ת��ԭ��," & _
        "       c.No As ���ݺ�,c.��ʶ��,c.����,c.�Ա�,c.����,To_Char(c.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From סԺ���ü�¼ c,����Ԥ����¼ d," & _
        "     ( Select d.No" & _
        "       From ����Ԥ����¼ d," & _
        "           (   Select Id As ����id From ���˽��ʼ�¼ Where �շ�ʱ��>=[1] And �շ�ʱ��<[2] ) l" & _
        "       Where d.����id = l.����id And d.��¼���� In (1, 11)" & _
        "       Group By d.No" & _
        "       Having d.No Is Not Null And Sum(d.���) - Sum(d.��Ԥ��) <> 0 " & _
        "      ) n" & _
        " Where d.No = n.No And d.��¼���� In (1, 11)" & _
        "       And c.����ID=d.����ID And c.��¼���� IN(2, 12, 3, 13, 5, 15) And c.���ʷ���=1"
        
        gstrSQL = gstrSQL & " Union ALL " & _
        " Select Distinct Decode(Mod(c.��¼����,10),2,Decode(c.�����־,2,'סԺ����','�������'),3,'�Զ�����',5,'���￨����') As ��������, '����ʱʹ�õ�Ԥ����δ����' as �޷�ת��ԭ��," & _
        "       c.No As ���ݺ�,c.��ʶ��,c.����,c.�Ա�,c.����,To_Char(c.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From ������ü�¼ c,����Ԥ����¼ d," & _
        "     ( Select d.No" & _
        "       From ����Ԥ����¼ d," & _
        "           (   Select Id As ����id From ���˽��ʼ�¼ Where �շ�ʱ��>=[1] And �շ�ʱ��<[2] ) l" & _
        "       Where d.����id = l.����id And d.��¼���� In (1, 11)" & _
        "       Group By d.No" & _
        "       Having d.No Is Not Null And Sum(d.���) - Sum(d.��Ԥ��) <> 0 " & _
        "      ) n" & _
        " Where d.No = n.No And d.��¼���� In (1, 11)" & _
        "       And c.����ID=d.����ID And c.��¼���� IN(2, 12, 3, 13, 5, 15) And c.���ʷ���=1" & _
        " Order By ��������,���ݺ� Desc"
        
        
    Case 2
        '���ﲡ��
        gstrSQL = _
        " Select Decode(Count(d.No),0,Null,'���˹Һŷ���δת��') ||Decode(Count(e.�Һ�id),0,Null,CHR(13)||CHR(10)||'����δת����ҽ������')  ||Decode(Count(a.�Һ�id),0,Null,CHR(13)||CHR(10)||'����δ�ڴ�֮ǰ���͵�ҽ��') as �޷�ת��ԭ��," & _
        "       r.No As �Һŵ�,r.�����,r.����,r.�Ա�,r.����,x.���� As �������,To_Char(r.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & _
        " From ���ű� x,���˹Һż�¼ r," & _
        "   (   Select No From ������ü�¼ Where �Ǽ�ʱ��>=[1] And �Ǽ�ʱ��<[2] And ��¼���� = 4) d," & _
        "   (   Select r.Id As �Һ�id From ����ҽ����¼ a, ���˹Һż�¼ r Where a.�Һŵ� = r.No And r.�Ǽ�ʱ��>=[1] And r.�Ǽ�ʱ��<[2] Group By r.Id Having Max(a.ͣ��ʱ��)>=[2]) a," & _
        "   (   Select a.�Һ�id From ������ü�¼ e, (Select a.Id, r.Id As �Һ�id  From ����ҽ����¼ a, ���˹Һż�¼ r  Where a.�Һŵ� = r.No And r.�Ǽ�ʱ��>=[1] And r.�Ǽ�ʱ��<[2]) a" & _
        "       Where e.ҽ����� = a.Id) e" & _
        " Where r.No = d.No(+) And r.Id = a.�Һ�id(+) And r.Id = e.�Һ�id(+)" & _
        "       And r.ִ��״̬<>2 And r.�Ǽ�ʱ��>=[1] And r.�Ǽ�ʱ��<[2] And x.Id=r.ִ�в���ID" & _
        " Group By r.No,r.�����,r.����,r.�Ա�,r.����,x.����,r.�Ǽ�ʱ��" & _
        " Having Count(d.No) > 0 Or Count(a.�Һ�id) > 0 Or Count(e.�Һ�id) > 0" & _
        " Order By ����ʱ�� Desc,�����"
    Case 3
        'סԺ����
        gstrSQL = _
            " Select '���˴���δת������' as �޷�ת��ԭ��,i.סԺ��,i.����,i.�Ա�,i.����,p.��ҳid As סԺ����,d.���� As סԺ����," & _
            "        To_Char(p.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժʱ��,To_Char(p.��Ժ����,'YYYY-MM-DD HH24:MI') as ��Ժʱ��" & _
            " From ���ű� d,������Ϣ i,������ҳ p" & _
            " Where p.��Ժ����>=[1] And p.��Ժ����<[2] And Nvl(p.����ת��, 0) <> 1" & _
                " And i.����ID=p.����ID And p.��Ժ����ID=d.ID" & _
                " And Exists (Select 1 From סԺ���ü�¼ Where ����id = p.����id And ��ҳid = p.��ҳid)" & _
            " Order BY ��Ժ���� Desc,סԺ��"

    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Case 4
        '�����������
        gstrSQL = _
            "Select '�ܼ���Ա�ķ���δת��' As �޷�ת��ԭ��, c.����, c.�Ա�, c.����, c.�����, c.������, b.������, a.����ʱ�� As ���ʱ��, d.���� As ������" & vbNewLine & _
            "From ���������Ա A, ��������¼ B, ������Ϣ C, ���ű� D" & vbNewLine & _
            "Where a.����ʱ�� > [1] And a.����ʱ�� < [2] And a.����id = b.Id And a.����id = c.����id And b.��첿��id = d.Id And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From ������ü�¼ X, ���������� Y" & vbNewLine & _
            "       Where y.����id = a.����id And y.����id = a.����id And x.�����־ = 4 And x.No = y.No And x.��¼���� = y.��¼����)" & vbNewLine & _
            "Order By a.����ʱ�� Desc, c.����, c.����id, b.������"
            
    End Select
    
    Set mrsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(strBegin), CDate(strEnd))
    
    With Me.vsData
        Set .DataSource = mrsData
        Err.Clear
        Call RestoreFlexState(vsData, App.ProductName & "\" & Me.Name & mintType)
        If .Rows = .FixedRows Then
            .Rows = .FixedRows + 1
            cmdExcel.Enabled = False
            cmdFind.Enabled = False
        End If
        If mintType = 2 Then
            .WordWrap = True
            .ColWidth(0) = 1800
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 0
        End If
        
        cboFind.Clear
        For i = 0 To .Cols - 1
            If InStr("���ݺ�,�Һŵ�,����,סԺ��,�����,��ʶ��", .TextMatrix(0, i)) > 0 Then
                cboFind.AddItem .TextMatrix(0, i)
                cboFind.ItemData(cboFind.NewIndex) = i
            End If
            .ColAlignment(i) = 1
        Next
        cboFind.ListIndex = 0
        
        .RowHeight(0) = 250
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
        .Row = 1
    End With
    
    Screen.MousePointer = 0
    LoadData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lblNote_Change()
    mlngBegin = 0
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdFind.Enabled Then Call cmdFind_Click
    ElseIf InStr("���ݺ�,�Һŵ�", cboFind.Text) > 0 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
