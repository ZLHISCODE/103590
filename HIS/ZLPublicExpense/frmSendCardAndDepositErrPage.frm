VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSendCardAndDepositErrPage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picEndDate 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   7635
      ScaleHeight     =   330
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   3675
      Width           =   1320
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   148832259
         CurrentDate     =   40777
      End
   End
   Begin VB.PictureBox picStartDate 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   5925
      ScaleHeight     =   300
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   3645
      Width           =   1425
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   148832259
         CurrentDate     =   40777
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsErrList 
      Height          =   3615
      Left            =   765
      TabIndex        =   0
      Top             =   1515
      Width           =   4545
      _cx             =   8017
      _cy             =   6376
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSendCardAndDepositErrPage.frx":0000
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   7
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
      Begin VB.PictureBox picErrImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   60
         Width           =   210
         Begin VB.Image imgErrImg 
            Height          =   195
            Left            =   0
            Picture         =   "frmSendCardAndDepositErrPage.frx":0059
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsthis 
      Left            =   570
      Top             =   690
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSendCardAndDepositErrPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------
'��ڱ���
Private mlngModule As Long
Private mblnNotRefresh As Boolean
Private mfrmMain As Object
Private mintӦ�ó��� As Integer
'----------------------------------------------------------------------
'2.�˵���ر���
Private mblnNotChange As Boolean
Private mlngPreID As Long   '�ϴ�ѡ����쳣ID
Private mobjCombox As CommandBarComboBox  '�����б�
Private mobjDateLable As CommandBarControl  '���ڿؼ�
Private Const conMenu_Combox = 3820   '������
Private Const conMenu_StartDate = 3824    '��ʼ����
Private Const conMenu_EndDate = 3825    '��ֹ����
Private Const conMenu_LableRange = 3827    '��
Private Const conMenu_LableDate = 3826    '��ֹ����
Private mintDateType As Integer
Private mlng�쳣ID As Long
Private mbln����������Ա As Boolean

'----------------------------------------------------------
'�ӿ�:
'  1.zlRefreshData-����ˢ������
'  2.zlInit-��ʼ���ӿ�

'----------------------------------------------------------
Public Sub zlRefreshData()
    Call LoadErrDataToGrid
    
End Sub

Public Function zlInit(ByVal frmMain As Object, ByVal intӦ�ó��� As Integer, ByVal lngModule As Long, Optional bln����������Ա As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:intӦ�ó���-1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
    '    bln����������Ա-�����ȡ��������Ա���쳣����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-28 17:16:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mintӦ�ó��� = intӦ�ó���: Set mfrmMain = frmMain: mlngModule = lngModule
    zlInit = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Sub Form_Load()
    Call zlDefCommandBars
End Sub
 Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    vsErrList.Top = Top
    vsErrList.Left = Left: vsErrList.Width = Right - Left
    vsErrList.Height = Bottom - Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
End Sub

Private Sub vsErrList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "�쳣�б�", False
End Sub
 Private Sub vsErrList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Then Exit Sub
    zl_VsGridRowChange vsErrList, OldRow, NewRow, OldCol, NewCol, GRD_GOTFOCUS_COLORSEL
    If OldRow = NewRow Then Exit Sub
    With vsErrList
        mlng�쳣ID = 0
        If .Row < 0 Or .ColIndex("ID") < 0 Then Exit Sub
        mlng�쳣ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Sub
   
Private Sub vsErrList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsErrList
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsErrList_GotFocus()
    zl_VsGridGotFocus vsErrList, GRD_GOTFOCUS_COLORSEL
End Sub

Private Sub vsErrList_LostFocus()
    zl_VsGridLostFocus vsErrList, GRD_LOSTFOCUS_COLORSEL
End Sub

Private Sub vsErrList_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "�쳣�б�", False
End Sub
Private Function ExcuteErrOper(Optional ByVal bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���쳣����
    '���:bln����-�Ƿ������쳣
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-28 16:37:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, lng����ID As Long, lng�쳣ID As Long, bln�����쳣 As Boolean
    Dim bln�������� As Boolean
    On Error GoTo errHandle
    
    With vsErrList
        lng�쳣ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        bln�����쳣 = Trim(.TextMatrix(.Row, .ColIndex("�쳣��ʽ"))) = "�����쳣"
        bln�������� = Val(.TextMatrix(.Row, .ColIndex("ͬ��״̬"))) < 2
        If lng�쳣ID = 0 Then Exit Function
    End With
    If bln���� And Not bln�������� Then
        MsgBox "��ǰ�쳣��¼�ѵ��ýӿڻ��Ѳ������ã��������ϣ���㡾�쳣���ա�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If bln�����쳣 Then
       If Not bln���� Then
            MsgBox "��ǰ�쳣��¼Ϊ�����쳣��¼���������գ���㡾�쳣���ϡ�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
       End If
       If MsgBox("���Ƿ����Ҫ���ϵ�ǰ���쳣����ô?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
       If Excute_ReCancel(lng�쳣ID) = False Then Exit Function
       MsgBox "���ϳɹ�", vbInformation + vbOKOnly, gstrSysName
       ExcuteErrOper = True
       Call LoadErrDataToGrid
       Exit Function
    End If
    
    'int��������-0-����;1-�쳣����;2-�쳣����
    If frmSendCardAndDepositErrEdit.zlShowWindow(Me, IIf(bln����, 2, 1), lng�쳣ID, mlngModule) = False Then Exit Function
    ExcuteErrOper = True
    Call LoadErrDataToGrid
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function Excute_ReCancel(ByVal lng�쳣ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���쳣���˲���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-12-02 15:16:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllErrData As Collection, strSql As String, rsTemp As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim objService As clsService
    
    On Error GoTo errHandle
    
    If zlGetServiceObject(objService) = False Then Exit Function
    strSql = "" & _
    "Select ID, ��������,nvl( �Ƿ�����,0) as �Ƿ�����, ҵ��id, �Ƿ�����, ����id, ��ҳid, Ԥ������, ҽ�ƿ�����, �����id, ��������, ͬ��״̬, ������Ϣ, �Ǽ�ʱ��, ����Ա���� " & _
    "     From ���˽����쳣��¼ " & _
    "     Where ID =[1] "
    
    'int����:1-ҽ�ƿ�����;2-������Ϣ�Ǽ�;3-������Ժ �Ǽ�;4-ԤԼ�ҺŽ���
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, lng�쳣ID)
    If rsTemp.EOF Then
        MsgBox "��ȡ�쳣����ʧ�ܣ������򲢷�ԭ���������ջ����ϣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!�Ƿ�����)) <> 1 Then
        MsgBox "��ǰ�쳣��¼�����쳣�����ϼ�¼������!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllErrData = New Collection
    cllErrData.Add Array("�쳣ID", Val(Nvl(rsTemp!ID)))
    cllErrData.Add Array("��������", Val(Nvl(rsTemp!��������)))
    cllErrData.Add Array("ҵ��ID", Val(Nvl(rsTemp!ҵ��ID)))
     
    'ɾ��ҽ�ƿ��䶯��¼
    gcnOracle.BeginTrans: blnTrans = True
    If Zl_���˽����쳣��¼_Modify(2, cllErrData) = False Then
         gcnOracle.RollbackTrans: blnTrans = False: Exit Function
    End If
    
    If objService.zl_PatiSvr_DelCardChangeInfo(Val(Nvl(rsTemp!����ID)), Val(Nvl(rsTemp!ҵ��ID)), Val(Nvl(rsTemp!�����ID)), Trim(Nvl(rsTemp!��������))) = False Then
       gcnOracle.RollbackTrans: blnTrans = False: Exit Function
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    Call LoadErrDataToGrid
    Excute_ReCancel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsErrList_DblClick()
    Call ExcuteErrOper
End Sub
Private Sub imgErrImg_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    vRect = zlControl.GetControlRect(picErrImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picErrImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsErrList, lngLeft, lngTop, imgErrImg.Height)
    zl_vsGrid_Para_Save mlngModule, vsErrList, Me.Name, "�쳣�б�", False
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl, cbrToolBar As CommandBar
    Dim objCustomControl As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHand:
    
    Set cbsthis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsthis.VisualTheme = xtpThemeOffice2003
    With cbsthis.Options
        
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
        
    End With
    cbsthis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsthis.DeleteAll
    Set cbrToolBar = cbsthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    
    cbrToolBar.EnableDocking xtpFlagStretched
    
    
    With cbrToolBar.Controls
        Set mobjCombox = .Add(xtpControlComboBox, conMenu_Combox, "ȱʡ��ʾ")
        mobjCombox.Width = 2600 / Screen.TwipsPerPixelX
        mobjCombox.Style = xtpComboLabel
        
        Set mobjDateLable = .Add(xtpControlLabel, conMenu_LableDate, "2017-01-01 23:59:59-2017-02-02 23:59:59")
        mobjDateLable.Visible = False
 
         Set objCustomControl = .Add(xtpControlCustom, conMenu_StartDate, "")
        objCustomControl.Handle = picStartDate.hWnd
        
        Set objControl = .Add(xtpControlLabel, conMenu_LableRange, " �� ")
        Set objCustomControl = .Add(xtpControlCustom, conMenu_EndDate, "")
        objCustomControl.Handle = picEndDate.hWnd
        

        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): objControl.BeginGroup = True
        objControl.Flags = xtpFlagRightAlign
     
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ErrReBalance, "�쳣����(&R)")
        objControl.Flags = xtpFlagRightAlign: objControl.IconId = 231
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ErrCancelBalance, "�쳣����(&C)")
        objControl.Flags = xtpFlagRightAlign
    End With
    
    For Each objControl In cbrToolBar.Controls
        If objControl.Type <> xtpControlLabel And objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlComboBox Then
          objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    Call InitErrDate
    Call LoadErrDataToGrid
    zlDefCommandBars = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean, objCombox As CommandBarComboBox
    On Error GoTo errHandle
    Select Case Control.ID
    Case conMenu_View_Refresh   'ˢ��
        Call LoadErrDataToGrid
    Case conMenu_Edit_ErrReBalance  '�쳣����
        ExcuteErrOper False
    Case conMenu_Edit_ErrCancelBalance  '�쳣����
        ExcuteErrOper True
    Case conMenu_Combox
        Set objCombox = Control
        mintDateType = objCombox.ListIndex
        Call ChargeComboxDate(mintDateType)
        Call LoadErrDataToGrid
        
    Case Else
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function ChargeComboxDate(ByVal intListIndex As Integer)
    Dim dtStartDate As Date, dtEndDate As Date
    
    intListIndex = intListIndex - 1
    Select Case intListIndex
        Case 0 '�����쳣
        Case 1 '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
        Case 2 '���2��
            dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 3 '���3��
            dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 4  '���һ��
            dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case 5  '����
            dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
            dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
        Case Else
            dtStartDate = CDate(Format(dtpStartDate.value, "yyyy-mm-dd") & " 00:00:00")
            dtEndDate = CDate(Format(dtpEndDate.value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    If Not mobjDateLable Is Nothing Then
        mobjDateLable.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS") & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
    End If
End Function

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Refresh   'ˢ��
    Case conMenu_Edit_ErrReBalance  '�쳣����
        Control.Enabled = True ' mlng�쳣ID <> 0
    Case conMenu_Edit_ErrCancelBalance  '�쳣����
        Control.Enabled = mlng�쳣ID <> 0
    Case conMenu_Combox
    Case conMenu_StartDate
       Control.Visible = mintDateType = 7
    Case conMenu_EndDate
       Control.Visible = mintDateType = 7
    Case conMenu_LableRange
       Control.Visible = mintDateType = 7
    Case conMenu_LableDate
       Control.Visible = mintDateType <> 7 And mintDateType <> 1
    Case Else
    End Select
    Exit Sub
End Sub
Private Function LoadErrDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����쳣���ݸ�����
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-11-05 22:21:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, lngID As Long, i As Long
    Dim strSql As String, strPreCardNO  As String
    Dim rsTemp As ADODB.Recordset
     
    On Error GoTo errHandle
    
     Call zlCommFun.ShowFlash("���ڼ��ز����쳣��������,���Ե�...", Me)
    
    Select Case mintDateType - 1
    Case 0 '�����쳣
        dtStartDate = CDate(Format("1900-01-01", "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format("3000-01-01", "yyyy-mm-dd") & " 23:59:59")
    Case 1 '����
        dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
    Case 2 'ǰһ��������
        dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 3 'ǰ����������
        dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 4  'ǰһ��������
        dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case 5  '����
        dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm") & "-01 00:00:00")
        dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
    Case Else
        dtStartDate = CDate(Format(dtpStartDate.value, "yyyy-mm-dd") & " 00:00:00")
        dtEndDate = CDate(Format(dtpEndDate.value, "yyyy-mm-dd") & " 23:59:59")
    End Select
    
    
    strSql = " " & _
    "   Select '' as ��־,A.ID,decode(a.��������,1,'ҽ�ƿ�����',2,'������Ϣ�Ǽ�',3,'������Ժ�Ǽ�',4,'ԤԼ�ҺŽ���','����')as  ��������, " & _
    "       decode(nvl(a.�Ƿ�����,0),1,'�����쳣','�շ��쳣') as �쳣��ʽ,a.ҵ��ID,decode(nvl(a.�Ƿ�����,0),0,'','��') as ������,a.����id,a.��ҳid, " & _
    "       a.����,a.�Ա�,a.����,a.�����,a.סԺ��,a.Ԥ������,a.Ԥ�����, " & _
    "       a.ҽ�ƿ�����,a.���ѽ��,a.�����ID,a.���������,a.��������,a.����Ա����,a.�Ǽ�ʱ��,a.ͬ��״̬,a.������Ϣ " & _
    "   From ���˽����쳣��¼ A " & _
    "   Where  A.��������=[1] And A.�Ǽ�ʱ�� between [2] and [3] " & IIf(mbln����������Ա, "", "And ����Ա����=[4]") & _
    "   Order by Decode(����Ա����,[4],1,0)"
    
    Set rsTemp = zlDatabase.OpenSQLRecordLob(strSql, Me.Caption, mintӦ�ó���, dtStartDate, dtEndDate, UserInfo.����)
       
       
    With vsErrList
        If .Row > 0 And .ColIndex("ID") >= 0 Then
            mlngPreID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            If mlngPreID <> 0 And .ColIndex("ID") >= 0 Then
                lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            End If
        End If
        
        .Redraw = flexRDNone
        .Clear: .Rows = 2: .Cols = 1
        .Cell(flexcpForeColor, 1, .FixedCols - 1, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpText, 0, 0, .Rows - 1, .Cols - 1) = ""
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        .Row = 1

        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            ''ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            .ColData(i) = "0||0"  '����ѡ��
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = True
                .ColData(i) = "-1||1"    '����ѡ��
            ElseIf .ColKey(i) Like "*ʱ��" Or .ColKey(i) Like "*����" Or .ColKey(i) = "״̬" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .ColData(.ColIndex("��������")) = "1||0": .ColData(.ColIndex("��־")) = "-1||1"
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
        zl_vsGrid_Para_Restore mlngModule, vsErrList, Me.Name, "�쳣����", False, True
        .ColWidth(.ColIndex("��־")) = 285
        .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    Call vsErrList_AfterRowColChange(-1, 0, vsErrList.Row, 0)
    Call zlCommFun.StopFlash
    LoadErrDataToGrid = True = True
    Exit Function
errHandle:
    Call zlCommFun.StopFlash
    vsErrList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitErrDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���쳣��������
    '����:���˺�
    '����:2019-11-05 22:16:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strValue As String
    Call GetRegInFor(g˽��ģ��, Me.Name, "�쳣���ݲ�ѯ", strValue)
    i = Val(strValue)
    With mobjCombox
        .Clear
        .AddItem "�����쳣���"
        .ListIndex = 1
        If i = 0 Then .ListIndex = 1
        .AddItem "����"
        If i = 1 Then .ListIndex = 2
        .AddItem "�������"
        If i = 2 Then .ListIndex = 3
        .AddItem "�������"
        If i = 3 Then .ListIndex = 4
        .AddItem "���һ��"
        If i = 4 Then .ListIndex = 5
        .AddItem "����"
        If i = 5 Then .ListIndex = 6
        .AddItem "�Զ���ʱ�䷶Χ"
        
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.value = dtpEndDate.MaxDate
        dtpStartDate.value = DateAdd("d", -7, dtpEndDate.MaxDate)
        mintDateType = mobjCombox.ListIndex
    End With
End Sub
