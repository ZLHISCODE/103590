VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmChargeRollingCurtainMgr 
   Caption         =   "�շ����ʹ���"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmChargeRollingCurtainMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   15105
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      ScaleHeight     =   255
      ScaleWidth      =   3495
      TabIndex        =   5
      Top             =   1425
      Visible         =   0   'False
      Width           =   3495
      Begin zL9CashBill.ComboxExpend cboType 
         Height          =   255
         Left            =   750
         TabIndex        =   6
         Top             =   0
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   450
         Appearance      =   0
         BorderStyle     =   1
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "����"
         FontSize        =   9
         Locked          =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   210
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   1
      Top             =   2790
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   2
         Top             =   -15
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picBalanceList 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1005
      ScaleHeight     =   1695
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   705
      Width           =   3210
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1800
         Left            =   270
         TabIndex        =   4
         Top             =   180
         Width           =   10740
         _cx             =   18944
         _cy             =   3175
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmChargeRollingCurtainMgr.frx":0442
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   10065
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      SimpleText      =   $"frmChargeRollingCurtainMgr.frx":04BC
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeRollingCurtainMgr.frx":0503
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18997
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "���˺�"
            TextSave        =   "���˺�"
            Object.ToolTipText     =   "��ǰ����Ա:���˺�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   15
      Top             =   -75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmChargeRollingCurtainMgr.frx":0D97
      Left            =   660
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeRollingCurtainMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
'Private mobjComBar As CommandBarComboBox
Private mfrmChargeRollingCurtain As frmChargeRollingCurtain
Private mfrmHistory As frmChargeRollingCurtainHistory
Private mstr������Ա���� As String
Private mstrPrevRollingType As String
Private Enum mPgIndex
    EM_PG_�����б� = 250101
    EM_PG_��ʷ�б� = 250102
End Enum
Private Enum mPaneIndex
    EM_PN_�ݴ�� = 1
    EM_PN_��ϸ�б� = 2
End Enum
Private mstrPreDate As String   '�ϴ�����ʱ��
Private mblnFirst As Boolean, mblnNotice As Boolean
Private mblnԤ���ֱ����� As Boolean

Private Function GetPreRollingCurtainTime(ByVal strTYPE As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ϴ�����ʱ��
    '���:intType-(0-�������(��ȫ������),1-�շ�,2-Ԥ��(21-����Ԥ��,22-סԺԤ��),3-����,4-�Һ�,5-���￨)
    '����:���ظ�ʽyyyy-mm-dd hh24:mi:ss
    '����:���˺�
    '����:2015-03-03 10:43:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    '����:
    '    1.�����ǰ��ָ���������ʱ,���¹�����(����ֹʱ��Ϊ׼):
    '      1) ���������Ч�����ʼ�¼ʱ
    '          a)�����ǰ������Ա�����������������ʼ�¼��,�������һ������ʱ��Ϊ׼
    '          b)�����ǰ������Ա���������������ʼ�¼�����һ�����ʼ�¼����ֹʱ��>��ǰ�������һ�����ʼ�¼���տ�ʱ���,���������������һ�����ʼ�¼����ֹʱ��Ϊ׼
    '          c)�����ǰ������Ա���������������ʼ�¼�����һ�����ʼ�¼����ֹʱ��<��ǰ�������һ�����ʼ�¼���տ�ʱ���,���Ե�ǰ�������һ�����ʼ�¼���տ�ʱ��Ϊ׼
    '      2)�����������Ч���ʼ�¼ʱ
    '          a)�����ǰ������Ա���ڰ�����������ʵ���Ч��¼ʱ,�������һ�����ʼ�¼����ֹʱ��Ϊ׼
    '          b)��������ڰ�����������ʵ���Ч��¼ʱ,�������3���¹�����
    '    2.�����ǰ�������������ʱ,�����¹�����
    '          a)������ڰ�����������ʵ���Ч��¼ʱ,�������һ�����ʼ�¼����ֹʱ��Ϊ׼
    '          b)��������ڰ�����������ʵ���Ч��¼ʱ,�������3���¹�����
    '    3.����������ʹ����¼,�������ʹ����¼�ĵǼ�ʱ��Ϊ׼
    '    4.δ������,ȱʡΪ��ȡ���ý�ʱ��
    '    5.���δ���ñ��ý�ģ�ȱʡʱ��Ϊ��ǰʱ��-1���µ�����������ϴ�ת��ʱ��
    
    '��¼����:1-�շ�Ա���˼�¼(�ɿ���)��2-�������տ��¼;3-���������˼�¼(��ɿ���); _
    '4-�����տ��¼;5-�ֹ��ɿ�(��ԭ�����ܱ��ֲ���);6-���ʹ���(�л�����ģʽ�󣬱�����������㣬��Ϊ��������¼)
    If strTYPE = "" Then Exit Function
    
    strSQL = "Select  Zl_Rollingcurtain_Lastdate([1],[2]) as ����ʱ�� From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����, strTYPE)
    If Not rsTemp.EOF Then
        GetPreRollingCurtainTime = Format(rsTemp!����ʱ��, "yyyy-mm-dd HH:MM:SS")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDefaultTime() As String
    '��ȡȱʡ���˽���ʱ��
    On Error GoTo errH
    Dim strSQL As String, rsTemp As ADODB.Recordset, strValue As String
    Dim datValue As Date, datNow As Date
    
    strValue = zlDatabase.GetPara("ȱʡ����ʱ��", glngSys, mlngModule, "", dtpTime, InStr(1, mstrPrivs, ";��������;") > 0)
    
    If strValue = "" Then GetDefaultTime = "": Exit Function
    
    datNow = zlDatabase.Currentdate
    strSQL = "Select 1 From ��Ա�սɼ�¼ Where ��ֹʱ�� >= [1] And ����ʱ�� Is Null And �տ�Ա = [2]"
    datValue = CDate(Format(datNow - IIf(Format(datNow, "hh:mm:ss") >= Format(strValue, "hh:mm:ss"), 0, 1), "yyyy-MM-dd") & " " & Format(strValue, "hh:mm:ss"))
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datValue, UserInfo.����)
    
    If Not rsTemp.EOF Then Exit Function
    
    GetDefaultTime = Format(datNow - IIf(Format(datNow, "hh:mm:ss") >= Format(strValue, "hh:mm:ss"), 0, 1), "yyyy-MM-dd") & " " & Format(strValue, "hh:mm:ss")

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function InitData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-11 14:20:10
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '��ȡ�ϴ�����ʱ��
    mstrPreDate = GetPreRollingCurtainTime(GetRollingType)
    Call LoadCurBalanceData
    
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadCurBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�ǰ��Ա�ɿ��������
    '����:���˺�
    '����:2015-03-03 12:06:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  decode(nvl(b.����,0),1,1,2,2,3,10,4,11,4) as ���, " & _
    "               A.���㷽ʽ, A.���, A.�ϴ�����ʱ�� " & _
    "   From ��Ա�ɿ���� A,���㷽ʽ  B" & _
    "   Where a.���㷽ʽ=b.����(+)  And A.�տ�Ա = [1] And A.���� = 1" & _
    "   Order by �ϴ�����ʱ�� Desc,���,���㷽ʽ"
    '--1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,
    '   5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
    Call InitGrid
    With vsBalance
        .Clear 1
        .Rows = 2: lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
            .TextMatrix(lngRow, .ColIndex("�ݴ���")) = Format(Val(Nvl(rsTemp!���)), "#,###0.00;-#,###0.00;0.00;-0.00")
            lngRow = lngRow + 1: .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "�ݴ���б�", False
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function zlShowChargeRollingCourtain(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
        '-------------------------------------------------------------------------------------------------
        '����:�շ�Ա���˽ӿ�
        '���:frmMain-���õ�������
        '       strOperatorName-�շ�Ա,Ϊ��ʱ,ȱʡΪ��ǰ����Ա
        '����:�շ����ʳɹ�һ������,����true,���򷵻�False
        '����:���˺�
        '����:2013-08-13 10:31:00
        '˵��:
        '-------------------------------------------------------------------------------------------------
        mlngModule = lngModule: mstrPrivs = strPrivs
        mstrPreDate = ""
        If CheckDepend = False Then Exit Function
        If InitData = False Then
            Err = 0: On Error Resume Next
            Unload Me
            Err.Clear: Err = 0
            Exit Function
        End If
        '��ʼ������
        Call InitFace
        mblnFirst = True
        If frmMain Is Nothing Then
            Me.Show
        Else
            Me.Show , frmMain
        End If
End Function

Public Sub BHShowList(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngMain As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2013-10-17 18:17:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mstrPreDate = ""
    If CheckDepend = False Then Exit Sub
    If InitData = False Then
        Err = 0: On Error Resume Next
        Unload Me
        Err.Clear: Err = 0
        Exit Sub
    End If
    '��ʼ������
    Call InitFace
    mblnFirst = True
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
    Me.ZOrder 0
End Sub

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:���˺�
    '����:2013-09-03 10:29:48
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    '��ǰ�ݴ��
    With vsBalance
        .Rows = 3: .Cols = 2
       .TextMatrix(0, 0) = "���㷽ʽ"
       .TextMatrix(0, 1) = "�ݴ���"
       For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
       Next
       .TextMatrix(1, .ColIndex("���㷽ʽ")) = "�ֽ�"
       .TextMatrix(1, .ColIndex("�ݴ���")) = "100"
       .TextMatrix(2, .ColIndex("���㷽ʽ")) = "֧Ʊ"
       .TextMatrix(2, .ColIndex("�ݴ���")) = "100"
       .AutoSizeMode = flexAutoSizeColWidth
       Call .AutoSize(0, .Cols - 1)
    End With
End Sub
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-09-03 14:43:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitPanel
    Call InitPage
    Set dkpMan.TabPaintManager.Font = vsBalance.Font
    Set dkpMan.PaintManager.CaptionFont = vsBalance.Font
    dkpMan.PanelPaintManager.StaticFrame = True
    stbThis.Panels(3).Text = UserInfo.����
    stbThis.Panels(3).ToolTipText = "��ǰ����Ա:" & UserInfo.����
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 18:21:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
    Dim objCustom As CommandBarControlCustom
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "����(&Z)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "��������(&D)")
        mcbrControl.IconId = 229
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮(&E)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "�ش�ɿ���(&R)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "�鿴��ϸ����(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("R"), conMenu_Edit_ChargeBook_Reprint
        .Add FCONTROL, Asc("S"), conMenu_Edit_RollingCurtain
        .Add FCONTROL, Asc("D"), conMenu_Edit_RollingCurtain_Cancel
        .Add FCONTROL, Asc("T"), conMenu_View_Detail
        .Add 0, VK_F6, conMenu_Edit_CheckCash
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        '.AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "����"): mcbrControl.BeginGroup = True
         mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Cancel, "��������"): mcbrControl.BeginGroup = True
         mcbrControl.IconId = 229
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮")
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "��ѯ��ϸ")
        mcbrControl.IconId = 2322
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
        Set objCustom = .Add(xtpControlCustom, conMenu_COMBOX_INTERFACE, "�������")
        objCustom.Flags = xtpFlagRightAlign
        objCustom.HideFlags = xtpNoHide
        objCustom.Handle = picType.hWnd
        objCustom.BeginGroup = True
        picType.BackColor = CommandBarsGlobalSettings.ColorManager.Color(XPCOLOR_TOOLBAR_FACE)
        
        With cboType
            .Clear
            .AddItem "0", "�������", True, True, True
            
            If InStr(1, mstr������Ա����, ",�����շ�Ա,") > 0 Then
                .AddItem 1, "�շ�", False, True, True
            End If
            If InStr(1, mstr������Ա����, ",Ԥ���տ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",��Ժ�Ǽ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",�����Ǽ���,") > 0 Then
                If mblnԤ���ֱ����� Then
                    .AddItem 21, "����Ԥ��", False, True, True
                    .AddItem 22, "סԺԤ��", False, True, True
                Else
                    .AddItem 2, "Ԥ��", False, True, True
                End If
            End If
            If InStr(1, mstr������Ա����, ",סԺ����Ա,") > 0 Then
                .AddItem 3, "����", False, True, True
            End If
            If InStr(1, mstr������Ա����, ",����Һ�Ա,") > 0 Then
                .AddItem 4, "�Һ�", False, True, True
            End If
            If InStr(1, mstr������Ա����, ",����Һ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",��Ժ�Ǽ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",�����Ǽ���,") > 0 Then
                .AddItem 5, "���￨", False, True, True
            End If
            If InStr(1, mstr������Ա����, ",����Һ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",��Ժ�Ǽ�Ա,") > 0 _
               Or InStr(1, mstr������Ա����, ",�����Ǽ���,") > 0 Then
                .AddItem 6, "���ѿ�", False, True, True
            End If
        End With
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    With dkpMan
        'Set .ImageList = zlCommFun.GetPubIcons
        Set objPane = .CreatePane(mPaneIndex.EM_PN_��ϸ�б�, 400, 400, DockLeftOf, Nothing)
        objPane.Title = "������Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
      
        Set objPane = .CreatePane(mPaneIndex.EM_PN_�ݴ��, 100, 100, DockRightOf, objPane)
        objPane.Title = "��ǰ�ݴ��": objPane.Options = PaneNoCloseable
        objPane.Handle = picBalanceList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Dim strRollingType As String
    Err = 0: On Error GoTo Errhand:
    If mfrmChargeRollingCurtain Is Nothing Then
        Set mfrmChargeRollingCurtain = New frmChargeRollingCurtain
        Load mfrmChargeRollingCurtain
    End If
    strRollingType = GetRollingType
    '��ʼ������
    Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.����, strRollingType, GetDefaultTime)
    If mfrmHistory Is Nothing Then
        Set mfrmHistory = New frmChargeRollingCurtainHistory
        Load mfrmHistory
    End If
    Call mfrmHistory.zlInitVar(Me, mlngModule, mstrPrivs)
    Set objItem = tbPage.InsertItem(EM_PG_�����б�, "����", mfrmChargeRollingCurtain.hWnd, 0)
    objItem.Tag = EM_PG_�����б�
    Set objItem = tbPage.InsertItem(EM_PG_��ʷ�б�, "��ʷ������Ϣ", mfrmHistory.hWnd, 0)
    objItem.Tag = EM_PG_��ʷ�б�
     With tbPage
        Set tbPage.PaintManager.Font = vsBalance.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboType_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    If GetRollingType = "" Then Node.Checked = True: Node.Selected = True
    Call ReloadData(Format(mfrmChargeRollingCurtain.dtpEndDate.Value, "yyyy-mm-dd hh:mm:ss"))
    mstrPrevRollingType = GetRollingType
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
     If Action = PaneActionAttached Then Cancel = True: Exit Sub
     If Action = PaneActionAttaching Then Cancel = True: Exit Sub
     If Action = PaneActionFloated Then Cancel = True: Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    DoEvents
    Call DefaultSetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�����б�
        Call mfrmChargeRollingCurtain.MainKeyDown(KeyCode, Shift)
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mblnԤ���ֱ����� = Val(zlDatabase.GetPara("Ԥ�����ʰ������סԺ�ֱ�����", glngSys, glngModul, "0")) = 1
    RestoreWinState Me, App.ProductName
    Call zlDefCommandBars
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmChargeRollingCurtain Is Nothing Then Unload mfrmChargeRollingCurtain
    If Not mfrmHistory Is Nothing Then Unload mfrmHistory
    Set mfrmChargeRollingCurtain = Nothing
    Set mfrmHistory = Nothing
End Sub
Private Sub picBalanceList_Resize()
    Err = 0: On Error Resume Next
    With vsBalance
        .Top = picBalanceList.ScaleTop + 20
        .Left = picBalanceList.ScaleLeft + 20
        .Width = picBalanceList.ScaleWidth - .Left * 2
        .Height = picBalanceList.ScaleHeight - .Top * 2 '.RowHeight(0) + .RowHeight(1) + 3 * 50
    End With
End Sub
 Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub ParameterSet()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2013-09-12 15:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strDefaultTime As String
    If frmChargeRollingCurtainSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    strDefaultTime = GetDefaultTime
    If strDefaultTime <> "" Then
        mfrmChargeRollingCurtain.dtpEndDate.Value = Format(strDefaultTime, "yyyy-MM-dd hh:mm:ss")
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub RollingCurtain()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʴ���
    '����:���˺�
    '����:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String
    If Val(tbPage.Selected.Tag) = EM_PG_��ʷ�б� Then Exit Sub
    Call mfrmChargeRollingCurtain.SaveDataWithCheck
End Sub
Private Sub RollingCurtainCancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ϴ���
    '����:���˺�
    '����:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, blnDel As Boolean
    Dim strRollingType As String
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then Exit Sub
    strRollingType = GetRollingType
    If mfrmHistory.CancelData() Then
        '�������ʺ�,���¶�ȡ�ϴ�����ʱ��
        Call InitData
        Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.����, strRollingType)
        Call mfrmChargeRollingCurtain.RefreshPage
        Exit Sub
    End If
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '�˳�(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '��ӡ����
        Case conMenu_File_Preview: Call zlPrintRpt(2)  'Ԥ��(&V)
        Case conMenu_File_Print: Call zlPrintRpt(1) '��ӡ(&P)
        Case conMenu_File_Excel: Call zlPrintRpt(3)  '�����&Excel��
        Case conMenu_File_Parameter: Call ParameterSet '��������
        Case conMenu_Edit_RollingCurtain: Call RollingCurtain  '����(&Z)
        Case conMenu_Edit_RollingCurtain_Cancel: Call RollingCurtainCancel '��������(&D)
        Case conMenu_Edit_CheckCash: Call CheckCash '�ֽ�㳮(&E)
        Case conMenu_Edit_ChargeBook_Reprint:  Call RePrintBill '�ش�ɿ���(&R)
        Case conMenu_View_Detail: Call ShowChargeList '�鿴��ϸ����(&V)
        Case conMenu_View_Refresh: zlRefresh 'ˢ��(&R)
        Case conMenu_View_StatusBar '״̬��(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_COMBOX_INTERFACE   '���ѡ��
            Call ReloadData
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                'ִ�з�������ǰģ��ı���
                Call CallCustomRpt(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub ReloadData(Optional ByVal strTime As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�������
    '����:���˺�
    '����:2015-03-03 12:15:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRollingType As String
    
    On Error GoTo errHandle
    strRollingType = GetRollingType
    If strRollingType = "" Then
        MsgBox "������ѡ��һ�����������ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    mstrPreDate = GetPreRollingCurtainTime(strRollingType)
    Call mfrmChargeRollingCurtain.zlInitVar(Me, mlngModule, mstrPrivs, mstrPreDate, UserInfo.����, strRollingType, strTime)
    Call mfrmChargeRollingCurtain.RefreshPage
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHavePrivs As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_RollingCurtain ' "����(&Z)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_�����б�
    Case conMenu_Edit_RollingCurtain_Cancel ' "��������(&D)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "��������")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_��ʷ�б�
        If Control.Enabled Then
            Control.Enabled = mfrmHistory.GetChargeRollingCurtainID <> 0 _
                And Not mfrmHistory.GetChargeRollingCurtainDel
        End If
    Case conMenu_View_Detail    '�鿴��ϸ����(&V)
        If Val(tbPage.Selected.Tag) = EM_PG_��ʷ�б� Then
            With mfrmHistory.vsRollingCurtain
                Control.Enabled = .RowSel >= 1 And .TextMatrix(.RowSel, .ColIndex("���ʵ���")) <> ""
            End With
        Else
            Control.Enabled = True
        End If
    Case conMenu_Edit_CheckCash ' "�ֽ�㳮(&E)")
        
    Case conMenu_Edit_ChargeBook_Reprint ' "�ش�ɿ���(&R)")
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ش�ɿ���") And zlStr.IsHavePrivs(mstrPrivs, "�ɿ����ӡ")
        Control.Visible = blnHavePrivs
        Control.Enabled = blnHavePrivs And Val(tbPage.Selected.Tag) = EM_PG_��ʷ�б�
        If Control.Enabled Then
            Control.Enabled = mfrmHistory.GetChargeRollingCurtainID <> 0 _
                And Not mfrmHistory.GetChargeRollingCurtainDel
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_COMBOX_INTERFACE
        If Not tbPage.Selected Is Nothing Then
            Control.Visible = Val(tbPage.Selected.Tag) <> EM_PG_��ʷ�б�
        Else
            Control.Visible = False
        End If

    End Select
End Sub
Private Function CheckDepend() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2013-09-04 17:10:03
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    CheckDepend = False
    
    On Error GoTo errHandle
     mstr������Ա���� = ""
    gstrSQL = "" & _
    "   Select  B.ID,A.��Ա����  " & _
    "   From ��Ա����˵�� A, ��Ա�� B " & _
    "   Where A.��Աid = B.ID And A.��Ա���� In ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���') And B.ID=[1] " & _
    "   Order By ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ����Ա�Ƿ�Ϊ��Ӧ������Ա", UserInfo.ID)
    If rsTemp.EOF Then
        ShowMsgbox "�㲻�߱�������Һ�Ա,�����շ�Ա,Ԥ���տ�Ա,סԺ����Ա,��Ժ�Ǽ�Ա,�����Ǽ��ˡ������ʣ�����ʹ�ø�ģ�飡"
        rsTemp.Close
        Exit Function
    End If
    Do While Not rsTemp.EOF
        If InStr(mstr������Ա���� & ",", "," & rsTemp!��Ա���� & ",") = 0 Then
            mstr������Ա���� = mstr������Ա���� & "," & rsTemp!��Ա����
        End If
        rsTemp.MoveNext
    Loop
    If mstr������Ա���� <> "" Then mstr������Ա���� = mstr������Ա���� & ","
    
    Set rsTemp = Get���㷽ʽ
    rsTemp.Filter = "����=1"
    If rsTemp.EOF Then
        rsTemp.Filter = 0
        ShowMsgbox "���㷽ʽ�в�����һ�������ֽ����ʵĽ��㷽ʽ,���ڽ��㷽ʽ����������!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Filter = 0
    rsTemp.Close
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub picType_Resize()
    On Error Resume Next
    With cboType
        .Left = 15
        .Top = 15
        .Width = picType.ScaleWidth - 30
        .Height = picType.ScaleHeight - 30
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call DefaultSetFocus
End Sub

Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "�ݴ���б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "�ݴ���б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub zlPrintRpt(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б�
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objRow As New zlTabAppRow, bytPrn As Byte
    Dim i As Long, lngRow As Long, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then
        '��ӡ������Ϣ
        Call mfrmChargeRollingCurtain.zlPrint(bytMode)
        Exit Sub
    End If
    Call mfrmHistory.zlPrint(bytMode)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub RePrintBill()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش�ɿ���
    '����:���˺�
    '����:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmHistory.RePrintBill
End Sub
Private Sub CheckCash()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֽ�㳮
    '����:���˺�
    '����:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then
        dblMoney = mfrmChargeRollingCurtain.GetCashMoney
    End If
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub
Private Sub zlRefresh()
    '���½�������ˢ��
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then
        Call mfrmChargeRollingCurtain.zlRefresh
    Else
        Call mfrmHistory.zlRefresh
    End If
    Call DefaultSetFocus
End Sub

Public Sub RefreshBasic()
    Call InitData
End Sub
Private Function GetRollingType() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:�����������
    '����:���˺�
    '����:2015-03-06 10:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTYPE As String
    On Error GoTo errHandle
    strTYPE = cboType.GetNodesCheckedDatas(False)
    GetRollingType = strTYPE
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowChargeList()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '����:���˺�
    '����:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRollingType As String
    
    On Error GoTo errHandle
    
    strRollingType = GetRollingType
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then
         Call mfrmChargeRollingCurtain.ShowChargeList(Me, strRollingType)
         Exit Sub
    End If
    '��ʷ������ʾ
    Call mfrmHistory.ShowChargeList(Me)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CallCustomRpt(ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo errHandle
    If Val(tbPage.Selected.Tag) = EM_PG_�����б� Then
         Call mfrmChargeRollingCurtain.CallCustomRpt(Me, lngSys, strRptCode)
         Exit Sub
    End If
    '��ʷ������ʾ
    Call mfrmHistory.CallCustomRpt(Me, lngSys, strRptCode)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub DefaultSetFocus()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���ȱʡ��λ
    '����:���˺�
    '����:2013-10-16 14:25:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_�����б�
        picType.Visible = True
        mfrmChargeRollingCurtain.zlDefaultSetFocus
    Case Else
        picType.Visible = False
        mfrmHistory.zlDefaultSetFocus
    End Select
End Sub
