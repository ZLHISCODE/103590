VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "ZLIDKIND.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmBlackListRecordManage 
   BorderStyle     =   0  'None
   Caption         =   "���˲�����¼"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList imgList16 
      Left            =   6120
      Top             =   7740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackListRecordManage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackListRecordManage.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5310
      Left            =   465
      ScaleHeight     =   5310
      ScaleWidth      =   9060
      TabIndex        =   0
      Top             =   1725
      Width           =   9060
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3705
         _Version        =   589884
         _ExtentX        =   6535
         _ExtentY        =   2514
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTitle 
         Height          =   360
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   7905
         _Version        =   589884
         _ExtentX        =   13944
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "���˲�����¼"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H8000000C&
         Height          =   735
         Left            =   5040
         Top             =   720
         Width           =   405
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGridPrint 
      Height          =   555
      Left            =   12990
      TabIndex        =   3
      Top             =   2055
      Visible         =   0   'False
      Width           =   645
      _cx             =   1961559154
      _cy             =   1961558995
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin zlIDKind.PatiIdentify patiFind 
      Height          =   345
      Left            =   10620
      TabIndex        =   4
      Top             =   375
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBlackListRecordManage.frx":0B34
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "���￨"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
End
Attribute VB_Name = "frmBlackListRecordManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String
Private mstr��Ϊ��� As String
 
Private Enum mEnm_RptHeadCol
    COL_ID = 0
    COL_ͼ��
    COL_��Ϊ���
    COL_��������
    COL_����
    COL_�Ա�
    COL_��������
    COL_�����
    COL_����ʱ��
    COL_����ԭ��
    COL_������ϸ˵��
    COL_����ʱ��
    COL_������Ϣ
    COL_�Ǽ���
    COL_������
    COL_����ʱ��
    COL_����ԭ��
    COL_�Ƿ�̶�
End Enum
Private mblnShowCancelRecord As Boolean '�Ƿ���ʾ�ѳ����Ĳ�����¼
Private mlngPreSelID As Long
Private mintFindType As Integer
Private mrs������¼ As Recordset
Private mcllFilter As Collection    '��������

Public Event zlActivate(ByVal frmSubForm As Form) '�¼�����
Public Event zlShowStatusText(ByVal bytPancel As Byte, ByVal strText As String)  '��ʾ״̬���ı�
Public Sub zlCancelBands()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ؼ����
    '����:���˺�
    '����:2018-11-15 15:48:53
    '��Ҫ�����ؽ�ǰ��ɾ���ؼ��󣬿��ܴ��ڰ󶨵Ŀؼ����ڹ�������������У����ɾ��ʱ������ؼ�һ��ɾ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrToolBar As CommandBar
    On Error GoTo errHandle
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "������")
    If cbrToolBar Is Nothing Then Exit Sub
    cbrToolBar.Controls.DeleteAll
    Set patiFind.Container = Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Public Sub zlInitComm(frmMain As Form, cbsThis As Object, ByVal strPrivs As String, ByVal lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ӿ�
    '���:objPati-����������
    '     cbsThis-�˵�����
    '     strPrivs-Ȩ�޴�
    '     lngModule-ģ���
    '����:���˺�
    '����:2018-11-08 11:28:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    Set mfrmMain = frmMain: Set mcbsMain = cbsThis
    mstrPrivs = strPrivs: mlngModule = lngModule
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function zlLoadData(ByVal str��Ϊ��� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-13 15:33:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strType As String
    On Error GoTo errHandle
    mstr��Ϊ��� = str��Ϊ���
    zlLoadData = LoadRecordDataToGrid(str��Ϊ���, mcllFilter)
     
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    
    Err = 0: On Error GoTo errHandle
 
 
 
    '     '�ļ��˵�
    '    '-----------------------------------------------------
    '    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    '    With cbrMenuBar.CommandBar.Controls
    '        '���������Excel֮��
    '        Set cbrControl = .Find(, conMenu_File_Excel)
    ''        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
    '    End With
    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���Ӳ�����¼(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ĳ�����¼(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��������¼(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "����������¼(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ������������¼(&T)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����(&F)", cbrControl.Index): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾ�ѳ����Ĳ�����¼(&S)", cbrControl.Index)
        cbrControl.Checked = mblnShowCancelRecord
        cbrControl.BeginGroup = True
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = GetCommbarFromName(mcbsMain, "������")
    If cbrToolBar Is Nothing Then
        Set cbrToolBar = mcbsMain.Add("������", xtpBarTop)
    End If
    
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup And cbrControl.Index > 1 Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ������", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����", cbrControl.Index + 1): cbrControl.BeginGroup = True
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
 
    '���󶨵Ŀؼ����붯̬���أ���Ϊ������һ����ɾ�������󶨵Ŀؼ��ľ���ͻ���0
    Set objCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    objCustom.Handle = patiFind.hwnd
    objCustom.flags = xtpFlagRightAlign
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("F"), conMenu_View_Filter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsAllowOperation(ByVal intOperationType As Byte, Optional ByRef strID_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�ļ�¼�Ƿ��������
    '���:intOperationType-0-�޸�;1-ɾ��;2-����;3-ȡ������;
    '����:true-����;False-������
    '����:���˺�
    '����:2018-11-09 11:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln�Ƿ�̶� As Boolean, bln���� As Boolean, bln�������� As Boolean
    On Error GoTo errHandle
    strID_Out = ""
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    strID_Out = rptData.SelectedRows(0).Record(COL_ID).Value
    If strID_Out = "" Then Exit Function
    
    '143491:���ϴ�,2019/8/7�������ϵͳ�̶��ġ���������Ϊ�����ɾ��
    bln�Ƿ�̶� = rptData.SelectedRows(0).Record(COL_�Ƿ�̶�).Value = "��"
    bln�������� = bln�Ƿ�̶� And rptData.SelectedRows(0).Record(COL_��Ϊ���).Value = "����"
    bln���� = rptData.SelectedRows(0).Record(COL_����ʱ��).Value <> ""
    Select Case intOperationType
    Case 0 '�޸�
        If bln�Ƿ�̶� And Not bln�������� Then Exit Function
        IsAllowOperation = Not bln����
    Case 1 'ɾ��
        If bln�Ƿ�̶� And Not bln�������� Then Exit Function
         IsAllowOperation = Not bln����
    Case 2 '����
         IsAllowOperation = Not bln����
    Case 3 'ȡ������
         IsAllowOperation = bln����
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnable As Boolean
    Dim blnStop As Boolean '�Ƿ���ͣ��
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next

    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�༭������¼")

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = rptData.Rows.Count > 0
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(0)   '�޸�
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(1)   'ɾ��
    Case conMenu_Edit_Stop
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(2)   '����
    Case conMenu_Edit_Reuse
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And IsAllowOperation(3)   'ȡ��������0-�޸�;1-ɾ��;2-����;3-ȡ������;
    Case conMenu_View_ShowStoped '��ʾ�ѳ����ļ�¼
        Control.Checked = mblnShowCancelRecord
    Case conMenu_View_Filter '����
    End Select
End Sub

Private Function ExecuteAddItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�����Ӳ�����¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    On Error GoTo errHandle
    'bytEditType-�༭���:0-����;1-�޸�;2-����;3-ȡ������;4-�鿴
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_����, mstr��Ϊ���) Then Exit Function
    
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    ExecuteAddItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteModifyItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���޸Ĳ�����¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    If IsAllowOperation(0, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_����ʱ��).Value) <> "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼�Ѿ�����,�������޸�!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
        If Trim(rptData.SelectedRows(0).Record(COL_�Ƿ�̶�).Value) <> "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼��ϵͳ�Զ����ɵģ���ֻ�ܳ���������¼,���������޸�!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    If strID = "" Then
        MsgBox "��ǰδѡ��Ҫɾ���Ĳ�����¼�����ܽ����޸Ĳ�����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'bytEditType-�༭���:0-����;1-�޸�;2-����;3-ȡ������;4-�鿴
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_�޸�, mstr��Ϊ���, strID) Then Exit Function

    
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    ExecuteModifyItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�й��˲���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmFilter As New frmBlackListRecordFilter
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    If Not frmFilter.zlShowEdit(mfrmMain, mlngModule, mcllFilter) Then Exit Function
    
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    ExecuteFilter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteStopItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�г���������¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    '0-�޸�;1-ɾ��;2-����;3-ȡ������;
    If IsAllowOperation(2, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_����ʱ��).Value) <> "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼�Ѿ�����,�������ٴγ���!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "��ǰδѡ��Ҫ�����Ĳ�����¼�����ܽ��г���������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'bytEditType-�༭���:0-����;1-�޸�;2-����;3-ȡ������;4-�鿴
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_����, mstr��Ϊ���, strID) Then Exit Function
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    
    ExecuteStopItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ExecuteCancelStopItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ȡ������������¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    '0-�޸�;1-ɾ��;2-����;3-ȡ������;
    If IsAllowOperation(2, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_����ʱ��).Value) = "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼δ������,������ȡ����������!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "��ǰδѡ��Ҫȡ�������Ĳ�����¼�����ܽ���ȡ������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'bytEditType-�༭���:0-����;1-�޸�;2-����;3-ȡ������;4-�鿴
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_ȡ������, mstr��Ϊ���, strID) Then Exit Function

    
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    ExecuteCancelStopItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteView() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ȡ������������¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-08 16:59:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmEdit As New frmBlackListEdit
    Dim strInfor As String, strID As String
    
    On Error GoTo errHandle
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    strID = rptData.SelectedRows(0).Record(COL_ID).Value
 
    If strID = "" Then
        MsgBox "��ǰδѡ��Ҫȡ�������Ĳ�����¼�����ܽ���ȡ������������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    'bytEditType-�༭���:0-����;1-�޸�;2-����;3-ȡ������;4-�鿴
    If Not frmEdit.zlShowEdit(mfrmMain, mlngModule, EM_RD_�鿴, mstr��Ϊ���, strID) Then Exit Function
    ExecuteView = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExcuteDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ɾ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 11:23:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strID As String, strSQL As String
    
    On Error GoTo errHandle
    If rptData.SelectedRows.Count <= 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Then Exit Function
    
    If IsAllowOperation(1, strID) = False Then
        If Trim(rptData.SelectedRows(0).Record(COL_����ʱ��).Value) <> "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼�Ѿ�����,������ɾ��!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
        If Trim(rptData.SelectedRows(0).Record(COL_�Ƿ�̶�).Value) <> "" Then
            strInfor = "���ˡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
            strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼��ϵͳ�Զ����ɵģ���ֻ�ܳ���������¼,��������ɾ��!"
            Call MsgBox(strInfor, vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If strID = "" Then
        MsgBox "��ǰδѡ��Ҫɾ���Ĳ�����¼�����ܽ���ɾ��������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If


    strInfor = "��" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "��"
    strInfor = strInfor & "�ķ���ʱ���ڡ�" & Trim(rptData.SelectedRows(0).Record(COL_��������).Value) & "�Ĳ�����¼��?"
    
    If MsgBox("��ȷ��Ҫɾ��" & strInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    strSQL = "Zl_���˲�����¼_Delete(" & strID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    
    ExcuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function



Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
     Err = 0: On Error GoTo errHandle
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem: Call ExecuteAddItem
    Case conMenu_Edit_Modify: Call ExecuteModifyItem
    Case conMenu_Edit_Delete: Call ExcuteDelete
    Case conMenu_Edit_Stop: Call ExecuteStopItem
    Case conMenu_Edit_Reuse: Call ExecuteCancelStopItem
    
    Case conMenu_View_ShowStoped '��ʾ�ѳ����Ĳ�����¼
        mblnShowCancelRecord = Not mblnShowCancelRecord
        Control.Checked = mblnShowCancelRecord
        Call zlDatabase.SetPara("��ʾ������¼", IIf(mblnShowCancelRecord, "1", "0"), glngSys, mlngModule)
        Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    Case conMenu_View_Refresh
        Call LoadRecordDataToGrid(mstr��Ϊ���, mcllFilter)
    Case conMenu_View_Find
         If patiFind.Visible And patiFind.Enabled Then patiFind.SetFocus
    Case conMenu_View_Filter: Call ExecuteFilter
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitRptColHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���б�
    '����:���˺�
    '����:2018-11-09 11:59:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objCol As ReportColumn, lngIdx As Long
    
    Err = 0: On Error GoTo errHandle
    
    With rptData
        .AutoColumnSizing = False '��ʹ���Զ��п�
        .AllowColumnRemove = False '�������϶�ɾ����
        .ShowGroupBox = True '��ʾ�����
        .ShowItemsInGroups = False '����ʾ�ѷ������
        .MultipleSelection = False '���������ѡ��
        .SetImageList Me.imgList16
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid '�������߸�ʽ
            .HorizontalGridStyle = xtpGridSolid '�������߸�ʽ
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ������..."
            .ShadeSortColor = .BackColor
            Set .CaptionFont = Me.Font
            Set .TextFont = Me.Font
        End With
    End With

    With rptData.Columns
        Set objCol = .Add(COL_ID, "ID", 50, True): objCol.Visible = False
        Set objCol = .Add(COL_ͼ��, "", 20, False)
        objCol.Groupable = False
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.AllowRemove = False
        
        Set objCol = .Add(COL_��Ϊ���, "��Ϊ���", 50, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_��������, "��������", 80, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����, "����", 50, True): objCol.Alignment = xtpAlignmentLeft
        Set objCol = .Add(COL_�Ա�, "�Ա�", 50, True): objCol.Alignment = xtpAlignmentLeft
        Set objCol = .Add(COL_��������, "��������", 80, True)
        Set objCol = .Add(COL_�����, "�����", 80, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True)
        
        Set objCol = .Add(COL_����ԭ��, "����ԭ��", 80, True)
        Set objCol = .Add(COL_������ϸ˵��, "������ϸ˵��", 200, True)
        
        
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True): objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Add(COL_������Ϣ, "������Ϣ", 50, True)
        Set objCol = .Add(COL_�Ǽ���, "�Ǽ���", 80, True)
        Set objCol = .Add(COL_������, "������", 80, True)
        Set objCol = .Add(COL_����ʱ��, "����ʱ��", 130, True)
        Set objCol = .Add(COL_����ԭ��, "����ԭ��", 100, True)
        Set objCol = .Add(COL_�Ƿ�̶�, "�Ƿ�ɾ��", 50, True): objCol.Visible = False
    End With
    
    With rptData
    '        '����Ϊ���ȱʡ��������
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(COL_����ʱ��)
        .SortOrder(0).SortAscending = True
        
        '����Ϊ����������������
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(COL_��Ϊ���)
'        .GroupsOrder(0).SortAscending = True
        .Columns(COL_��Ϊ���).Visible = False
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetBlackListRecords(ByVal strType As String, ByVal cllFilter As Collection, ByRef rsBlackLists_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������¼����
    '���:cllFilter-����(Array("����",ֵ )
    '    strType-��ǰ���
    '����:rsBlackLists_Out-���غ�������¼����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 12:06:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSQL As String, varData As Variant, i As Long
    Dim dt����ʱ��_Start  As Date, dt����ʱ��_End As Date
    Dim dt����ʱ��_Start  As Date, dt����ʱ��_End As Date
    Dim dt����ʱ��_Start  As Date, dt����ʱ��_End As Date
    Dim str����ԭ�� As String, str�Ǽ��� As String, str������ As String
    Dim lng����ID As Long
    Dim dtCurdate As Date
    
    strWhere = ""
    If strType <> "" Then strWhere = " And  A.��Ϊ���=[1]"
    If cllFilter Is Nothing Then
        Set cllFilter = New Collection
        dtCurdate = zlDatabase.Currentdate
        dt����ʱ��_End = Format(dtCurdate, "yyyy-mm-dd 23:59:59")
        dt����ʱ��_Start = Format(DateAdd("m", -6, dtCurdate), "yyyy-mm-dd 00:00:00") 'ȱʡ����
        strWhere = strWhere & " And ����ʱ�� Between [3] and [4]"
    Else
     
        For i = 1 To cllFilter.Count
            varData = cllFilter(i)
            
            Select Case varData(0)
            Case "����ID"
                lng����ID = Val(varData(1))
                strWhere = strWhere & " And A.����ID=[2]"
            Case "����ʱ��"
                dt����ʱ��_End = CDate(varData(2))
                dt����ʱ��_Start = CDate(varData(1))
                strWhere = strWhere & " And A.����ʱ�� Between [3] and [4]"
            Case "����ʱ��"
                dt����ʱ��_End = CDate(varData(2))
                dt����ʱ��_Start = CDate(varData(1))
                strWhere = strWhere & " And A.����ʱ�� Between [5] and [6]"
            Case "����ʱ��"
                dt����ʱ��_End = CDate(varData(2))
                dt����ʱ��_Start = CDate(varData(1))
                strWhere = strWhere & " And A.����ʱ�� Between [7] and [8]"
            Case "����ԭ��"
                str����ԭ�� = varData(1)
                strWhere = strWhere & " And A.����ԭ��=[9]"
            Case "�Ǽ���"
                str�Ǽ��� = varData(1)
                strWhere = strWhere & " And A.�Ǽ���=[10]"
            Case "������"
                str������ = varData(1)
                strWhere = strWhere & " And A.������=[11]"
            End Select
        Next
    End If
    If Not mblnShowCancelRecord Then strWhere = strWhere & "  And A.����ʱ�� is NULL"
    
    strSQL = "" & _
    " Select a.Id,a.��Ϊ���, a.����ID,b.���� as ��������,b.�Ա�,b.����,to_char(b.��������,'yyyy-mm-dd') as ��������," & _
    "       b.�����, to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��, a.����ԭ��  , a.����˵�� , " & _
    "       to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��  ," & vbNewLine & _
    "       a.������Ϣ, a.�Ǽ��� ,a.����ԭ��, a.������, to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,nvl(C.�Ƿ�̶�,0) as �Ƿ�̶�" & vbNewLine & _
    " From ���˲�����¼ A, ������Ϣ B,������Ϊ���� C" & vbNewLine & _
    " Where a.����ID+0 = b.����Id(+)  and a.��Ϊ���=C.����(+) " & vbNewLine & strWhere & _
    " Order by a.��Ϊ���,a.����ʱ��"
    
    Set mrs������¼ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strType, lng����ID, dt����ʱ��_Start, dt����ʱ��_End, dt����ʱ��_Start, _
        dt����ʱ��_End, dt����ʱ��_Start, dt����ʱ��_End, str����ԭ��, str�Ǽ���, str������)
        
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InsertRowData(ByVal strID As String, ByVal str��Ϊ��� As String, ByVal str�������� As String, ByVal str���� As String, _
    ByVal str�Ա� As String, ByVal str�������� As String, ByVal str����� As String, ByVal str����ʱ�� As String, ByVal str����ԭ�� As String, _
    ByVal str������ϸ˵�� As String, ByVal str����ʱ�� As String, ByVal str������Ϣ As String, _
    ByVal str�Ǽ��� As String, ByVal str������ As String, str����ʱ�� As String, str����ԭ�� As String, ByVal bln�Ƿ�̶� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���б��в���������
    '����:���˺�
    '����:2018-11-09 13:43:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strTemp As String
    Dim i As Long
    
    Err = 0: On Error GoTo errHandle
    With rptData
        
        Set objRecord = .Records.Add()
        
        Set objItem = objRecord.AddItem(strID)
        Set objItem = objRecord.AddItem("")
        objItem.Icon = IIf(bln�Ƿ�̶�, 1, 0) 'ͼ������
        
        Set objItem = objRecord.AddItem(str��Ϊ���)
        Set objItem = objRecord.AddItem(str��������)
        Set objItem = objRecord.AddItem(str����)
        Set objItem = objRecord.AddItem(str�Ա�)
        Set objItem = objRecord.AddItem(str��������)
        Set objItem = objRecord.AddItem(str�����)
        Set objItem = objRecord.AddItem(str����ʱ��)
        Set objItem = objRecord.AddItem(str����ԭ��)
        Set objItem = objRecord.AddItem(str������ϸ˵��)
        Set objItem = objRecord.AddItem(str����ʱ��)
        Set objItem = objRecord.AddItem(str������Ϣ)
        Set objItem = objRecord.AddItem(str�Ǽ���)
        Set objItem = objRecord.AddItem(str������)
        Set objItem = objRecord.AddItem(str����ʱ��)
        Set objItem = objRecord.AddItem(str����ԭ��)
        Set objItem = objRecord.AddItem(IIf(bln�Ƿ�̶�, "��", ""))
 
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function LoadRecordDataToGrid(ByVal str��Ϊ��� As String, ByVal cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����ؼ�
    '���:cllFilter-��Ҫ���˵�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 13:40:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim i As Long, j As Long, lngSelectRow As Long, strNewItem As String
    
    Err = 0: On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    
    If rptData.SelectedRows.Count > 0 Then lngSelectRow = rptData.SelectedRows(0).Index
    
    rptData.Records.DeleteAll
    
    If GetBlackListRecords(str��Ϊ���, cllFilter, mrs������¼) Then Exit Function
    With mrs������¼
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Call InsertRowData(Nvl(!ID), Nvl(!��Ϊ���), Nvl(!��������), Nvl(!����), _
                 Nvl(!�Ա�), Nvl(!��������), Nvl(!�����), Nvl(!����ʱ��), Nvl(!����ԭ��), Nvl(!����˵��), _
                Nvl(!����ʱ��), Nvl(!������Ϣ), Nvl(!�Ǽ���), _
                Nvl(!������), Nvl(!����ʱ��), Nvl(!����ԭ��), Val(Nvl(!�Ƿ�̶�)) = 1)
            .MoveNext
        Loop
    End With
    With rptData
        For i = 0 To .Records.Count - 1
            If i > .Records.Count - 1 Then Exit For
            If .Records(i).Item(COL_����ʱ��).Value <> "" Then
                For j = 0 To .Columns.Count - 1
                    .Records(i).Item(j).ForeColor = vbRed ' &H8000000C
                Next
            End If
        Next
    End With
    
    Call rptData.Populate '���������Ը��½���
    If rptData.Rows.Count > 0 Then '����ѡ������ʾ�ڿɼ�����
        If strNewItem <> "" Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow Then
                    If rptData.Rows(i).Record(COL_��������).Caption = strNewItem Then
                        rptData.FocusedRow = rptData.Rows(i)
                        Exit For
                    End If
                End If
            Next
        Else
            If lngSelectRow = 0 Then
                rptData.FocusedRow = rptData.Rows(0)
            ElseIf lngSelectRow > rptData.Rows.Count - 1 Then
                rptData.FocusedRow = rptData.Rows(rptData.Rows.Count - 1)
            Else
                rptData.FocusedRow = rptData.Rows(lngSelectRow)
            End If
        End If
    End If
    Call SetReportControlBackColorAlternate(rptData)
    RaiseEvent zlShowStatusText(2, "��ǰ����" & mrs������¼.RecordCount & "�����˲�����¼")
    Screen.MousePointer = vbDefault
    Exit Function
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picBack
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
End Sub

Private Sub patiFind_FindPatiArfter(ByVal objCard As zlOneCardComLib.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlOneCardComLib.clsPatientInfo, objCardData As zlOneCardComLib.clsPatientInfo, strErrMsg As String, blnCancel As Boolean)
    Dim cllFilter As Collection, lngPatiID As Long
    If objHisPati Is Nothing Then
        If patiFind.GetCurCard.���� Like "*��*��*" Then
            lngPatiID = GetPatient(ShowName)
        Else
            lngPatiID = 0
        End If
    Else
        lngPatiID = objHisPati.����ID
    End If
    
    Set cllFilter = New Collection
    cllFilter.Add Array("����ID", lngPatiID), "����ID"
    Call LoadRecordDataToGrid(mstr��Ϊ���, cllFilter)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.ActiveControl Is Nothing Then
        rptData.SetFocus
    ElseIf Not Me.ActiveControl Is patiFind Then
        rptData.SetFocus
    End If
    RaiseEvent zlActivate(Me)
End Sub

Private Sub InitFace()

    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2018-11-09 15:32:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFind As String, objCard As zlOneCardComLib.Card, i As Long
    Dim objCards As zlOneCardComLib.Cards, strKindstr As String, dtCurdate As Date
    Dim dt����ʱ��_End As Date, dt����ʱ��_Start As Date
    On Error GoTo errHandle
    
    mblnShowCancelRecord = Val(zlDatabase.GetPara("��ʾ������¼", glngSys, mlngModule, "0")) = 1
    Call InitRptColHead
    
    Set mcllFilter = New Collection
    dtCurdate = zlDatabase.Currentdate
    dt����ʱ��_End = Format(dtCurdate, "yyyy-mm-dd 23:59:59")
    dt����ʱ��_Start = Format(DateAdd("m", -6, dtCurdate), "yyyy-mm-dd 00:00:00") 'ȱʡ����
    mcllFilter.Add Array("����ʱ��", dt����ʱ��_Start, dt����ʱ��_End), "����ʱ��"
    
    strKindstr = "��|��������￨|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;ס|סԺ��|0;��|�ֻ���|0"
    Call patiFind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindstr, gstrProductName)
    patiFind.objIDKind.AllowAutoICCard = True
    patiFind.objIDKind.AllowAutoIDCard = True
    
    Set objCards = patiFind.objIDKind.Cards
    If Not objCards Is Nothing Then
        strFind = Val(zlDatabase.GetPara("�ϴβ������", glngSys, mlngModule, ""))  '����ȱʡ��
        If strFind <> "" Then
            For i = 1 To objCards.Count
                Set objCard = objCards(i)
                If objCard.���� = strFind Then
                    If patiFind.GetKindIndex(objCard.�ӿ����) >= 0 Then
                        patiFind.IDKindIDX = i + 1
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Load()
    mlngPreSelID = -1: Call InitFace
    RestoreWinState Me, App.ProductName
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName
    If Not patiFind.GetCurCard Is Nothing Then
        Call zlDatabase.SetPara("�ϴβ������", patiFind.GetCurCard.����, glngSys, mlngModule)
    End If
    If Not mrs������¼ Is Nothing Then Set mrs������¼ = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����
    If patiFind.Visible And patiFind.Enabled Then patiFind.ActiveFastKey
End Sub

Private Sub picBack_Resize()
    Err = 0: On Error Resume Next
    With picBack
        shpBorder.Move 0, 0, .ScaleWidth - 6, .ScaleHeight - 6
        stcTitle.Move .ScaleLeft, .ScaleTop, .ScaleWidth
        rptData.Left = .ScaleLeft + 10
        rptData.Top = stcTitle.Top + stcTitle.Height
        rptData.Width = .ScaleWidth - 30
        rptData.Height = .ScaleHeight - stcTitle.Height - 30
    End With
End Sub

Private Sub rptData_ColumnOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub

Private Sub rptData_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo errHandle
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    
    Me.SetFocus: RaiseEvent zlActivate(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim blnStop As Boolean, bln�Ƿ�̶� As Boolean, lngID As Long
    
    Err = 0: On Error GoTo errHandle
    If rptData.SelectedRows.Count = 0 Then Exit Sub
    If rptData.SelectedRows(0).GroupRow Then Exit Sub
    
    lngID = Val(rptData.SelectedRows(0).Record(COL_ID).Value)
    blnStop = rptData.SelectedRows(0).Record(COL_����ʱ��).Value <> ""
    bln�Ƿ�̶� = rptData.SelectedRows(0).Record(COL_�Ƿ�̶�).Value <> ""
    
    If lngID = 0 Then Exit Sub
    
    If zlStr.IsHavePrivs(mstrPrivs, "�༭������¼") And Not blnStop And Not bln�Ƿ�̶� Then
        Call ExecuteModifyItem  '�༭
    Else
        Call ExecuteView '�鿴
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim lngID As Long
    
    Err = 0: On Error GoTo errHandle
    lngID = 0
    If rptData.SelectedRows.Count <> 0 Then
        With rptData.SelectedRows(0)
            If Not .GroupRow Then
                lngID = Val(.Record(COL_ID).Value)
            End If
        End With
    End If
    If mlngPreSelID = lngID Then Exit Sub
    mlngPreSelID = lngID
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub rptData_SortOrderChanged()
    Call SetReportControlBackColorAlternate(rptData)
End Sub
Private Sub zlDataPrint(bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2018-11-09 15:57:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte, strHiddenCols As String
    
    Err = 0: On Error GoTo errHandle
    
    If UserInfo.���� = "" Then Call GetUserInfo
    
    '��ReportControlת��ΪVSFlexGrid
    strHiddenCols = CStr(COL_ID) & "," & CStr(COL_ͼ��) & "," & CStr(COL_�Ƿ�̶�)
    If zlGetVsfGrid(rptData, vsfGridPrint, strHiddenCols) = False Then Exit Sub
    
    objOut.Title.Text = "���˲�����¼�嵥"
    Set objOut.Body = vsfGridPrint
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "������Ϊ���" & IIf(mstr��Ϊ��� = "", "�������", mstr��Ϊ���)
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub stcTitle_GotFocus()
    On Error Resume Next
    If rptData.Visible Then rptData.SetFocus
End Sub
    
Private Function GetPatient(ByVal str���� As String) As Long
    Dim strCard As String, strPati As String
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If gblnShowCard Then
            strCard = "A.���￨�� as ���￨,A.���￨�� as ���￨��,"
        Else
            strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,A.���￨�� as ���￨��,"
        End If
        
        'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
        strPati = _
            " Select A.����ID ID,A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
            "   B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
            "   To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
            "   A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
            "   Nvl(a.��������,Decode(a.����,Null,'��ͨ����','ҽ������')) ��������" & _
            " From ������Ϣ A,���ű� B,���ű� C" & _
            " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & _
            " Order by A.����,A.�Ǽ�ʱ�� Desc"
        
        vRect = zlControl.GetControlRect(patiFind.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, patiFind.Height, blnCancel, False, True, str���� & "%")
        
        If rsTmp Is Nothing Then GetPatient = 0: Exit Function
        If blnCancel Then GetPatient = 0: Exit Function
        
        GetPatient = Val(Nvl(rsTmp!����ID))

        Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

