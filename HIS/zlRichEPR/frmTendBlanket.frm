VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmTendBlanket 
   Caption         =   "�����ص����"
   ClientHeight    =   7170
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   11085
   Icon            =   "frmTendBlanket.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11085
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6675
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4275
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   0
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   795
      Width           =   3510
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Top             =   1305
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6810
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendBlanket.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
   Begin MSComctlLib.ImageList ilsList 
      Left            =   6435
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendBlanket.frx":70E6
            Key             =   "K1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendBlanket.frx":D948
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendBlanket.frx":DAA2
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfPrint 
      Height          =   465
      Left            =   5520
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7965
      Top             =   2460
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
            Picture         =   "frmTendBlanket.frx":DBFC
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTendBlanket.frx":DD56
      Left            =   375
      Top             =   15
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendBlanket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngUpKey As Long
Private mblnOK As Boolean
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mblnDataChanged As Boolean
Private mlngTmp As Long
Private mstrPrivs As String
Private mblnNew As Boolean
Private mblnContiue As Boolean
Private WithEvents mfrmTendBlanketEdit As frmTendBlanketEdit
Attribute mfrmTendBlanketEdit.VB_VarHelpID = -1

'######################################################################################################################

Private Property Let DataChanged(ByVal blnData As Boolean)
    mfrmTendBlanketEdit.DataChanged = blnData

    If mfrmTendBlanketEdit.DataChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If

End Property

Private Property Get DataChanged() As Boolean
    If Not (mfrmTendBlanketEdit Is Nothing) Then
        DataChanged = mfrmTendBlanketEdit.DataChanged
    End If
End Property

Public Function ShowEdit(ByVal frmMain As Object, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mblnOK = False
    mstrPrivs = strPrivs
 

    If ExecuteCommand("��ʼ����") = False Then Exit Function

    Call ExecuteCommand("ˢ������")

    DataChanged = False

    Me.Show 1, frmMain

    ShowEdit = mblnOK

End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL

    Call mclsVsf.PrintData(bytMode, "�����ص�����嵥", msfPrint)

End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����

    Call CommandBarInit(cbsMain)
    cbsMain.Options.LargeIcons = True

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)

    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "���ӱ��(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "��������(&A)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ�����(&D)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&R)")

    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)


    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)

    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings

        .Add 0, vbKeyF5, conMenu_View_Refresh           'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help              '����

        .Add FCONTROL, vbKeyP, conMenu_File_Print       '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem     '����
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save '����
    End With

End Function


Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 400, DockLeftOf, Nothing)
    objPane.Title = "�嵥"
    objPane.Options = PaneNoCaption


    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing)
    objPane.Title = "��ϸ"
    objPane.Options = PaneNoCaption

    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)


End Sub

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim bytMode As Byte
    Dim intRow As Integer
    Dim objItem As Object

    On Error GoTo errHand

    Call SQLRecord(rsSQL)


    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        Set mclsVsf = New clsVsf

        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, False, ilsList)
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("�ص���Ŀ", 3600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��Ƿ���", 900, flexAlignCenterCenter, flexDTString, "", , True)
            Call .AppendColumn("������ɫ", 900, flexAlignCenterCenter, flexDTString, "", , True)
            Call .AppendColumn("�����ɫ", 0, flexAlignCenterCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���ͼ��", 900, flexAlignCenterCenter, flexDTString, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)

            .AppendRows = True
        End With

        Call InitDockPannel
        Call InitCommandBar
        Call RestoreWinState(Me, App.ProductName)

    '--------------------------------------------------------------------------------------------------------------
    Case "ˢ������"

        Call ExecuteCommand("��ȡ����")
        Call ExecuteCommand("ˢ��״̬")
        Call ExecuteCommand("��ȡ��������")

    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"

        mclsVsf.ClearGrid

        strSQL = "Select 'K1' As ͼ��,��� As ID,�ص���Ŀ,��Ƿ���,�����ɫ From �����ص���� Where �ϼ���� Is Null Order By ���"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rs.BOF = False Then

            Call mclsVsf.LoadGrid(rs)

            For intRow = 1 To vsf(0).Rows - 1
                Call ExecuteCommand("��ʾͼ��", intRow)
                Call ExecuteCommand("��ʾ��ɫ", intRow)
            Next

        End If

    '--------------------------------------------------------------------------------------------------------------
    Case "��ʾ��ɫ"

        With vsf(0)
            '������ɫ
            strTmp = .TextMatrix(Val(varParam(0)), .ColIndex("�����ɫ"))
            On Error Resume Next
            Set objItem = Nothing
            Set objItem = ils16.ListImages("K" & Val(strTmp))
            On Error GoTo errHand

            If objItem Is Nothing Then Call SetColorIcon(Me, "K" & Val(strTmp), Val(strTmp), ils16)
            Set .Cell(flexcpPicture, Val(varParam(0)), .ColIndex("������ɫ")) = ils16.ListImages("K" & Val(strTmp)).Picture
            .Cell(flexcpPictureAlignment, Val(varParam(0)), .ColIndex("������ɫ")) = flexAlignCenterCenter
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʾͼ��"

        With vsf(0)
            strTmp = zlBlobRead(9, Val(.RowData(Val(varParam(0)))))

            If Dir(strTmp) <> "" And strTmp <> "" Then
                
                picIcon.Cls
                Call DrawPicture(picIcon, strTmp, 0, 0, picIcon.Width, picIcon.Height)

                strTmp = CreateTmpFile
                Call SavePicture(picIcon.Image, strTmp)
                If Dir(strTmp) <> "" And strTmp <> "" Then

                    Set .Cell(flexcpPicture, Val(varParam(0)), .ColIndex("���ͼ��")) = VB.LoadPicture(strTmp)
                    .Cell(flexcpPictureAlignment, Val(varParam(0)), .ColIndex("���ͼ��")) = 4
                    Kill strTmp

                Else
                    Set .Cell(flexcpPicture, Val(varParam(0)), .ColIndex("���ͼ��")) = Nothing
                End If

            Else
                Set .Cell(flexcpPicture, Val(varParam(0)), .ColIndex("���ͼ��")) = Nothing
            End If
        End With

    '--------------------------------------------------------------------------------------------------------------
    Case "�������"

        mclsVsf.ClearGrid
        DataChanged = False

    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡ��������"

        Call mfrmTendBlanketEdit.RefreshData(Val(vsf(0).RowData(vsf(0).Row)))

    '--------------------------------------------------------------------------------------------------------------
    Case "��ȡָ������"

        strSQL = "Select 'K1' As ͼ��,��� As ID,�ص���Ŀ,��Ƿ���,�����ɫ From �����ص���� Where ���=[1]"

        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngTmp)
        If rs.BOF = True Then Exit Function

        With vsf(0)
            intRow = mclsVsf.FindRow(mlngTmp, -1)
            If intRow > 0 Then
                '�Ѽ���
                .Row = intRow
            Else
                'δ����
                If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                .Row = .Rows - 1
            End If

            Call mclsVsf.LoadGridRow(.Row, rs)

            Call ExecuteCommand("��ʾͼ��", .Row)

            Call ExecuteCommand("��ʾ��ɫ", .Row)

        End With

        Call ExecuteCommand("ˢ��״̬")
    '--------------------------------------------------------------------------------------------------------------
    Case "���ӱ��"

        mblnNew = True

        If Val(vsf(0).RowData(vsf(0).Rows - 1)) > 0 Then vsf(0).Rows = vsf(0).Rows + 1
        vsf(0).Row = vsf(0).Rows - 1
        vsf(0).ShowCell vsf(0).Row, vsf(0).Col

        Call mfrmTendBlanketEdit.NewData

        Exit Function

    '--------------------------------------------------------------------------------------------------------------
    Case "ɾ�����"

        If Val(vsf(0).RowData(vsf(0).Row)) = 0 Then Exit Function

        If MsgBox("���Ƿ����Ҫɾ����ǰ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSQL = "zl_�����ص����_Delete(" & Val(vsf(0).RowData(vsf(0).Row)) & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
            ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        End If
        Exit Function

    '--------------------------------------------------------------------------------------------------------------
    Case "�Ƴ����"
        If vsf(0).Rows > 2 Then
            vsf(0).RemoveItem vsf(0).Row
            mclsVsf.AppendRows = True
        Else
            mclsVsf.ClearGrid
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case "�ָ�����"

        '1.�ָ���������
        '----------------------------------------------------------------------------------------------------------
        If mfrmTendBlanketEdit.DataChanged Then
            If Val(vsf(0).RowData(vsf(0).Row)) = 0 And vsf(0).Rows > 2 Then
                vsf(0).Rows = vsf(0).Rows - 1
                vsf(0).Row = vsf(0).Rows - 1
            End If

            Call ExecuteCommand("��ȡ��������")
            mfrmTendBlanketEdit.DataChanged = False
        End If

        mblnNew = False
    '--------------------------------------------------------------------------------------------------------------
    Case "У������"

        '1.У����ϸ����
        '----------------------------------------------------------------------------------------------------------
        If mfrmTendBlanketEdit.DataChanged Then
            If mfrmTendBlanketEdit.ValidData = False Then Exit Function
        End If

        ExecuteCommand = True

        Exit Function
    '--------------------------------------------------------------------------------------------------------------
    Case "��������"

        mlngTmp = Val(vsf(0).RowData(vsf(0).Row))

        '1.������ϸ����
        '----------------------------------------------------------------------------------------------------------
        If mfrmTendBlanketEdit.DataChanged Then

            If mfrmTendBlanketEdit.SaveData(rsSQL, mlngTmp) = False Then Exit Function

        End If

        If SQLRecordExecute(rsSQL, Me.Caption) Then
            ExecuteCommand = True

            '����ͼƬ
            If SQLRecordSavePicture(rsSQL, Me.Caption) Then

            End If

        End If

        Exit Function

    End Select

    ExecuteCommand = True

    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem               '����

        mblnContiue = False
        Call ExecuteCommand("���ӱ��")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append               '��������

        mblnContiue = True
        Call ExecuteCommand("���ӱ��")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                'ɾ��

        If ExecuteCommand("ɾ�����") Then
            Call ExecuteCommand("�Ƴ����")
        End If


    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save                  '��������

        If ExecuteCommand("У������") And DataChanged Then
            If ExecuteCommand("��������") Then

                DataChanged = False

                Call ExecuteCommand("��ȡָ������")

                If mblnContiue Then
                    Call ExecuteCommand("���ӱ��")
                End If

            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle                  '�ָ�����

        If ExecuteCommand("�ָ�����") Then

        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                            '���
        mclsVsf.ClearGrid
        DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call zlRptPrint(2)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call zlRptPrint(1)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call zlRptPrint(3)

    '--------------------------------------------------------------------------------------------------------------
    Case Else

         '��ҵ���޹صĹ��ܣ������Ĺ���
        Call CommandBarExecutePublic(Control, Me)

    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup                                  '�༭���˵�

        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (DataChanged = False And Control.Visible)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (Val(vsf(0).RowData(vsf(0).Row)) > 0 And DataChanged = False And Control.Visible)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle

        Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = (DataChanged And Control.Visible)

    '--------------------------------------------------------------------------------------------------------------
    Case Else

         '��ҵ���޹صĹ��ܣ������Ĺ���
         Call CommandBarUpdatePublic(Control, Me)

    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmTendBlanketEdit = New frmTendBlanketEdit
        Call mfrmTendBlanketEdit.InitData(Me, IsPrivs(mstrPrivs, "��ɾ��"))

        Item.Handle = mfrmTendBlanketEdit.hWnd
    End Select
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Call SetPaneRange(dkpMain, 2, 230, 15, 230, Me.ScaleHeight)

    dkpMain.RecalcLayout

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveWinState(Me, App.ProductName)

    Set mclsVsf = Nothing

    Unload mfrmTendBlanketEdit

End Sub

Private Sub mclsVsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(0).RowData(Row)) <= 0)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf.AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case Index
    Case 0
        If OldRow = NewRow Then Exit Sub
        Call ExecuteCommand("��ȡ��������")
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Call mclsVsf.RestoreRow(mclsVsf.SaveKey)
    vsf(Index).ShowCell vsf(Index).Row, vsf(Index).Col
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    mclsVsf.SaveKey = Val(vsf(Index).RowData(vsf(Index).Row))
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = mclsVsf.ColIndex("ͼ��"))
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    mclsVsf.MoveColumn = (vsf(Index).MouseRow = 0)

End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object

    If Button <> 2 Then Exit Sub

    If cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup

End Sub


