VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMipComView 
   Caption         =   "ͨ����Ϣ����"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13410
   Icon            =   "frmMipComView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   13410
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4245
      Index           =   1
      Left            =   240
      ScaleHeight     =   4245
      ScaleWidth      =   2370
      TabIndex        =   10
      Top             =   1695
      Width           =   2370
      Begin XtremeSuiteControls.TaskPanel tpl 
         Height          =   4770
         Left            =   345
         TabIndex        =   11
         Top             =   495
         Width           =   3210
         _Version        =   589884
         _ExtentX        =   5662
         _ExtentY        =   8414
         _StockProps     =   64
         Behaviour       =   1
         ItemLayout      =   2
         HotTrackStyle   =   3
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3885
      Index           =   0
      Left            =   5505
      ScaleHeight     =   3885
      ScaleWidth      =   5070
      TabIndex        =   3
      Top             =   3900
      Width           =   5070
      Begin RichTextLib.RichTextBox txtText 
         Height          =   1515
         Left            =   15
         TabIndex        =   9
         Top             =   945
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   2672
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMipComView.frx":0A02
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   915
         Index           =   1
         Left            =   15
         ScaleHeight     =   915
         ScaleWidth      =   5040
         TabIndex        =   4
         Top             =   15
         Width           =   5040
         Begin VB.Label lblLinkTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҩƷ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   915
            MouseIcon       =   "frmMipComView.frx":0A9F
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   630
            Width           =   1080
         End
         Begin VB.Label lblLinkType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    �ӣ�"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   45
            TabIndex        =   12
            Top             =   630
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Index           =   1
            Left            =   915
            TabIndex        =   8
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    �⣺"
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ�䣺"
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   6
            Top             =   90
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2014-01-10 16:58:00"
            Height          =   180
            Index           =   2
            Left            =   915
            TabIndex        =   5
            Top             =   90
            Width           =   1710
         End
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7425
      TabIndex        =   2
      Top             =   585
      Width           =   1575
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2010
      Index           =   2
      Left            =   5535
      ScaleHeight     =   2010
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   1680
      Width           =   2700
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1785
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   2670
         _cx             =   4710
         _cy             =   3149
         Appearance      =   0
         BorderStyle     =   0
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   3285
      Top             =   2445
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMipComView.frx":0DA9
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   210
      Top             =   750
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMipComView.frx":20CF
      Left            =   375
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMipComView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'��������

Private Enum Command
    ��ʼ�ؼ�
    ��ע���
    ɾ����Ϣ
    �Ķ���Ϣ
    ����Ķ�
    ˢ������
    ˢ����Ϣ
End Enum

Private mlngModualCode As Long
Private mstrSQL As String
Private mclsVsf(0) As zlVSFlexGrid.clsVsf
Private mblnStartUp As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mstrDataFile As String
Private mclsMipReceiptData As clsMipReceiptData
Private mstrCurrentGroup As String
Private mobjParentForm As Object

Public Event OpenLink(ByVal bytLinkType As Byte, ByVal strLinkPara As String)
Public Event AfterReadMessage()

'######################################################################################################################
'�ӿڷ���
Public Function ShowForm(ByVal objParentForm As Object, ByVal strDataFile As String, Optional ByVal blnOnlyNew As Boolean = False)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrDataFile = strDataFile
    
    mstrCurrentGroup = "G1"
    If blnOnlyNew Then
        mstrCurrentGroup = "G2"
        If tpl.Groups.Count > 0 Then
            tpl.Groups(1).Items(1).Selected = False
            tpl.Groups(1).Items(2).Selected = True
        End If
    End If
    
    Set mobjParentForm = objParentForm
    Me.Show , mobjParentForm
    
    Call ExecuteCommand(Command.ˢ����Ϣ)
    Call ExecuteCommand(Command.ˢ������)
    
End Function

'######################################################################################################################
Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, True, gfrmMipResource.ils16)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[���]", False)
        Call .AppendColumn("", 300, flexAlignCenterCenter, flexDTBoolean, "", "[ѡ��]", False)
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("�Ķ����", 0, flexAlignLeftCenter, flexDTString, , "receive_read", True, , , True)
        Call .AppendColumn("�ı�����", 0, flexAlignLeftCenter, flexDTString, , "receive_text", True, , , True)
        Call .AppendColumn("��������", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_type", True, , , True)
        Call .AppendColumn("���ӱ���", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_title", True, , , True)
        Call .AppendColumn("���Ӳ���", 0, flexAlignLeftCenter, flexDTString, , "receive_lnk_para", True, , , True)
        
        Call .AppendColumn("ʱ��", 1800, flexAlignLeftCenter, flexDTString, , "receive_date", True)
        Call .AppendColumn("����", 1800, flexAlignLeftCenter, flexDTString, , "receive_topic", True)
                        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
        
    End With
        
    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExecuteCommand(ByVal enmCommand As Command, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim rsTmp As zlDataSQLite.SQLiteRecordset
    Dim rsCondition As ADODB.Recordset
    Dim strTmp As String
    Dim intRow As Integer
    Dim varTmp As Variant
    Dim blnMuliSelect As Boolean
    Dim strTemp As String
    Dim lngCount As Long
    Dim lngLoop As Long
    Dim strLine As String
    
    On Error GoTo errHand
    
    Select Case enmCommand
    '------------------------------------------------------------------------------------------------------------------
    Case Command.��ʼ�ؼ�
                
        Call InitGrid
        Call InitCommandBar
        Call InitDockPannel
        Call InitTaskPanel
        
        Set mclsMipReceiptData = New clsMipReceiptData
        Call mclsMipReceiptData.Initialize(mstrDataFile)
                
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ɾ����Ϣ
        With vsf(0)
                        
            blnMuliSelect = False
            For intRow = 1 To .Rows - 1
                If Val(Abs(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                    blnMuliSelect = True
                    Exit For
                End If
            Next
            
            
            If mclsMipReceiptData.OpenDataFile = True Then
                If blnMuliSelect = True Then
                    If MsgBox("��ȷ��Ҫɾ���Ѿ���ѡ�Ľ�����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        For lngLoop = 1 To .Rows - 1
                            If Abs(Val(.TextMatrix(lngLoop, .ColIndex("ѡ��")))) = 1 And .TextMatrix(lngLoop, .ColIndex("id")) <> "" Then
                                Call mclsMipReceiptData.DeleteReceiveMessage(.TextMatrix(lngLoop, .ColIndex("id")))
                            End If
                        Next
                        Call ExecuteCommand(Command.ˢ����Ϣ)
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                    If MsgBox("��ȷ��Ҫɾ����ǰѡ���еĽ�����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If mclsMipReceiptData.DeleteReceiveMessage(.TextMatrix(.Row, .ColIndex("id"))) Then
                            Call ExecuteCommand(Command.ˢ����Ϣ)
                        End If
                    End If
                End If
                mclsMipReceiptData.CloseDataFile
            End If
            
            
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.�Ķ���Ϣ
        With vsf(0)
            If mclsMipReceiptData.OpenDataFile = True Then
                blnMuliSelect = False
                For intRow = 1 To .Rows - 1
                    If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 Then
                        blnMuliSelect = True
                        Exit For
                    End If
                Next
                
                If blnMuliSelect = True Then
                    If MsgBox("��ȷ��Ҫ����Щ��Ϣ�����Ϊ���Ķ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        For lngLoop = 1 To .Rows - 1
                            If Abs(Val(.TextMatrix(intRow, .ColIndex("ѡ��")))) = 1 And .TextMatrix(intRow, .ColIndex("id")) <> "" And .TextMatrix(intRow, .ColIndex("�Ķ����")) <> "1" Then
                                If mclsMipReceiptData.UpdateReceiveMessageReaded(.TextMatrix(intRow, .ColIndex("id"))) Then
                                    .TextMatrix(intRow, .ColIndex("�Ķ����")) = "1"
                                    .Cell(flexcpFontBold, intRow, 1, intRow, .Cols - 1) = False
                                End If
                            End If
                        Next
                        RaiseEvent AfterReadMessage
                    End If
                ElseIf .TextMatrix(.Row, .ColIndex("id")) <> "" And .TextMatrix(.Row, .ColIndex("�Ķ����")) <> "1" Then
                    
                    If mclsMipReceiptData.UpdateReceiveMessageReaded(.TextMatrix(.Row, .ColIndex("id"))) Then
                        .TextMatrix(.Row, .ColIndex("�Ķ����")) = "1"
                        .Cell(flexcpFontBold, .Row, 1, .Row, .Cols - 1) = False
                        RaiseEvent AfterReadMessage
                    End If
                    
                End If
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ����Ϣ
        
        With vsf(0)
            
            If mclsMipReceiptData.OpenDataFile() = True Then
            
                mclsVsf(0).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
                mclsVsf(0).ClearGrid
                Set rsCondition = zlCommFun.CreateCondition
                
                
                rsTmp = mclsMipReceiptData.ReadReceiveMessage("Count", rsCondition)
                If rsTmp.DataSet.BOF = False Then
                    If Val(rsTmp.DataSet("δ����").Value) > 0 Then
                        tpl.Groups(1).Items(2).Caption = "δ����Ϣ[" & Val(rsTmp.DataSet("δ����").Value) & "]"
                        tpl.Groups(1).Items(2).Bold = True
                    Else
                        tpl.Groups(1).Items(2).Caption = "δ����Ϣ"
                        tpl.Groups(1).Items(2).Bold = False
                    End If
                    tpl.Reposition
                End If
                
'                Call zlCommFun.SetCondition(rsCondition, "Start_Date", Format(dtp(0).Value, dtp(0).CustomFormat))
'                Call zlCommFun.SetCondition(rsCondition, "End_Date", Format(dtp(1).Value, dtp(1).CustomFormat))

                
                Call zlCommFun.SetCondition(rsCondition, "receive_read", IIf(mstrCurrentGroup = "G2", 1, 0))
                
                If Trim(txtLocation.Text) = "" Then
                    
                    rsTmp = mclsMipReceiptData.ReadReceiveMessage("FilterData", rsCondition)
                    If rsTmp.DataSet.BOF = False Then ExecuteCommand = mclsVsf(0).LoadDataSource(rsTmp.DataSet.DataSource)

                Else
                    Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
                    Call zlCommFun.SetCondition(rsCondition, "FilterText", Trim(txtLocation.Text))
                    
                    rsTmp = mclsMipReceiptData.ReadReceiveMessage("FilterData", rsCondition)
                    If rsTmp.DataSet.BOF = False Then ExecuteCommand = mclsVsf(0).LoadDataSource(rsTmp.DataSet.DataSource)
                    
                End If
    
                Call mclsVsf(0).RestoreRow(mclsVsf(0).SaveKey, .ColIndex("id"))
                
                For intRow = 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, .ColIndex("�Ķ����"))) = 0 Then
                        .Cell(flexcpFontBold, intRow, 1, intRow, .Cols - 1) = True
                    End If
                Next
                
                mclsMipReceiptData.CloseDataFile
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case Command.ˢ������
        With vsf(0)
            txtText.Text = ""
            lblLinkType.Tag = ""
            lblLinkTitle.Caption = ""
            lblLinkTitle.Tag = ""
            txtText.Text = .TextMatrix(.Row, .ColIndex("�ı�����"))
            lbl(2).Caption = .TextMatrix(.Row, .ColIndex("ʱ��"))
            lbl(1).Caption = .TextMatrix(.Row, .ColIndex("����"))
            lblLinkType.Tag = .TextMatrix(.Row, .ColIndex("��������"))
            lblLinkTitle.Caption = .TextMatrix(.Row, .ColIndex("���ӱ���"))
            lblLinkTitle.Tag = .TextMatrix(.Row, .ColIndex("���Ӳ���"))
            lblLinkTitle.Visible = (lblLinkTitle.Caption <> "")
            
            '���Ϊδ�Ķ�������Ϊ���Ķ�
            Call ExecuteCommand(Command.�Ķ���Ϣ, 0)

    
        End With
    End Select
        
    GoTo EndHand

    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mclsMipReceiptData.CloseDataFile
    '------------------------------------------------------------------------------------------------------------------
EndHand:
End Function


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
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call zlCommFun.CommandBarInit(cbsMain)
'    cbsMain.VisualTheme = xtpThemeNativeWinXP
    Set cbsMain.Icons = frmMipResource.imgPublic.Icons
    cbsMain.Options.LargeIcons = True
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap


    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True, , , "�˳�������־���Ĺ���")
    
    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ(&A)", , , , "����ǰ�б��е�����������Ϊ��ѡ״̬")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ClsAll, "ȫ��(&C)", , , , "����ǰ�б��е�����������Ϊ�ǹ�ѡ״̬")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "���(&D)", True, , , "�����ǰ�л��߹�ѡ�е�ͨ����Ϣ")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "���Ϊ�Ѷ�(&M)", True, , , "����ǰ�л��߹�ѡ�е�ͨ����Ϣ����Ϊ�Ѷ�״̬")
    
    
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = zlCommFun.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", , , , "��ʾ/���ع�������ť")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", , , , "��ʾ/���ع�������ť�ϵ���������")
    Set objControl = zlCommFun.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", , , , "���ù�������ťͼ��Ϊ��ͼ���Сͼ��")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)", , , , "��ʾ/����״̬��")
    
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True, , , "����ǰ���õ���������ˢ��ͨ����Ϣ����")
    
    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)", , , , "��ʾ����ͨ����Ϣ���ĵĲ���˵��")
    Set objControl = zlCommFun.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True, , , "��ʾ�й�ͨ����Ϣ�����˵��")
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        
            
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", True, , , , , "����ǰ�б��е�����������Ϊ��ѡ״̬")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "ȫ��", , , , , , "����ǰ�б��е�����������Ϊ�ǹ�ѡ״̬")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "���", True, , , , , "�����ǰ�л��߹�ѡ�е�ͨ����Ϣ")
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "���Ϊ�Ѷ�", True, , , , , "����ǰ�л��߹�ѡ�е�ͨ����Ϣ����Ϊ�Ѷ�״̬")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlLabel, conMenu_View_LocationItem, "����", False, , xtpButtonIconAndCaption)
    objControl.IconId = conMenu_View_Find
        
    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Refresh, "ˢ��", , , , , , "����ǰ���õ���������ˢ��ͨ����Ϣ����")
            
    cbsMain.StatusBar.Visible = True
    cbsMain.StatusBar.IdleText = "׼��"
    Call cbsMain.StatusBar.AddPane(0)
    Call cbsMain.StatusBar.SetPaneText(0, cbsMain.StatusBar.IdleText)
    Call cbsMain.StatusBar.SetPaneStyle(0, SBPS_STRETCH)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_CAPS)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_NUM)
    Call cbsMain.StatusBar.AddPane(ID_INDICATOR_SCRL)

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add FCONTROL, vbKeyDelete, conMenu_Edit_Delete     '���
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll          'ȫѡ
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll       'ȫ��
    End With
        
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, objPane)
    objPane.Title = "��¼"
    objPane.Options = PaneNoCaption
        
    Set objPane = dkpMain.CreatePane(3, 800, 100, DockBottomOf, objPane)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call zlCommFun.DockPannelInit(dkpMain)

End Sub

Private Sub InitTaskPanel()
    
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    
    With tpl
        .SetIconSize 24, 24
        Call .Icons.AddIcons(ImageManager1.Icons)
        .VisualTheme = xtpTaskPanelThemeNativeWinXP
        .Behaviour = xtpTaskPanelBehaviourToolbox
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        
        .SetMargins 5, 5, 5, 5, 5
        .SetItemInnerMargins 0, 5, 0, 5
        .SelectItemOnFocus = True
                        
        Set objGroup = .Groups.Add(0, "����")
        objGroup.Expandable = False
        objGroup.CaptionVisible = False
        
        Set objItem = objGroup.Items.Add(1, "������Ϣ", xtpTaskItemTypeLink, 3)
        objItem.Tag = "G1"
        objItem.Tooltip = "��ǰ����վ�����յ�����Ϣ"
        If mstrCurrentGroup = objItem.Tag Then objItem.Selected = True
                
        Set objItem = objGroup.Items.Add(2, "δ����Ϣ", xtpTaskItemTypeLink, 2)
        objItem.Tag = "G2"
        objItem.Tooltip = "��ǰ����վ�����յ���δ����Ϣ"
        If mstrCurrentGroup = objItem.Tag Then objItem.Selected = True
        
        .Reposition
    
    End With
    
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As Object
    Dim blnMuliSelect As Boolean
    
    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_SelAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 1
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ClsAll
        
        With vsf(0)
            .Cell(flexcpText, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
    
        Call ExecuteCommand(Command.ɾ����Ϣ)
                        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        
        blnMuliSelect = False
        With vsf(0)
            For lngLoop = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("ѡ��")))) = 1 Then
                    blnMuliSelect = True
                    Call ExecuteCommand(Command.�Ķ���Ϣ, 1)
                    Exit For
                End If
            Next
        End With
        If blnMuliSelect = False Then Call ExecuteCommand(Command.�Ķ���Ϣ, 0)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh               'ˢ��
                
        Call ExecuteCommand(Command.ˢ����Ϣ)
        Call ExecuteCommand(Command.ˢ������)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_StatusBar
        cbsMain.StatusBar.Visible = Not cbsMain.StatusBar.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Close
    
        Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
        Select Case Control.id
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
    
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            
            Control.Enabled = (Trim(.TextMatrix(.Row, .ColIndex("id"))) <> "" And Control.Visible)
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button            '������
            If cbsMain.Count >= 2 Then
                Control.Checked = cbsMain(2).Visible
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Text              'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Size              '��ͼ��
            Control.Checked = cbsMain.Options.LargeIcons
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_View_StatusBar                 '״̬��
            Control.Checked = cbsMain.StatusBar.Visible
        
        End Select
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.Handle = picPane(1).hWnd
    Case 2
        Item.Handle = picPane(2).hWnd
    Case 3
        Item.Handle = picPane(0).hWnd
    End Select
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    mlngModualCode = 1005
    
    Call ExecuteCommand(Command.��ʼ�ؼ�)
    Call ExecuteCommand(Command.��ע���)

'    If Not (gcnOracle Is Nothing) Then Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call zlCommFun.SetPaneRange(dkpMain, 1, 100, 15, 100, Me.ScaleHeight)
    Call zlCommFun.SetPaneRange(dkpMain, 3, 15, 100, Me.ScaleWidth, 200)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
    Set mobjFindKey = Nothing
            
End Sub

Private Sub lblLinkTitle_Click()
    
    RaiseEvent OpenLink(Val(lblLinkType.Tag), lblLinkTitle.Tag)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        picBack(1).Move 15, 15, picPane(Index).Width - 30
        txtText.Move 15, picBack(1).Top + picBack(1).Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (picBack(1).Top + picBack(1).Height + 15) - 15
    Case 1
        tpl.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 2
        vsf(0).Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Private Sub tpl_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    mstrCurrentGroup = Item.Tag
    Call ExecuteCommand(Command.ˢ����Ϣ)
    Call ExecuteCommand(Command.ˢ������)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        txtLocation.Tag = ""
        
        Dim obj As CommandBarControl
        
        Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
        If Not (obj Is Nothing) Then
            If obj.Enabled = True Then Call cbsMain_Execute(obj)
        End If
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf(Index).AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    If OldRow <> NewRow Then
        '���Ϊδ�Ķ�������Ϊ���Ķ�
        Call ExecuteCommand(Command.ˢ������)
    End If
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_AfterSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        Call mclsVsf(Index).RestoreRow(mclsVsf(Index).SaveKey, .ColIndex("id"))
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf(Index).AppendRows = True
End Sub

Private Sub vsf_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    With vsf(Index)
        mclsVsf(Index).SaveKey = Trim(.TextMatrix(.Row, .ColIndex("id")))
    End With
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   Call mclsVsf(Index).BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Dim objMenu As CommandBarControl
    
    Set objMenu = cbsMain.FindControl(, conMenu_Edit_Modify, False)
    If Not (objMenu Is Nothing) Then
        If objMenu.Enabled = True Then
            Call cbsMain_Execute(objMenu)
        End If
    End If
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf(Index).KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf(Index).KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mclsVsf(Index).MoveColumn = (vsf(Index).MouseRow = 0)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call zlCommFun.SendLMouseButton(vsf(Index).hWnd, X, Y)
        Select Case Index
        Case 0
            If mclsVsf(Index).MoveColumn = False Then
                Call ShowConetneMenu(1).ShowPopup
            End If
        End Select
        
    End Select
End Sub

Public Function ShowConetneMenu(Optional ByVal bytPlace As Byte = 1) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    
    Select Case bytPlace
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_SelAll, "ȫ����ѡ(&A)")
        cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ����ѡ(&U)")
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "�����Ϣ(&D)")
        cbrPopupItem.BeginGroup = True
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "����ˢ��(&R)")
        cbrPopupItem.BeginGroup = True
                
    End Select
    
    Set ShowConetneMenu = cbrPopupBar
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf(0).EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf(0).ValidateEdit(Col, Cancel)
End Sub



