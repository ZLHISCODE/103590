VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISAuditSafeKeep 
   Caption         =   "��������¼"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmCISAuditSafeKeep.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   11670
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5835
      Index           =   2
      Left            =   45
      ScaleHeight     =   5835
      ScaleWidth      =   8100
      TabIndex        =   0
      Top             =   465
      Width           =   8100
      Begin VB.PictureBox picPane 
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   0
         Left            =   45
         ScaleHeight     =   840
         ScaleWidth      =   7860
         TabIndex        =   1
         Top             =   4890
         Width           =   7860
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   1065
            TabIndex        =   9
            Top             =   30
            Width           =   6660
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "�˳�(&Q)"
            Height          =   350
            Left            =   6615
            TabIndex        =   7
            Top             =   345
            Width           =   1100
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "��ѯ(&F)"
            Height          =   350
            Left            =   5475
            TabIndex        =   6
            Top             =   345
            Width           =   1100
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1065
            TabIndex        =   2
            Top             =   390
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   487194627
            CurrentDate     =   38083
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   3345
            TabIndex        =   3
            Top             =   390
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   487194627
            CurrentDate     =   38083
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�������(&1)"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   75
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "���ʱ��(&2)"
            Height          =   180
            Index           =   8
            Left            =   0
            TabIndex        =   5
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   9
            Left            =   3180
            TabIndex        =   4
            Top             =   435
            Width           =   180
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1200
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmCISAuditSafeKeep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mlngMoual As Long

Private mrsCondition    As ADODB.Recordset
Private mclsVsf(0)      As clsVsf

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object, ByVal lngMoual As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngMoual = lngMoual
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Or ExecuteCommand("ˢ������") = False Then Exit Function
    
End Function


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call RptPrint(2)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call RptPrint(1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call RptPrint(3)
        
    End Select
    
End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With vfgThis
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               'Ԥ��,��ӡ,�����Excel
        
            Control.Enabled = (.Rows > .FixedRows + 1)
        
        End Select
        
    End With
    
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
    Dim objExtendedBar As CommandBar

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsThis)
    Set cbsThis.Icons = frmPubResource.imgApp.Icons
    cbsThis.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = True
    
    '�ļ�
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    Call CreateHelpMenu(cbsThis)
    
    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
    With cbsThis.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
    End With
    
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = vfgThis
    objPrint.Title.Text = "��������¼�嵥"
    
    Set objPrint.Title.Font = vfgThis.Font

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
                
        Call InitCommandBar
        
        Set mclsVsf(0) = New clsVsf
        With mclsVsf(0)
                    Call .Initialize(Me.Controls, vfgThis, True, True, frmPubResource.GetImageList(16))
                    Call .ClearColumn
                    Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("����״ֵ̬", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("����ת��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("���￨��", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                    Call .AppendColumn("����", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
'                    Call .AppendColumn("���ʱ��", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True, , , True)
                    
                    Call .AppendColumn("", 240, flexAlignCenterCenter, flexDTBoolean, , "[ѡ��]", False)
                    Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[ͼ��]", False)
                    Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[·��]", False)
                    Call .AppendColumn("����", 810, flexAlignLeftCenter, flexDTString, , , True)
    
                    Call .AppendColumn("סԺ��", 900, flexAlignLeftCenter, flexDTDecimal, , , True)
                    
                    Call .AppendColumn("����", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("����ȼ�", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("סԺҽʦ", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                    Call .AppendColumn("����״̬", 1080, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("��Ժ����", 1080, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("���״̬", 840, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("�ύ��", 810, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("�ύʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("�����", 990, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("���ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                    Call .AppendColumn("�������", 2000, flexAlignLeftCenter, flexDTString, , , True)
                    Call .AppendColumn("��Ժ����ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
              
                    .SysHidden(.ColIndex("ID")) = True
                    .SysHidden(.ColIndex("����id")) = True
                    .SysHidden(.ColIndex("��ҳid")) = True
                    .SysHidden(.ColIndex("����״ֵ̬")) = True
                    .SysHidden(.ColIndex("���￨��")) = True
                    .SysHidden(.ColIndex("����")) = True
'                    .SysHidden(.ColIndex("���ʱ��")) = True
                    .SysHidden(.ColIndex("����ת��")) = True
                    .SysHidden(.ColIndex("��Ժ����ID")) = True
                    
                    Call .InitializeEdit(True, False, False)
                    Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
            .AppendRows = True
        End With
        DoEvents
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        dtp(0).Value = Format(DateAdd("d", -7, Now()), "YYYY-MM-DD 00:00:00")
        dtp(1).Value = Format(Now(), "YYYY-MM-DD 23:59:59")
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
    
        mclsVsf(0).ClearGrid
        
        Set rs = gclsPackage.GetAduitPatientSafeKeep(txt����.Text, CStr(dtp(0).Value), CStr(dtp(1).Value))
        If rs.BOF = False Then
            Call mclsVsf(0).LoadDataSource(rs)
'            rs.MoveFirst
'            Do Until rs.EOF
'                With vfgThis
'                    .Rows = .Rows + 1
'                    .TextMatrix(.Rows - 2, .ColIndex("ID")) = rs!ID
'                    .TextMatrix(.Rows - 2, .ColIndex("����id")) = rs!����ID
'                    .TextMatrix(.Rows - 2, .ColIndex("��ҳid")) = rs!��ҳID
'                    .TextMatrix(.Rows - 2, .ColIndex("����״ֵ̬")) = zlCommFun.NVL(rs!����״ֵ̬, 0)
'                    .TextMatrix(.Rows - 2, .ColIndex("����ת��")) = zlCommFun.NVL(rs!����ת��)
'                    .TextMatrix(.Rows - 2, .ColIndex("���￨��")) = zlCommFun.NVL(rs!���￨��)
'                    .TextMatrix(.Rows - 2, .ColIndex("����")) = zlCommFun.NVL(rs!����)
'                    .TextMatrix(.Rows - 2, .ColIndex("���ʱ��")) = zlCommFun.NVL(rs!���ʱ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("����")) = zlCommFun.NVL(rs!����)
'                    .TextMatrix(.Rows - 2, .ColIndex("סԺ��")) = zlCommFun.NVL(rs!סԺ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("����")) = zlCommFun.NVL(rs!����)
'                    .TextMatrix(.Rows - 2, .ColIndex("����ȼ�")) = zlCommFun.NVL(rs!����ȼ�)
'                    .TextMatrix(.Rows - 2, .ColIndex("סԺҽʦ")) = zlCommFun.NVL(rs!סԺҽʦ)
'                    .TextMatrix(.Rows - 2, .ColIndex("��Ժ����")) = zlCommFun.NVL(rs!��Ժ����)
'                    .TextMatrix(.Rows - 2, .ColIndex("���״̬")) = zlCommFun.NVL(rs!���״̬)
'
'                    .TextMatrix(.Rows - 2, .ColIndex("�ύ��")) = zlCommFun.NVL(rs!�ύ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("�ύʱ��")) = zlCommFun.NVL(rs!�ύʱ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("������")) = zlCommFun.NVL(rs!������)
'                    .TextMatrix(.Rows - 2, .ColIndex("����ʱ��")) = zlCommFun.NVL(rs!����ʱ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("���ʱ��")) = zlCommFun.NVL(rs!���ʱ��)
'                    .TextMatrix(.Rows - 2, .ColIndex("�������")) = zlCommFun.NVL(rs!�������)
'                    .TextMatrix(.Rows - 2, .ColIndex("��Ժ����ID")) = zlCommFun.NVL(rs!��Ժ����ID)
'                End With
'                rs.MoveNext
'            Loop
'
'            mclsVsf(0).AppendRows = True
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case Else
    
        If Control.ID > 400 And Control.ID < 500 Then
           
        Else
             '��ҵ���޹صĹ��ܣ������Ĺ���
            Call CommandBarExecutePublic(Control, Me, vfgThis, "��������¼�嵥")
            
        End If
        
    
    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long
    Dim lngScaleTop  As Long
    Dim lngScaleRight  As Long
    Dim lngScaleBottom  As Long
    
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    picPane(2).Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call ExecuteCommand("ˢ������")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf(0) = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        vfgThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height - picPane(0).Height
        picPane(0).Move 0, vfgThis.Top + vfgThis.Height, vfgThis.Width
        
        txt����.Move txt����.Left, txt����.Top, picPane(0).Width - txt����.Left - 30
        
        cmdCancel.Move picPane(0).Width - cmdCancel.Width - 30, cmdCancel.Top
        cmdOK.Move cmdCancel.Left - cmdOK.Width - 30
        mclsVsf(0).AppendRows = True
    End Select
End Sub

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, Optional ByVal objPrnVsf As Object, Optional ByVal strPrintTitle As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim objControl As Object
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    Dim bytMode As Byte
        
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet              '��ӡ����
    
        Call zlPrintSet
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel               '��ӡ����,Ԥ������,�����Excel
        
        If objPrnVsf Is Nothing Then Exit Function
        
        If Not SearchPrintData(objPrnVsf, frmPubResource.msfPrint) Then
            MsgBox "���ӡ�����粻�������ݣ������¼��ӣ�", vbInformation, ParamInfo.ϵͳ����
            Exit Function
        End If
        
        '���ô�ӡ��������
        Set objPrint.Body = frmPubResource.msfPrint
        objPrint.Title.Text = strPrintTitle
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add("")
        Call objAppRow.Add("��ӡʱ��:" & Now())
        Call objPrint.BelowAppRows.Add(objAppRow)

        Select Case Control.ID
        Case conMenu_File_Print
            bytMode = zlPrintAsk(objPrint)
            If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
        Case conMenu_File_Preview
            zlPrintOrView1Grd objPrint, 2
        Case conMenu_File_Excel
            zlPrintOrView1Grd objPrint, 3
        End Select
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
    
        For lngLoop = 2 To frmMain.cbsMain.count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '��ť����
    
        For lngLoop = 2 To frmMain.cbsMain.count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '��ͼ��
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '״̬��
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '��������
    
        Call ShowHelp(App.ProductName, frmMain.hWnd, frmMain.Name, Int((ParamInfo.ϵͳ��) / 100))
        
    Case conMenu_Help_Web_Home          'Web�ϵ�����
        
        Call zlHomePage(frmMain.hWnd)
        
    Case conMenu_Help_Web_Forum         'Web�ϵ���̳
    
        Call zlWebForum(frmMain.hWnd)
        
    Case conMenu_Help_Web_Mail          '���ͷ���
        
        Call zlMailTo(frmMain.hWnd)
            
    Case conMenu_Help_About             '����
        
        Call ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_File_Exit              '�˳�
    
        Unload frmMain
            
    End Select
    
    CommandBarExecutePublic = True
End Function

Private Sub vfgThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vfgThis_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf(0).AppendRows = True
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
 mclsVsf(0).AppendRows = True
End Sub

