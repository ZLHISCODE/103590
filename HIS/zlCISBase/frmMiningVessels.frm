VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMiningVessels 
   Caption         =   "��Ѫ������"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   Icon            =   "frmMiningVessels.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   11700
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   4950
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   635
      SimpleText      =   $"frmMiningVessels.frx":000C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMiningVessels.frx":0053
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   645
      Width           =   7860
      _cx             =   13864
      _cy             =   6350
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
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
      AutoResize      =   0   'False
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.TextBox txtEdit 
         Height          =   375
         Left            =   5865
         TabIndex        =   2
         Top             =   435
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin XtremeCommandBars.ImageManager imgICON 
      Left            =   915
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMiningVessels.frx":08E7
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   300
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMiningVessels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API declares
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
Private Const mconMenu_FilePopup = 1 '�ļ�
Private Const mconMenu_ManagePopup = 2 '����
Private Const mconMenu_EditPopup = 3 '�༭
Private Const mconMenu_ReportPopup = 4 '����
Private Const mconMenu_ViewPopup = 7 '�鿴
Private Const mconMenu_ToolPopup = 8 '����
Private Const mconMenu_HelpPopup = 9 '����

'�ļ��˵�
Private Const mconMenu_File_Open = 100            '*��(&O)��
Private Const mconMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Private Const mconMenu_File_Preview = 102         '*Ԥ��(&V)
Private Const mconMenu_File_Print = 103           '*��ӡ(&P)
Private Const mconMenu_File_Excel = 104           '�����&Excel��

Private Const mconMenu_File_MedRecSetup = 1051        '��ӡ����(&S)
Private Const mconMenu_File_MedRecPreview = 1052      '��ӡԤ��(&P)

Private Const mconMenu_File_RowPrint = 121        '��¼��ӡ(&R)

Private Const mconMenu_File_Exit = 191            '*�˳�(&X)

'�༭

Private Const mconMenu_Manage_Append = 3001     '*����(&Y)
Private Const mconMenu_Manage_Delete = 3004     '*ɾ��(&D)
Private Const mconMenu_Manage_Modify = 3003       '*�޸�(&M)
Private Const mconMenu_Manage_ModifyNo = 228       '*�޸ı���(&M)
Private Const mconMenu_Manage_deleCos = 21205       '*ɾ������

Private Const mconMenu_Manage_Stop = 3503     '*����(&C)
Private Const mconMenu_Manage_Cancle = 3014   'ȡ��

'�鿴�˵�
Private Const mconMenu_View_ToolBar = 701              '������(&T)
Private Const mconMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Private Const mconMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Private Const mconMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Private Const mconMenu_View_StatusBar = 702            '״̬��(&S)
Private Const mconMenu_View_Append = 703               '������Ϣ(&A)
Private Const mconMenu_View_Expend = 711               'չ��/�۵���(&X)
Private Const mconMenu_View_Expend_CurCollapse = 7111     '�۵���ǰ��(&C)
Private Const mconMenu_View_Expend_CurExpend = 7112       'չ����ǰ��(&E)
Private Const mconMenu_View_Expend_AllCollapse = 7113     '�۵�������(&L)
Private Const mconMenu_View_Expend_AllExpend = 7114       'չ��������(&X)
Private Const mconMenu_View_Find = 721                 '*����(&F)
Private Const mconMenu_View_FindNext = 722             '��������(&N)
Private Const mconMenu_View_FindType = 723             '���ҷ�ʽ(&Y)
Private Const mconMenu_View_Filter = 731               '*���ݹ���(&I),�Ӵ���Ĺ��˹���
Private Const mconMenu_View_Notify = 732               '*ҽ������(&B)
Private Const mconMenu_View_Busy = 733                 '����æ(&M)
Private Const mconMenu_View_Hide = 741                 '*����(&H)
Private Const mconMenu_View_Show = 742                 '*��ʾ(&S)
Private Const mconMenu_View_Backward = 743             '*����(&B)
Private Const mconMenu_View_Forward = 744              '*ǰ��(&F)
Private Const mconMenu_View_Option = 781               'ѡ��(&O)
Private Const mconMenu_View_Refresh = 791              '*ˢ��(&R)
Private Const mconMenu_View_Jump = 792                 '��ת(&J)

'�����˵�
Private Const mconMenu_Help_Help = 901        '*��������(&H)
Private Const mconMenu_Help_Web = 902         '&WEB�ϵ�����
Private Const mconMenu_Help_Web_Home = 9021       '������ҳ(&H)
Private Const mconMenu_Help_Web_Forum = 9023   '������̳(&F)
Private Const mconMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Private Const mconMenu_Help_About = 991       '����(&A)��

'������������
'*********************************************************************
'CommandBar���г�������
Private Const mXTP_ID_WINDOW_LIST = 35000 '�����б�
Private Const mXTP_ID_TOOLBARLIST = 59392 '�������б�
Private Const mID_INDICATOR_CAPS = 59137 '״̬������д��
Private Const mID_INDICATOR_NUM = 59138 '״̬�������֣�
Private Const mID_INDICATOR_SCRL = 59139 '״̬����������

'CommandBar�����ȼ�
Private Const mFSHIFT = 4
Private Const mFCONTROL = 8
Private Const mFALT = 16

Private mlngModul As Long, mstrPrivs As String

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    Dim strNo As String, strSql As String, strNewNo As String
    
    Select Case Control.ID
    Case mconMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case mconMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case mconMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hWnd)
    Case mconMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case mconMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case mconMenu_Help_Help '����
        Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_File_Exit '�˳�
        Unload Me
        
    Case mconMenu_Manage_Modify  '�޸�
        vfgList.SelectionMode = flexSelectionFree
        vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_ModifyNo '�޸ı���
        If vfgList.Row > 0 Then
            
            strNo = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("Modify")))
            strNewNo = InputBox("������" & strNo & "��Ϊ��" & vbNewLine & "(ע���޸ı���󣬼�����Ŀ�еĹ��뽫ͬ��������)")
            If strNewNo <> "" And strNo <> "" Then
                With vfgList
                    For i = .FixedRows To .Rows - 1
                        If strNewNo = .TextMatrix(i, .ColIndex("Modify")) Then
                            MsgBox strNewNo & "�����б����ظ�����������һ�����룡"
                            Exit Sub
                        End If
                    Next
                End With
            
                strSql = "Zl_��Ѫ������_Clear(1,'" & strNo & "','" & strNewNo & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "�޸Ĺ���")
                Call LoadData
            End If
        End If
    Case mconMenu_Manage_deleCos
        With vfgList
            .TextMatrix(.Row, .ColIndex("��Ӧ����")) = ""
            .TextMatrix(.Row, .ColIndex("����ID")) = ""
        End With
    Case mconMenu_Manage_Stop '����

        If SaveData = True Then
            vfgList.Editable = flexEDNone
            vfgList.SelectionMode = flexSelectionByRow
            Call LoadData
        End If
    Case mconMenu_Manage_Cancle  'ȡ��
        vfgList.Editable = flexEDNone
        vfgList.SelectionMode = flexSelectionByRow
        Call LoadData
    Case mconMenu_Manage_Append '����
        With vfgList
            .SelectionMode = flexSelectionFree
            .Editable = flexEDKbdMouse
            .Rows = .Rows + 1
            If .Rows > 1 Then
                If Val(.TextMatrix(.Rows - 2, .ColIndex("����"))) <> 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(.TextMatrix(.Rows - 2, .ColIndex("����"))) + 1
                End If
            End If
            .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("��ɫ")) = vbWhite
            .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("��ɫ")) = 100
            .Select .Rows - 1, .ColIndex("����")
            
        End With
    Case mconMenu_Manage_Delete 'ɾ��
    
        If vfgList.Row > 0 Then
            'ɾ��ǰ����Ƿ���ʹ�ã�ʹ��������ա�
            If MsgBox("��ʾ��ɾ������Ŀ�󣬼�����Ŀ�ж�Ӧ�Ĺ������ý���գ��Ƿ������", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                strNo = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("Modify")))
                If strNo <> "" Then
                    strSql = "Zl_��Ѫ������_Clear(1,'" & strNo & "',Null)"
                    Call zlDatabase.ExecuteProcedure(strSql, "�������")
                End If
                Call vfgList.RemoveItem(vfgList.Row)
            End If
        End If
    End Select

End Sub

Private Function SaveData() As Boolean
    '��������
    Dim i As Integer, iRow As Integer
    Dim strSql() As String, str���� As String, str���� As String
    Dim lngColor As Long '-214748363
    On Error GoTo errHandle
    
    
    With vfgList
        
        If .Rows > 1 Then
            ReDim strSql(.Rows - 1)
        End If
        For i = 1 To .Rows - 1
            str���� = Replace(Trim(.TextMatrix(i, .ColIndex("����"))), "'", "")
            If str���� = "" Then
                MsgBox "���벻��Ϊ�գ�", vbInformation, gstrSysName
                Exit Function
            End If
            If IsNumeric(str����) = False Then
                MsgBox "���벻��Ϊ�ַ���", vbInformation, gstrSysName
                Exit Function
            End If
            str���� = Replace(Trim(.TextMatrix(i, .ColIndex("����"))), "'", "")
            If str���� = "" Then
                MsgBox "���Ʋ���Ϊ�գ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            '�������Ƿ����ظ�
            For iRow = 1 To .Rows - 1
                If i <> iRow Then
                    If str���� = Replace(Trim(.TextMatrix(iRow, .ColIndex("����"))), "'", "") Then
                        MsgBox "����[" & str���� & "]�ظ����������", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
            
            lngColor = Val(.Cell(flexcpFloodColor, i, .ColIndex("��ɫ")))
            If lngColor = -214748363 Then lngColor = 0
            strSql(i) = "Zl_��Ѫ������_Update('" & str���� & "','" & str���� & "','" & Replace(Trim(.TextMatrix(i, .ColIndex("���"))), "'", "") & "'," & _
                        "'" & Replace(Trim(.TextMatrix(i, .ColIndex("��Ӽ�"))), "'", "") & "','" & Replace(Trim(.TextMatrix(i, .ColIndex("��Ѫ��"))), "'", "") & "'," & _
                        lngColor & "," & Val(.TextMatrix(i, .ColIndex("����ID"))) & ")"
        Next
        
        If vfgList.Rows > 1 Then
            'gstrSql = "Zl_��Ѫ������_Clear"
            gcnOracle.BeginTrans
            'Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            For i = LBound(strSql) To UBound(strSql)
                If strSql(i) Like "Zl_��Ѫ������_Update*" Then
                    Call zlDatabase.ExecuteProcedure(strSql(i), Me.Caption)
                End If
            Next
            gcnOracle.CommitTrans
            SaveData = True
        End If
    End With
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.vfgList
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top - stbThis.Height
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case mconMenu_Manage_Stop, mconMenu_Manage_deleCos '����
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_Cancle 'ȡ��
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_ModifyNo, mconMenu_Manage_Delete ', mconMenu_Manage_deleCos
        Control.Enabled = Not (vfgList.Editable = flexEDKbdMouse)
     
    End Select
End Sub

Private Sub Form_Load()

    '�˵�������
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
    cbsMain.Icons = imgICON.Icons
    Call initMenus
    Call LoadData
    
End Sub

Private Sub initVfgList()
    Dim strHead As String
    '1 ����� 4 ���� 7 �Ҷ���
    strHead = "����,600,4;����,1500,1;���,1500,1;��Ӽ�,1800,1;��Ѫ��,1500,1;��ɫ,600,1;��Ӧ����,2800,1;����ID,0,1;Modify,0,1"
    Call SetVsFlexGridHead(strHead, vfgList)
    
End Sub

Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Private Sub LoadData()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    strSql = "Select A.����,A.����,A.���,A.��Ӽ�,A.��Ѫ��,A.��ɫ,A.����ID,B.����||' '||B.���� as ��Ӧ���� From ��Ѫ������ A,�շ���ĿĿ¼ B Where A.����ID=B.ID(+) Order by A.����"
    Dim lngColor As Long
    
    On Error GoTo errHandle
    With vfgList
        .Clear
        Call initVfgList
        '.ColComboList(.ColIndex("��ɫ")) = "..."
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Do Until rsTmp.EOF
             
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsTmp.Fields("����"))
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(rsTmp.Fields("����"))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Nvl(rsTmp.Fields("���"))
            .TextMatrix(.Rows - 1, .ColIndex("��Ӽ�")) = Nvl(rsTmp.Fields("��Ӽ�"))
            .TextMatrix(.Rows - 1, .ColIndex("��Ѫ��")) = Nvl(rsTmp.Fields("��Ѫ��"))
            .TextMatrix(.Rows - 1, .ColIndex("Modify")) = Nvl(rsTmp.Fields("����"))
            
            .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("��ɫ")) = 100
            lngColor = Val(Nvl(rsTmp.Fields("��ɫ")))
            If lngColor = 0 Then lngColor = -214748363
            .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("��ɫ")) = lngColor
            
            .TextMatrix(.Rows - 1, .ColIndex("����ID")) = Nvl(rsTmp.Fields("����ID"))
            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ����")) = Trim(Nvl(rsTmp.Fields("��Ӧ����")))
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '��ѡ��
        .SelectionMode = flexSelectionByRow
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) = "Update"
End Sub

Private Sub vfgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEdit As Boolean
    Call vfgList_BeforeEdit(NewRow, NewCol, blnEdit)
    If blnEdit Then
        vfgList.ComboList = ""
        'vsList.FocusRect = flexFocusLight
    Else
        'vsList.FocusRect = flexFocusSolid
        If NewCol = vfgList.ColIndex("��Ӧ����") Then
            vfgList.ComboList = "..."
        Else
            vfgList.ComboList = ""
        End If
    End If
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
    If Col = vfgList.ColIndex("��ɫ") Then
        Cancel = True
    ElseIf Col = vfgList.ColIndex("����") Then
        If vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) <> "" Then
            '��ʹ�õı��벻�ܸ�
            Cancel = True
        End If
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_DblClick()

    Dim pt As POINTAPI
    
    With vfgList
        If .Editable = flexEDNone Then Exit Sub
        
        If .MouseCol = .ColIndex("��ɫ") Then
            pt.x = .ColPos(.MouseCol) \ Screen.TwipsPerPixelX
            pt.y = (.RowPos(.MouseRow) + .RowHeight(.MouseRow)) \ Screen.TwipsPerPixelY
            ClientToScreen .hWnd, pt
            
            frmSelColor.lblRow = .MouseRow
            frmSelColor.lblCol = .MouseCol
            frmSelColor.lngColor = .Cell(flexcpFloodColor, .MouseRow, .MouseCol)
            frmSelColor.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
            frmSelColor.Show vbModal, Me
        End If
    End With
End Sub

Private Sub initMenus()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = mconMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, mconMenu_File_Excel, "�����&Excel��")

        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ManagePopup, "�༭(&E)", -1, False)
    objMenu.ID = mconMenu_ManagePopup
    With objMenu.CommandBar.Controls
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "����(&S)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "ȡ��(&C)")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "����(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "ɾ��(&D)")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_ModifyNo, "�޸ı���(&N)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_deleCos, "ɾ������"): objControl.BeginGroup = True
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = mconMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, mconMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, mconMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)") '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = mconMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, mconMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, mconMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��") '����
        Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ") '����

        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "ȡ��")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_ModifyNo, "�޸ı���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_deleCos, "ɾ������"): objControl.BeginGroup = True
        
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "ɾ��")

        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�") '����
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, mconMenu_File_Print '��ӡ
        .Add 0, vbKeyF1, mconMenu_Help_Help '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet '��ӡ����
        .AddHiddenCommand mconMenu_File_Excel '�����Excel
    End With

    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vfgListButtonClick(Row, Col)
End Sub

Private Sub vfgList_EnterCell()
    On Error GoTo errHandle
    With vfgList
    
        If .Col = .ColIndex("��Ӧ����") And .Row > 0 Then
            If txtEdit.Tag = "False" Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
                txtEdit.Tag = "True"
            End If
        Else
            txtEdit.Tag = "False"
        End If
        
        Dim blnCancle As Boolean
        Call vfgList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
 With vfgList
    If (.Col = .ColIndex("��Ӧ����")) And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        If .Col = .ColIndex("��Ӧ����") And vfgList.ComboList = "..." Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                txtEdit.Text = .EditText
                Call vfgList_CellButtonClick(.Row, .Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End If
 End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vfgList
            If Col = .ColIndex("��Ӧ����") Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call vfgListButtonClick(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            
            ElseIf Col + 1 > .Cols - 3 Then
                If Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    
                    If Val(.TextMatrix(Row, .ColIndex("����"))) <> 0 Then
                        .TextMatrix(Row + 1, .ColIndex("����")) = Val(.TextMatrix(Row, .ColIndex("����"))) + 1
                    End If
                End If
                .Cell(flexcpFloodColor, .Rows - 1, .ColIndex("��ɫ")) = vbWhite
                .Cell(flexcpFloodPercent, .Rows - 1, .ColIndex("��ɫ")) = 100
                .Select Row + 1, .ColIndex("����")
                
            Else
                .Select Row, Col + 1
            End If
        End With
    End If
End Sub


Private Sub vfgList_LeaveCell()
    Dim blnCancle As Boolean
    
    With vfgList
        Call vfgList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_RowColChange()
    On Error GoTo errHandle
    With vfgList
        If txtEdit.Tag = "True" Then
            txtEdit.Left = .CellLeft
            txtEdit.Top = .CellTop
            txtEdit.Height = .CellHeight - 12
            txtEdit.Width = .CellWidth - 12
        End If
    End With
    Exit Sub
errHandle:
    If err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vfgListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo errHandle
    
    If Col = vfgList.ColIndex("��Ӧ����") Then
        '��ȡ����
        '--------------------------------------------------------------------------------------
            strInput = DelInvalidChar(UCase(Trim(txtEdit)))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            If strInput = "" Then
                strSql = "Select A.ID, A.����, A.����, A.���, A.���㵥λ, B.�ּ�, A.��������," & vbNewLine & _
                        "       Decode(A.�������, 1, '����', 2, 'סԺ', '�����סԺ') As �������" & vbNewLine & _
                        "From (Select �ּ�, �շ�ϸĿid From �շѼ�Ŀ Where (��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                        IIf(gstrPriceClass = "", " And �۸�ȼ� Is Null ", " And �۸�ȼ� = [1] ") & ") B," & vbNewLine & _
                        "     �շ���ĿĿ¼ A" & vbNewLine & _
                        "Where A.ID = B.�շ�ϸĿid And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.��� = '4'"
            Else
                strSql = "Select /*+ rule */" & vbNewLine & _
                    " A.ID, A.����, A.����, A.���, A.���㵥λ, B.�ּ�, A.��������," & vbNewLine & _
                    " Decode(A.�������, 1, '����', 2, 'סԺ', '�����סԺ') As �������" & vbNewLine & _
                    "From �շ���Ŀ���� E," & vbNewLine & _
                    "     (Select �ּ�, �շ�ϸĿid From �շѼ�Ŀ Where (��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                    IIf(gstrPriceClass = "", " And �۸�ȼ� Is Null ", " And �۸�ȼ� = [1] ") & ") B," & vbNewLine & _
                    "     �շ���ĿĿ¼ A" & vbNewLine & _
                    "Where A.ID = E.�շ�ϸĿid And A.ID = B.�շ�ϸĿid And E.���� = 1 And" & vbNewLine & _
                    "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.��� = '4' And" & vbNewLine & _
                    "      (E.���� Like '%" & strInput & "%' Or A.���� Like '%" & strInput & "%' Or A.���� Like '%" & strInput & "%')"

            End If

            vRect = zlControl.GetControlRect(txtEdit.hWnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����", False, "", "ѡ�����", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, gstrPriceClass)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vfgList
                        .EditText = Trim(Nvl(rsTmp.Fields("����")) & " " & Nvl(rsTmp.Fields("����")))
                        .TextMatrix(.Row, .ColIndex("��Ӧ����")) = Trim(Nvl(rsTmp.Fields("����")) & " " & Nvl(rsTmp.Fields("����")))
                        .TextMatrix(.Row, .ColIndex("����ID")) = Nvl(rsTmp.Fields("ID"), "")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

