VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmLabMBReagent 
   Caption         =   "ø���Լ�����"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmLabMBReagent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4290
      Left            =   90
      ScaleHeight     =   4290
      ScaleWidth      =   8895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   8895
      Begin VB.Frame fraTool 
         Caption         =   "��������"
         Height          =   975
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8850
         Begin VB.TextBox txt��Ŀ 
            Height          =   300
            Left            =   4725
            TabIndex        =   11
            Top             =   255
            Width           =   1950
         End
         Begin VB.TextBox txt���� 
            Height          =   300
            Left            =   1035
            TabIndex        =   8
            Top             =   570
            Width           =   5640
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   1035
            TabIndex        =   5
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   73400323
            CurrentDate     =   40553
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   2520
            TabIndex        =   7
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   73400323
            CurrentDate     =   40553
         End
         Begin VB.Label lbl��Ŀ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������Ŀ"
            Height          =   180
            Left            =   3930
            TabIndex        =   10
            Top             =   285
            Width           =   720
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Լ����� "
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lblЧ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Լ�Ч��               �� "
            Height          =   180
            Left            =   240
            TabIndex        =   6
            Top             =   285
            Width           =   2340
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   3765
         Left            =   705
         TabIndex        =   2
         Top             =   1320
         Width           =   7875
         _cx             =   13891
         _cy             =   6641
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
            Left            =   5460
            TabIndex        =   3
            Top             =   2340
            Visible         =   0   'False
            Width           =   1125
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      SimpleText      =   $"frmLabMBReagent.frx":000C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMBReagent.frx":0053
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15478
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   135
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgICON 
      Bindings        =   "frmLabMBReagent.frx":08E7
      Left            =   525
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmLabMBReagent.frx":08FB
   End
End
Attribute VB_Name = "frmLabMBReagent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' API declares
Private Type POINTAPI
        x As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'�������ݲ˵�ID����:*��ʾ��ͼ��
'*********************************************************************
Private Const mconMenu_FilePopup = 1    '�ļ�
Private Const mconMenu_ManagePopup = 2  '����
Private Const mconMenu_EditPopup = 3    '�༭
Private Const mconMenu_ReportPopup = 4  '����
Private Const mconMenu_ViewPopup = 7    '�鿴
Private Const mconMenu_ToolPopup = 8     '����
Private Const mconMenu_HelpPopup = 9    '����

'�ļ��˵�
Private Const mconMenu_File_Open = 100              '*��(&O)��
Private Const mconMenu_File_PrintSet = 101          '*��ӡ����(&S)��
Private Const mconMenu_File_Preview = 102           '*Ԥ��(&V)
Private Const mconMenu_File_Print = 103             '*��ӡ(&P)
Private Const mconMenu_File_Excel = 104             '�����&Excel��

Private Const mconMenu_File_MedRecSetup = 1051      '��ӡ����(&S)
Private Const mconMenu_File_MedRecPreview = 1052    '��ӡԤ��(&P)

Private Const mconMenu_File_RowPrint = 121        '��¼��ӡ(&R)

Private Const mconMenu_File_Exit = 191            '*�˳�(&X)

'�༭

Private Const mconMenu_Manage_Append = 3001     '*����(&Y)
Private Const mconMenu_Manage_Delete = 3004     '*ɾ��(&D)
Private Const mconMenu_Manage_Modify = 3003       '*�޸�(&M)

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
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_File_Exit                             '�˳�
        Unload Me
        
    Case mconMenu_Manage_Modify                         '�޸�
        vfgList.SelectionMode = flexSelectionFree
        vfgList.Editable = flexEDKbdMouse
         
    Case mconMenu_Manage_Stop                           '����

        If SaveData = True Then
            vfgList.Editable = flexEDNone
            vfgList.SelectionMode = flexSelectionByRow
            Call LoadData
        End If
    Case mconMenu_Manage_Cancle                         'ȡ��
        vfgList.Editable = flexEDNone
        vfgList.SelectionMode = flexSelectionByRow
        Call LoadData
    Case mconMenu_Manage_Append                         '����
        With vfgList
            .SelectionMode = flexSelectionFree
            .Editable = flexEDKbdMouse
            .Rows = .Rows + 1

            .TextMatrix(.Rows - 1, .ColIndex("�Լ�����")) = Trim(.TextMatrix(.Rows - 2, .ColIndex("�Լ�����")))
            .TextMatrix(.Rows - 1, .ColIndex("�Լ�Ч��")) = Format(Now + 365, "yyyy-MM-dd")
            .Select .Rows - 1, .ColIndex("�Լ�����")
            
        End With
    Case mconMenu_Manage_Delete 'ɾ��
    
        If vfgList.Row > 0 Then
            vfgList.RemoveItem (vfgList.Row)
            vfgList.SelectionMode = flexSelectionFree
            vfgList.Editable = flexEDKbdMouse
        End If
    Case mconMenu_View_Refresh 'ˢ��
        Call LoadData
    End Select

End Sub

Private Function SaveData() As Boolean
    '��������
    Dim i As Integer, iRow As Integer
    Dim strSQL() As String, str���� As String, strЧ�� As String
    Dim blnRollBack As Boolean
    On Error GoTo ErrHandle
    
    
    With vfgList
        
        If .Rows > 1 Then
            ReDim strSQL(.Rows - 1)
        End If
        For i = 1 To .Rows - 1
            str���� = Replace(Trim(.TextMatrix(i, .ColIndex("�Լ�����"))), "'", "")
'            If str���� = "" Then
'                MsgBox "�Լ����Ų���Ϊ�գ�", vbInformation, gstrSysName
'                Exit Function
'            End If
            strЧ�� = Replace(Trim(.TextMatrix(i, .ColIndex("�Լ�Ч��"))), "'", "")
            If strЧ�� = "" Then
                
                MsgBox "�Լ�Ч�ڲ���Ϊ�գ�", vbInformation, gstrSysName
                Exit Function
            Else
                If IsDate(strЧ��) = False Then
                    MsgBox "�Լ�Ч�ڲ���ȷ���밴YYYY-MM-DD��ʽ��д��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '�������Ƿ����ظ�
            For iRow = 1 To .Rows - 1
                If i <> iRow Then
                    If str���� = Replace(Trim(.TextMatrix(iRow, .ColIndex("�Լ�����"))), "'", "") Then
                        MsgBox "�Լ�����[" & str���� & "]�ظ����������", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
            If str���� <> "" Then
                strSQL(i) = "Zl_����ø���Լ�_Edit(0,'" & str���� & "',To_Date('" & strЧ�� & "','YYYY-MM-DD'),'" & Replace(Trim(.TextMatrix(i, .ColIndex("�Լ�����"))), "'", "") & "'," & _
                            "'" & Replace(Trim(.TextMatrix(i, .ColIndex("���Է���"))), "'", "") & "'," & _
                            IIf(.TextMatrix(i, .ColIndex("������Ŀ")) = "", "Null", .TextMatrix(i, .ColIndex("��ĿID"))) & ")"
            End If
        Next
        
        If vfgList.Rows >= 2 Then
            gstrSql = "Zl_����ø���Լ�_Edit(1)"
            gcnOracle.BeginTrans
            blnRollBack = True
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            For i = LBound(strSQL) To UBound(strSQL)
                If strSQL(i) Like "Zl_����ø���Լ�_Edit*" Then
                    Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
                End If
            Next
            gcnOracle.CommitTrans
            SaveData = True
        Else
            gstrSql = "Zl_����ø���Լ�_Edit(1)"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            SaveData = True
        End If
    End With
    Exit Function
ErrHandle:
    If blnRollBack = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.picMain
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
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
    Case mconMenu_Manage_Stop, mconMenu_Manage_Cancle '����,ȡ��
        Control.Enabled = vfgList.Editable = flexEDKbdMouse
    Case mconMenu_Manage_Modify, mconMenu_Manage_Append, mconMenu_Manage_Delete, _
         mconMenu_View_Refresh  '�޸ģ�����,ɾ��,ˢ��
        Control.Enabled = vfgList.Editable = flexEDNone
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
    dtpStart.Value = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
    dtpStart.Value = CDate("3000-01-01")
    txt���� = ""
    txt��Ŀ = ""
    Call initMenus
    Call LoadData
    
End Sub

Private Sub initVfgList()
    Dim strHead As String
    '1 ����� 4 ���� 7 �Ҷ���
    strHead = "�Լ�����,1200,4;�Լ�Ч��,1000,1;�Լ�����,2800,1;���Է���,2500,1;������Ŀ,2500,1;��ĿId,0,1;Modify,0,1"
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
    Dim strSQL As String
    Dim rsTmp As New adodb.Recordset
    Dim dateS As Date, dateE As Date
    Dim strWhere As String, strItem As String, strSang As String
    
    dateS = CDate(Format(dtpStart.Value, "yyyy-MM-dd"))
    dateE = CDate(Format(dtpEnd.Value, "yyyy-MM-dd"))
    
    If dateE < dateS Then
        dtpStart = dateE
        dtpEnd = dateS
        dateS = CDate(Format(dtpStart.Value, "yyyy-MM-dd"))
        dateE = CDate(Format(dtpEnd.Value, "yyyy-MM-dd"))
    End If
    strWhere = ""
    If Trim(txt��Ŀ) <> "" Then
        strWhere = strWhere & " And B.���� Like [3] "
        strItem = "%" & DelInvalidChar(Trim(txt��Ŀ)) & "%"
    End If
    If Trim(txt����) <> "" Then
        strWhere = strWhere & " And A.�Լ����� Like [4] "
        strSang = "%" & DelInvalidChar(Trim(txt����)) & "%"
    End If
    
    strSQL = "Select A.�Լ�����, A.�Լ�Ч��, A.�Լ�����, A.���Է���, A.������Ŀid, B.����" & vbNewLine & _
            "From ����ø���Լ� A, ������ĿĿ¼ B" & vbNewLine & _
            "Where A.������Ŀid = B.ID(+) And A.�Լ�Ч�� between [1] and [2] " & vbNewLine & _
             strWhere & _
            "Order By A.�Լ�Ч�� Desc"

    With vfgList
        .Clear
        Call initVfgList
        '.ColComboList(.ColIndex("��ɫ")) = "..."
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dateS, dateE, strItem, strSang)
        Do Until rsTmp.EOF
             
            .TextMatrix(.Rows - 1, .ColIndex("�Լ�����")) = zlCommFun.Nvl(rsTmp.Fields("�Լ�����"))
            .TextMatrix(.Rows - 1, .ColIndex("�Լ�Ч��")) = zlCommFun.Nvl(rsTmp.Fields("�Լ�Ч��"))
            .TextMatrix(.Rows - 1, .ColIndex("�Լ�����")) = zlCommFun.Nvl(rsTmp.Fields("�Լ�����"))
            .TextMatrix(.Rows - 1, .ColIndex("���Է���")) = zlCommFun.Nvl(rsTmp.Fields("���Է���"))
            .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = zlCommFun.Nvl(rsTmp.Fields("����"))
            .TextMatrix(.Rows - 1, .ColIndex("��ĿID")) = zlCommFun.Nvl(rsTmp.Fields("������ĿId"))
            
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '��ѡ��
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub picMain_Resize()
    On Error Resume Next
    With Me.fraTool
        .Left = picMain.ScaleLeft
        .Top = picMain.ScaleTop
        .Width = picMain.ScaleWidth
    End With
    With Me.vfgList
        .Left = picMain.ScaleLeft
        .Top = Me.fraTool.Top + Me.fraTool.Height
        .Width = picMain.ScaleWidth
        .Height = picMain.ScaleHeight - Me.fraTool.Height
    End With

End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vfgList.TextMatrix(Row, vfgList.ColIndex("Modify")) = "Update"
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub vfgList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vfgList.Editable = flexEDNone Then Exit Sub
    If NewCol = vfgList.ColIndex("������Ŀ") Then
        vfgList.ComboList = "..."
    Else
        vfgList.ComboList = ""
    End If
End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vfgListButtonClick(Row, Col)
End Sub

Private Sub vfgList_DblClick()

'    Dim pt As POINTAPI
'
'    With vfgList
'        If .Editable = flexEDNone Then Exit Sub
'
'        If .MouseCol = .ColIndex("��ɫ") Then
'            pt.x = .ColPos(.MouseCol) \ Screen.TwipsPerPixelX
'            pt.y = (.RowPos(.MouseRow) + .RowHeight(.MouseRow)) \ Screen.TwipsPerPixelY
'            ClientToScreen .hWnd, pt
'
'            frmSelColor.lblRow = .MouseRow
'            frmSelColor.lblCol = .MouseCol
'            frmSelColor.lngColor = .Cell(flexcpFloodColor, .MouseRow, .MouseCol)
'            frmSelColor.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
'            frmSelColor.Show vbModal, Me
'
'        End If
'    End With
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
        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
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
       ' Set objControl = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��") '����
       ' Set objControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ") '����

        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Stop, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Cancle, "ȡ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Append, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, mconMenu_Manage_Delete, "ɾ��")

        Set objControl = .Add(xtpControlButton, mconMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�") '����
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
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

Private Sub vfgList_EnterCell()
    On Error GoTo ErrHandle
    With vfgList
    
        If .Col = .ColIndex("������Ŀ") And .Row > 0 Then
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
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    With vfgList
    If (.Col = .ColIndex("������Ŀ")) And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    Else
        If .Col = .ColIndex("������Ŀ") And vfgList.ComboList = "..." Then
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
            If Col = .ColIndex("������Ŀ") Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call vfgListButtonClick(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
            
            ElseIf Col + 1 > .Cols - 3 Then
                If Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                End If
                If .TextMatrix(Row + 1, .ColIndex("�Լ�����")) = "" Then .TextMatrix(Row + 1, .ColIndex("�Լ�����")) = Trim(.TextMatrix(Row, .ColIndex("�Լ�����")))
                If .TextMatrix(Row + 1, .ColIndex("�Լ�Ч��")) = "" Then .TextMatrix(Row + 1, .ColIndex("�Լ�Ч��")) = Format(Now + 365, "yyyy-MM-dd")
                
                .Select Row + 1, .ColIndex("�Լ�Ч��")
            Else
                .Select Row, Col + 1
            End If
        End With
    End If
End Sub

Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, mconMenu_ManagePopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub
Private Sub vfgListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New adodb.Recordset
    Dim strSQL   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    
    If Col = vfgList.ColIndex("������Ŀ") Then
        '��ȡ����
        '--------------------------------------------------------------------------------------
            strInput = UCase(Trim(txtEdit))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            If strInput = "" Then
                strSQL = "Select C.ID, C.����, C.����, A.��д" & vbNewLine & _
                    "From ������Ŀ A, ���鱨����Ŀ B, ������ĿĿ¼ C" & vbNewLine & _
                    "Where A.������Ŀid = B.������Ŀid And B.������Ŀid = C.ID And ��Ŀ��� = 4 And C.�����Ŀ = 0"
            Else
                strSQL = "Select C.ID, C.����, C.����, A.��д" & vbNewLine & _
                    "From ������Ŀ A, ���鱨����Ŀ B, ������ĿĿ¼ C" & vbNewLine & _
                    "Where A.������Ŀid = B.������Ŀid And B.������Ŀid = C.ID And ��Ŀ��� = 4 And C.�����Ŀ = 0 " & vbNewLine & _
                    " And (C.���� like '%" & UCase(strInput) & "%' or C.���� like '%" & UCase(strInput) & "%' or ��д Like '%" & UCase(strInput) & "%')"
            End If

            vRect = GetControlRect(txtEdit.hWnd)
            Set rsTmp = New adodb.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ����Ŀ", False, "", "", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vfgList
                        .EditText = Trim(zlCommFun.Nvl(rsTmp.Fields("����")))
                        .TextMatrix(.Row, .ColIndex("������Ŀ")) = Trim(zlCommFun.Nvl(rsTmp.Fields("����")))
                        .TextMatrix(.Row, .ColIndex("��ĿID")) = zlCommFun.Nvl(rsTmp.Fields("ID"), "")
                    End With
                End If
                Set rsTmp = Nothing
            Else
                With vfgList
                    .EditText = ""
                    .TextMatrix(.Row, .ColIndex("������Ŀ")) = ""
                End With
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

