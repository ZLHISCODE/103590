VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmProcMain 
   Caption         =   "�Զ�����̹���"
   ClientHeight    =   6996
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   13176
   Icon            =   "frmProcMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6996
   ScaleWidth      =   13176
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picMain 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   0
      ScaleHeight     =   4752
      ScaleWidth      =   12612
      TabIndex        =   0
      Top             =   600
      Width           =   12612
      Begin VB.PictureBox picHeader 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEBD7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1800
         Left            =   15
         ScaleHeight     =   1800
         ScaleWidth      =   12540
         TabIndex        =   1
         Top             =   15
         Width           =   12540
         Begin VB.Frame fraSplit 
            Height          =   45
            Left            =   0
            TabIndex        =   28
            Top             =   1360
            Width           =   12840
         End
         Begin VB.PictureBox picStep 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Index           =   1
            Left            =   7285
            ScaleHeight     =   324
            ScaleWidth      =   936
            TabIndex        =   26
            Top             =   720
            Width           =   960
            Begin VB.Label lblStep 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "���̵���"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   84
               Width           =   720
            End
         End
         Begin VB.PictureBox picStep 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   350
            Index           =   0
            Left            =   1775
            ScaleHeight     =   324
            ScaleWidth      =   1296
            TabIndex        =   24
            Top             =   720
            Width           =   1320
            Begin VB.Label lblStep 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "Ӧ��ϵͳ����"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   85
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "����(&G)"
            Height          =   350
            Index           =   3
            Left            =   8880
            TabIndex        =   20
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "���(&J)"
            Height          =   350
            Index           =   2
            Left            =   5555
            TabIndex        =   17
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "�������(&C)"
            Height          =   350
            Index           =   1
            Left            =   3725
            TabIndex        =   14
            Top             =   720
            Width           =   1200
         End
         Begin VB.CommandButton cmdFun 
            Caption         =   "�ռ�(&S)"
            Height          =   350
            Index           =   0
            Left            =   45
            TabIndex        =   11
            Top             =   720
            Width           =   1100
         End
         Begin VB.ComboBox cboProcState 
            Height          =   276
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1476
            Width           =   1410
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   8916
            TabIndex        =   5
            ToolTipText     =   "��ֱ�Ӱ��س������й���"
            Top             =   1488
            Width           =   1695
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�û�����"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   3240
            TabIndex        =   4
            Top             =   1536
            Width           =   1305
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�հ׹���"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1872
            TabIndex        =   3
            Top             =   1536
            Width           =   1305
         End
         Begin VB.OptionButton optType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "�䶯����"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   528
            TabIndex        =   2
            Top             =   1536
            Width           =   1305
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---��"
            Height          =   180
            Index           =   3
            Left            =   6736
            TabIndex        =   23
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   360
         End
         Begin VB.Label lblWarn 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   $"frmProcMain.frx":6852
            ForeColor       =   &H002222B2&
            Height          =   540
            Left            =   48
            TabIndex        =   21
            Top             =   120
            Width           =   6348
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---��"
            Height          =   180
            Index           =   4
            Left            =   8326
            TabIndex        =   19
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblResult 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "�����������嵥"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   5555
            TabIndex        =   18
            Top             =   1120
            Width           =   1260
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---��"
            Height          =   180
            Index           =   2
            Left            =   5006
            TabIndex        =   16
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---��"
            Height          =   180
            Index           =   1
            Left            =   3176
            TabIndex        =   15
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblResult 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "��������嵥"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   48
            TabIndex        =   13
            Top             =   1120
            Width           =   1080
         End
         Begin VB.Label lblNext 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEBD7&
            Caption         =   "---��"
            Height          =   180
            Index           =   0
            Left            =   1226
            TabIndex        =   12
            Top             =   805
            Width           =   468
         End
         Begin VB.Label lblProcState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "״̬"
            Height          =   180
            Left            =   6360
            TabIndex        =   10
            Top             =   1536
            Width           =   360
         End
         Begin VB.Label lblLocation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ��"
            Height          =   180
            Left            =   8388
            TabIndex        =   6
            Top             =   1536
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1752
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   7332
         _cx             =   12938
         _cy             =   3096
         Appearance      =   1
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
         GridColor       =   -2147483632
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
         FixedCols       =   1
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmProcMain.frx":6904
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
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      TabIndex        =   8
      Top             =   6636
      Width           =   13176
      _ExtentX        =   23241
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmProcMain.frx":699B
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20355
            MinWidth        =   8819
            Text            =   "��ǰ���д�����0����������0��"
            TextSave        =   "��ǰ���д�����0����������0��"
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
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   1080
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":722F
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":77C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":7D63
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":80B5
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":E917
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15179
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15641
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   1800
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15B09
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15C50
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15D9F
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":15F21
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":160E7
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":1656A
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":169E5
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcMain.frx":16E05
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'===����
'==============================================================
Private mfrmProgramEdit As frmProcEdit
Private mfrmBuildScript As frmProcBuildScript
Private mfrmProcedureRelating As frmProcRelating
Private mfrmCollectUpdate As frmProcCollectUpdate
Private mintProcType As Integer
Private mblnReading As Boolean
Private mlngProcID As Long '��ǰ��ID
Private mobjMain As Object
Private mobjTip As clsTipSwap  '������ʾ
Private Enum OptProcType
    OPT_�䶯���� = 0
    OPT_�հ׹��� = 1
    OPT_�û����� = 2
End Enum

Private Enum ProcCol
    PC_��� = 0
    PC_ѡ�� = 1
    PC_���� = 2
    PC_״̬ = 3
    PC_˵�� = 4
End Enum


Private Enum cmdFun
    CF_�ռ� = 0
    CF_������� = 1
    CF_��� = 2
    CF_���� = 3
End Enum

'==============================================================
'==�����ӿ�
'==============================================================
Public Function ShowMe(ByVal objParent As Object)
    Me.Show 1, objParent
End Function

'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub cboProcState_Click()
    If cboProcState.Tag = "" Then
        If IsSelData Then
            mlngProcID = Val(vsfMain.RowData(vsfMain.Row))
        Else
            mlngProcID = 0
        End If
    End If
    '��ȡ��������
    Call RefreshData
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo errHand
    
    Dim rs As ADODB.Recordset
    Select Case Control.Id
    Case conMenu_File_PrintSet
        '��ӡ����
        Call zlPrintSet
    Case conMenu_File_Preview
        'Ԥ��
        PrintProcs 2
    Case conMenu_File_Print
        '��ӡ
        PrintProcs 1
    Case conMenu_File_Excel
        '�����Excel
        PrintProcs 3
    Case conMenu_Edit_NewItem
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        Call mfrmProgramEdit.ShowMe(Me, 0, mintProcType)
    Case conMenu_Edit_Modify
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        If vsfMain.RowData(vsfMain.Row) > 0 Then
            If mfrmProgramEdit.ShowMe(Me, vsfMain.RowData(vsfMain.Row), mintProcType) Then
                Call RefreshData
            End If
        End If
    Case conMenu_Edit_Disuse
        If MsgBox("��ȷ�������������" & vbCrLf & "�˲����Ὣ��������ǰ�Ĺ��̼�¼��Ϊ�ϴι��̼�¼��", vbOKCancel + vbInformation + vbDefaultButton2, "�������") = vbOK Then
            gcnOracle.Execute "Zl_Zlproceduretext_Move()"
            Call RefreshData
        End If
    Case conMenu_Edit_Audit
        If mfrmCollectUpdate Is Nothing Then Set mfrmCollectUpdate = New frmProcCollectUpdate
        If mfrmCollectUpdate.ShowMe(Me, 1) Then Call RefreshData
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Change_PaitNote
        Set rs = gclsBase.GetProcByState(1, 2)
        If rs.BOF = False Then
            MsgBox "��⵽�й��̻�δ������ɣ����Ƚ��е����������ɡ�", vbInformation + vbOKOnly, "�������"
            Exit Sub
        End If
        If mfrmBuildScript Is Nothing Then
            Set mfrmBuildScript = New frmProcBuildScript
        End If
        Call mfrmBuildScript.ShowMe(Me)
    Case conMenu_Edit_Delete
        Call FunDeleteProc
    Case conMenu_Edit_Untread
        Call FunRestoreProc
    Case conMenu_Edit_Word
        If mfrmCollectUpdate Is Nothing Then Set mfrmCollectUpdate = New frmProcCollectUpdate
        If mfrmCollectUpdate.ShowMe(Me) Then Call RefreshData
    Case conMenu_Edit_Confirm 'ȷ�ϵ���
        Call FunConfirmProc
    Case conMenu_File_Exit
        Unload Me
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '������
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text      '��ť����
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size      '��ͼ��
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    Case conMenu_View_StatusBar         '״̬��
        stbThis.Visible = Not stbThis.Visible
        cbsMain.RecalcLayout
    Case conMenu_Help_Help              '��������
'        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((ParamInfo.ϵͳ��) / 100))
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me)
    End Select
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If stbThis.Visible Then lngBottom = lngBottom - stbThis.Height
    picMain.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnSelData As Boolean
    With vsfMain
        Select Case Control.Id
            Case conMenu_Edit_Delete
                Control.Visible = Not optType(OPT_�հ׹���).value
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Modify
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Untread
                Control.Visible = (optType(OPT_�䶯����).value Or optType(OPT_�հ׹���).value)
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_Edit_Confirm 'ȷ�ϵ���
                Control.Visible = optType(OPT_�û�����).value
                Control.Enabled = IsSelData() And Control.Visible
            Case conMenu_View_ToolBar_Button            '������
                If cbsMain.Count >= 2 Then
                    Control.Checked = cbsMain(2).Visible
                End If
            Case conMenu_View_ToolBar_Text              'ͼ������
                If cbsMain.Count >= 2 Then
                    Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
                End If
            Case conMenu_View_ToolBar_Size              '��ͼ��
                Control.Checked = cbsMain.Options.LargeIcons
            Case conMenu_View_StatusBar                 '״̬��
                Control.Checked = stbThis.Visible
        End Select
    End With
End Sub

Private Sub cmdFun_Click(Index As Integer)
    Dim objControl  As CommandBarControl
    Set objControl = cbsMain.FindControl(xtpControlButton, Decode(Index, CF_�ռ�, conMenu_Edit_Word, CF_�������, conMenu_Edit_Disuse, CF_���, conMenu_Edit_Audit, CF_����, conMenu_Manage_Change_PaitNote))
    
    If Not objControl Is Nothing Then
        Call cbsMain_Execute(objControl)
    End If
End Sub

Private Sub Form_Load()
    'Ӧ��OEMͼ��
    Call ApplyOEM(stbThis)
    '��ʼ���˵�
    Call InitCommandBar
    'Ĭ��չʾ�䶯����
    optType(OPT_�䶯����).value = True
    Call OptType_Click(OPT_�䶯����)
    '��ȡ��������
    Call RefreshData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mfrmBuildScript Is Nothing) Then Unload mfrmBuildScript
    If Not (mfrmCollectUpdate Is Nothing) Then Unload mfrmCollectUpdate
    If Not (mfrmProcedureRelating Is Nothing) Then Unload mfrmProcedureRelating
    If Not (mfrmProgramEdit Is Nothing) Then Unload mfrmProgramEdit
End Sub

Private Sub OptType_Click(Index As Integer)
    Dim arrTmp As Variant, strTMp As String, i As Integer
    Dim strOldType As String, intIndex As Integer
    If IsSelData Then
        mlngProcID = Val(vsfMain.RowData(vsfMain.Row))
    Else
        mlngProcID = 0
    End If
    mintProcType = (Index + 1)
    strTMp = "ȫ��,-1,�����,0,������,1,������,2,�ѵ���,3,�ޱ仯,4"
    strOldType = cboProcState.Text: intIndex = -1
    arrTmp = Split(strTMp, ",")
    cboProcState.Clear
    For i = LBound(arrTmp) To UBound(arrTmp) Step 2
        cboProcState.AddItem arrTmp(i)
        cboProcState.ItemData(cboProcState.NewIndex) = Val(arrTmp(i + 1))
        If intIndex = -1 Then
            If arrTmp(i) = strOldType Then intIndex = cboProcState.NewIndex
        End If
    Next
    If intIndex = -1 Then intIndex = 0
    cboProcState.Tag = "ˢ��" '��ʶ����Ҫ���»�ȡ��ǰ��ID
    cboProcState.ListIndex = intIndex
    cboProcState.Tag = ""
End Sub

Private Sub optType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Call ShowTips(Index)
End Sub

Private Sub picHeader_Resize()
    On Error Resume Next
    txtLocation.Move picHeader.ScaleWidth - txtLocation.Width - 75
    lblLocation.Move txtLocation.Left - lblLocation.Width - 30
    cboProcState.Move lblLocation.Left - cboProcState.Width - 60
    lblProcState.Move cboProcState.Left - lblProcState.Width - 30
    fraSplit.Width = picHeader.ScaleWidth - fraSplit.Left
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picHeader.Move 15, 15, picMain.ScaleWidth - 30
    vsfMain.Move 15, picHeader.Top + picHeader.Height + 15, picMain.ScaleWidth - 30, picMain.ScaleHeight - (picHeader.Top + picHeader.Height + 15) - 15
End Sub

Private Sub txtLocation_GotFocus()
    Call gclsBase.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    
    If KeyAscii = vbKeyReturn Then
        intCol = vsfMain.ColIndex("����")
        lngRow = vsfMain.FindRow(UCase(txtLocation.Text), intCol, 2, vsfMain.Row + 1)
        If lngRow = -1 Then
            lngRow = vsfMain.FindRow(UCase(txtLocation.Text), intCol, 2)
        End If
        If lngRow > 0 And vsfMain.Row <> lngRow Then
            vsfMain.Row = lngRow
            vsfMain.ShowCell vsfMain.Row, vsfMain.Col
        End If
        Call gclsBase.LocationObj(txtLocation)
    End If
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfMain
        .Redraw = False
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, PC_���, .Rows - 1, PC_���) = Color.���ɫ
            .Cell(flexcpFontBold, .FixedRows, PC_���, .Rows - 1, PC_���) = False
            .Cell(flexcpFontBold, NewRow, PC_���, NewRow, PC_���) = True
            .Cell(flexcpForeColor, NewRow, PC_���, NewRow, PC_���) = Color.��ɫ
        End If
        .Redraw = True
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    Call SetSerial
End Sub

Private Sub vsfMain_BeforeSort(ByVal Col As Long, Order As Integer)
    If Col = PC_ѡ�� Then
        Call SelRow
        Order = flexSortNone
    End If
End Sub

Private Sub vsfMain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <= PC_ѡ�� Then
        Cancel = True
    End If
End Sub

Private Sub vsfMain_Click()
    If vsfMain.Col = PC_ѡ�� Then
        vsfMain.ExplorerBar = flexExNone
    Else
        vsfMain.ExplorerBar = flexExSort
    End If
    If vsfMain.Col = PC_ѡ�� Then
        Call SelRow(vsfMain.MouseRow)
    End If
End Sub

Private Sub vsfMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cbrPopupBar As CommandBar
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call gclsBase.SendLMouseButton(vsfMain.hwnd, x, y)
        Set cbrPopupBar = gclsBase.CopyMenu(cbsMain, 2)
        If cbrPopupBar Is Nothing Then Exit Sub
        cbrPopupBar.ShowPopup
    End Select
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

'==============================================================
'==˽�з���
'==============================================================
Private Sub InitCommandBar()
    '******************************************************************************************************************
    '���ܣ���ʼ�˵�������
    '��������
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel��")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Word, "�ռ��Ǽ�(&S)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "�½��Ǽ�(&N)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Disuse, "�������(&C)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "������(&J)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸Ĺ���(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "�ָ�����(&R)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Confirm, "ȷ�ϵ���(&T)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Change_PaitNote, "���ɽű�(&G)", True)
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrWebSustainer)
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "��ҳ(&H)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "��̳(&F)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
'    objBar.SetIconSize 16, 16
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Word, "�ռ�")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "�½�")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "���", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Untread, "�ָ�")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Confirm, "ȷ�ϵ���(&T)")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Manage_Change_PaitNote, "����", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Disuse, "�������(&C)", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
    
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
        
    End With
End Sub

Private Sub RefreshData()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngRow As Long, lngCurRow As Long
    Dim lng������ As Long, lng������ As Long
    Dim intState As Integer
    
    intState = cboProcState.ItemData(cboProcState.ListIndex)
    
    '���ԭ������
    strSQL = "Select Id, Decode(����, 1, '��׼����', 2, '�հ׹���', 3, '�û�����') As ����, ���� As ����," & vbNewLine & _
            "       Decode(״̬,0,'�����', 1, '������', 2, '������', 3, '�ѵ���', 4, '�ޱ仯') As ״̬����, ״̬, ˵��, �޸���Ա, �޸�ʱ��, �ϴ��޸���Ա," & vbNewLine & _
            "       �ϴ��޸�ʱ��" & vbNewLine & _
            "From Zlprocedure" & vbNewLine & _
            "Where ���� = [1]" & IIf(intState = -1, "", " And ״̬=[2] ")
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�����嵥", mintProcType, intState)
    With vsfMain
        .Redraw = flexRDNone
        .Cell(flexcpPicture, 0, PC_ѡ��) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, PC_ѡ��) = flexAlignCenterCenter
        .Rows = vsfMain.FixedRows
        .Tag = ""
        .RowData(0) = rsTmp.RecordCount
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, PC_���) = lngRow
            .TextMatrix(lngRow, PC_����) = rsTmp!���� & ""
            .TextMatrix(lngRow, PC_״̬) = rsTmp!״̬���� & ""
            .Cell(flexcpData, lngRow, PC_״̬) = rsTmp!״̬
            .TextMatrix(lngRow, PC_˵��) = rsTmp!˵�� & ""
            .RowData(lngRow) = Val(rsTmp!Id & "")
            If .RowData(lngRow) = mlngProcID Then
                lngCurRow = lngRow
            End If
            If .TextMatrix(lngRow, PC_״̬) = "������" Then
                .Cell(flexcpForeColor, lngRow, PC_״̬) = vbRed
                lng������ = lng������ + 1
            ElseIf .TextMatrix(lngRow, PC_״̬) = "������" Then
                .Cell(flexcpForeColor, lngRow, PC_״̬) = vbBlue
                lng������ = lng������ + 1
            Else
                .Cell(flexcpForeColor, lngRow, PC_״̬) = &H80000008
            End If
            rsTmp.MoveNext
        Loop
        If .Rows <> vsfMain.FixedRows Then
            vsfMain.Row = vsfMain.FixedRows
            If lngCurRow = 0 Then lngCurRow = vsfMain.FixedRows
        End If
        Call vsfMain_AfterRowColChange(-1, -1, lngCurRow, lngCurRow)
        .Redraw = flexRDDirect
    End With
    stbThis.Panels(2).Text = "��ǰ���д����� " & lng������ & " ��,������ " & lng������ & " ����"
End Sub

Private Sub SelRow(Optional ByVal lngRow As Long)
'���ܣ�����ѡ��vsDetailParas����ȡ��ѡ��
'          lngRow=0-ѡ���ȡ��ѡ�������У�>0ѡ���ȡ��ѡ��ָ����
    Dim blnSel As Boolean, i As Long
    
    With vsfMain
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngRow = 0 Then
            blnSel = Val(.ColData(PC_ѡ��)) = 0
            .Cell(flexcpPicture, lngRow, PC_ѡ��) = imgEdit.ListImages(IIf(blnSel, "AllCheck", "UnCheck")).Picture
            .ColData(PC_ѡ��) = IIf(blnSel, 1, 0) '���ͼ��״̬
            For i = .FixedRows To .Rows - 1
                If Val(.RowData(i)) <> 0 Then
                    .TextMatrix(i, PC_ѡ��) = IIf(blnSel, -1, 0)
                End If
            Next
            If blnSel Then
                .Tag = Val(.RowData(0))
            Else
                .Tag = 0
            End If
        Else
            If Val(.RowData(lngRow)) <> 0 Then
                blnSel = Val(.TextMatrix(lngRow, PC_ѡ��)) = 0
                .TextMatrix(lngRow, PC_ѡ��) = IIf(blnSel, -1, 0)
                .Tag = (Val(.Tag) + IIf(blnSel, 1, -1))
                If Val(.Tag) = 0 Then '���еĶ�δѡ����ͼ�����Ϊ����δ��ѡ
                    .Cell(flexcpPicture, 0, PC_ѡ��) = imgEdit.ListImages("UnCheck").Picture
                    .ColData(PC_ѡ��) = 0
                ElseIf Val(.Tag) = Val(.RowData(0)) Then '���еĶ�ѡ����ͼ�����Ϊ������ѡ
                    .Cell(flexcpPicture, 0, PC_ѡ��) = imgEdit.ListImages("AllCheck").Picture
                    .ColData(PC_ѡ��) = 1
                End If
            End If
        End If
    End With
End Sub

Private Sub SetSerial()
'���ܣ����������
    Dim i As Long
    With vsfMain
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, PC_���) = i
        Next
        If .Rows - 1 > 0 Then
            .Cell(flexcpForeColor, .FixedRows, PC_���, .Rows - 1, PC_���) = Color.���ɫ
            .Cell(flexcpFontBold, .FixedRows, PC_���, .Rows - 1, PC_���) = False
        End If
        If .Row > 0 Then
            .Cell(flexcpFontBold, .Row, PC_���, .Row, PC_���) = True
            .Cell(flexcpForeColor, .Row, PC_���, .Row, PC_���) = Color.��ɫ
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FunDeleteProc()
'���ܣ�ɾ���û��洢����
    Dim i As Long, lngCout As Long
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer
    Dim arrTmp As Variant, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnOperate As Boolean
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  'û��ѡ����鿴��ǰ��
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "ɾ�����̣�" & .TextMatrix(.Row, PC_����)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "ɾ������" & Decode(mintProcType, ProcType.�䶯����, "�䶯����", ProcType.�հ׹���, "�հ׹���", "�û�����") & "(����" & Val(.Tag) & "��)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%�ı�ѡ�����ȡ������
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_ѡ��)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_ѡ��)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(i, PC_����)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "ɾ�����¹��̣�" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(����" & Val(.Tag) & "��)")

        End If
        If MsgBox("ȷ��" & strProcMsg & "��?", vbInformation + vbOKCancel, "�������") = vbOK Then
            If mfrmProcedureRelating Is Nothing Then Set mfrmProcedureRelating = New frmProcRelating
            If Not mfrmProcedureRelating.CheckRelation(Me, strProcIDs) Then Exit Sub
            strSQL = GetProcSQL(strProcIDs)
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "ɾ���洢����", mintProcType, strProcIDs)
            blnOperate = True: lngCout = rsTmp.RecordCount
            Call ShowFlash("����ɾ�����̣����Ժ�", 0)
            For i = 1 To rsTmp.RecordCount
                Call ShowFlash("����ɾ�����̡�" & rsTmp!���� & "��", i / lngCout)
                strSQL = "Zl_Zlprocedure_Delete(" & rsTmp!Id & ")"
                Call ExecuteProcedure(strSQL, "ɾ���洢����")
                rsTmp.MoveNext
            Next
            Call ShowFlash("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FunRestoreProc()
'���ܣ����䶯���̻�հ׹��ָ̻�Ϊ���α�׼�洢����
    Dim i As Long, lngCout As Long, rsTmp As ADODB.Recordset
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer, strPreID As String, strPreName As String
    Dim strSQL As String, strProcText As String
    Dim blnOperate As Boolean
    Dim lngTotal As Long
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  'û��ѡ����鿴��ǰ��
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "�ָ����̣�" & .TextMatrix(.Row, PC_����)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "�ָ�����" & Decode(mintProcType, ProcType.�䶯����, "�䶯����", ProcType.�հ׹���, "�հ׹���", "�û�����") & "(����" & Val(.Tag) & "��)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%�ı�ѡ�����ȡ������
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_ѡ��)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_ѡ��)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(.Row, PC_����)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "ȷ���ָ����¹�����?�������£�" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(����" & Val(.Tag) & "��)")

        End If
        If mintProcType <> ProcType.�հ׹��� Then
            strProcMsg = "����ִ�к��ڱ��������У������ٶԸù��̽��й����ҽ����ݿ��иù��̶���ָ�Ϊ����֮��ı�׼���̡�" & vbNewLine & strProcMsg
        Else
            strProcMsg = "�����Ὣ���ݿ��иù��̶���ָ�Ϊ����֮���׼���̡�" & vbNewLine & strProcMsg
        End If
        If MsgBox(strProcMsg, vbInformation + vbOKCancel, "�������") = vbOK Then
            strSQL = "Select a.Id, a.����, b.���, b.����" & vbNewLine & _
                        "From (" & GetProcSQL(strProcIDs) & ") a, Zlproceduretext b" & vbNewLine & _
                        "Where a.Id = b.����id  And b.���� = [3]" & vbNewLine & _
                        "Order By a.Id, b.���"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "�ָ��洢����", mintProcType, strProcIDs, ProcTextType.���α�׼����)
            strPreID = "": strSQL = "": strProcText = ""
            blnOperate = True
            Call ShowFlash("���ڻָ����̣����Ժ�", 0, Me, True)
            lngTotal = Val(.Tag): lngCout = 0
            If lngTotal = 0 Then lngTotal = 1
            Do While Not rsTmp.EOF
                If rsTmp!Id & "" <> strPreID Then
                    If strPreID <> "" Then
                        lngCout = lngCout + 1
                        strProcText = strProcText
                        Call ShowFlash("���ڻָ����̡�" & strPreName & "��", lngCout / lngTotal)
                        Call gcnOldOra.Execute(strProcText)
                        If mintProcType <> ProcType.�հ׹��� Then
                            strSQL = "Zl_Zlprocedure_Delete(" & strPreID & ")"
                            Call ExecuteProcedure(strSQL, "ɾ���洢����")
                        End If
                    End If
                    strPreID = rsTmp!Id
                    strProcText = rsTmp!����
                    strPreName = rsTmp!����
                Else
                    strProcText = strProcText & rsTmp!����
                End If
                rsTmp.MoveNext
            Loop
            If strPreID <> "" Then
                strProcText = strProcText
                Call ShowFlash("���ڻָ����̡�" & strPreName & "��", 100)
                Call gcnOldOra.Execute(strProcText)
                If mintProcType <> ProcType.�հ׹��� Then
                    strSQL = "Zl_Zlprocedure_Delete(" & strPreID & ")"
                    Call ExecuteProcedure(strSQL, "ɾ���洢����")
                End If
            End If
            ShowFlash ("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FunConfirmProc()
'���ܣ�ȷ���Ѿ���������
'���ܣ����䶯���̻�հ׹��ָ̻�Ϊ���α�׼�洢����
    Dim i As Long, lngCout As Long, rsTmp As ADODB.Recordset
    Dim strProcIDs As String, strProcMsg As String
    Dim intSel As Integer
    Dim strSQL As String
    Dim blnOperate As Boolean
    Dim lngTotal As Long
    
    On Error GoTo errH
    With vsfMain
        intSel = -1
        If Val(.Tag) = 0 Then  'û��ѡ����鿴��ǰ��
            If .Row > 0 Then
                strProcIDs = .RowData(.Row)
                strProcMsg = "ȷ���Ѿ��������¹��̣�" & .TextMatrix(.Row, PC_����)
            End If
        ElseIf Val(.Tag) = .RowData(0) Then
            strProcIDs = "*" & mintProcType
            strProcMsg = "ȷ���Ѿ����������û�����(����" & Val(.Tag) & "��)"
        ElseIf Val(.Tag) > .RowData(0) * 0.9 Then '90%�ı�ѡ�����ȡ������
            strProcIDs = "-" & mintProcType
            intSel = 0
        End If
        If strProcMsg = "" Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, PC_ѡ��)) = intSel Then
                    strProcIDs = strProcIDs & "," & .RowData(i)
                End If
                If lngCout < 5 Then
                    If Val(.TextMatrix(i, PC_ѡ��)) = -1 Then
                        lngCout = lngCout + 1
                        strProcMsg = strProcMsg & vbNewLine & .TextMatrix(.Row, PC_����)
                    End If
                End If
            Next
            If strProcIDs Like ",*" Then
                strProcIDs = Mid(strProcIDs, 2)
            End If
            strProcMsg = "ȷ���Ѿ��������¹�����?�������£�" & strProcMsg & vbNewLine & _
                                IIf(lngCout = Val(.Tag), "", "... ..." & vbNewLine & "(����" & Val(.Tag) & "��)")

        End If
        If MsgBox(strProcMsg, vbInformation + vbOKCancel, "�������") = vbOK Then
            strSQL = GetProcSQL(strProcIDs)
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "ȷ�ϵ����洢����", mintProcType, strProcIDs)
            blnOperate = True
            Call ShowFlash("���ڵ������̣����Ժ�", 0, Me, True)
            lngTotal = Val(.Tag): lngCout = 0
            If lngTotal = 0 Then lngTotal = 1
            Do While Not rsTmp.EOF
                Call ShowFlash("���ڵ������̡�" & rsTmp!���� & "��", lngCout / lngTotal)
                strSQL = "Zl_Zlprocedure_Confirm(" & rsTmp!Id & ")"
                Call ExecuteProcedure(strSQL, "�����洢����")
                rsTmp.MoveNext
            Loop
            ShowFlash ("")
        Else
            Exit Sub
        End If
        blnOperate = False
    End With
    Call RefreshData
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    If blnOperate Then RefreshData
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function GetProcSQL(strIDs As String) As String
'����ID����ȡ׼��ɾ����ָ��Ĵ洢����
'strIDs=ID������*���� ��ʾ�����͵����ж���-����,ID1...:��ʾ������ȥ���ض���ID,ID1,...����ʾֻ��ȡ��ЩID
    Dim strProcs As String, intType As Integer, strTMp As String
    Dim lngPos As String, i As Integer

    '��ȡ���μ��Ĵ洢����
    If strIDs Like "[*]*" Then
        strProcs = "Select Id, Upper(����) ����, Upper(������) ������ From Zlprocedure Where ���� = [1]"
        intType = Val(Mid(strIDs, 2))
    ElseIf strIDs Like "-*" Then
        lngPos = InStr(strIDs, ",")
        strTMp = Mid(strIDs, 1, lngPos - 1)
        strIDs = Mid(strIDs, lngPos + 1)
        intType = Val(Mid(strTMp, 2))
        strProcs = "Select Id, Upper(����) ����, Upper(������) ������" & vbNewLine & _
                    "From Zlprocedure a, Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)) b" & vbNewLine & _
                    "Where ���� = [1] And a.Id = b.Column_Value(+) And b.Column_Value Is Null"
    Else
        strProcs = "Select Id, Upper(����) ����, Upper(������) ������" & vbNewLine & _
                        "From Zlprocedure" & vbNewLine & _
                        "Where Id In (Select Column_Value From Table(Cast(f_Num2list([2]) As Zltools.t_Numlist)))"
    End If
    GetProcSQL = strProcs
End Function

Private Function IsSelData() As Boolean
    If vsfMain.Row >= vsfMain.FixedRows Then
        IsSelData = vsfMain.RowData(vsfMain.Row) <> 0
    Else
        IsSelData = False
    End If
End Function

Private Sub PrintProcs(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    
    '��ͷ
    objOut.Title.Text = "���̹���" & Decode(mintProcType, ProcType.�䶯����, "�䶯����", ProcType.�հ׹���, "�հ׹���", "�û�����") & "(" & cboProcState.Text & ")"
    objOut.Title.Font.name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "ʱ�䣺" & Format(CurrentDate(), "yyyy-MM-dd HH:mm:ss")
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objOut.Body = vsfMain
    '���
    vsfMain.Redraw = False
    lngRow = vsfMain.Row: lngCol = vsfMain.Col
        
    If bytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytMode
    End If
    vsfMain.Row = lngRow: vsfMain.Col = lngCol
    vsfMain.Redraw = True
End Sub

