VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRASetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9360
   Icon            =   "frmRASetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   8040
      TabIndex        =   3
      Top             =   6100
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   6840
      TabIndex        =   2
      Top             =   6100
      Width           =   1095
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "frmRASetup.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFind(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwDept"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAll(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFind(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "ҽ��(&2)"
      TabPicture(1)   =   "frmRASetup.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFind(1)"
      Tab(1).Control(1)=   "chkAll(1)"
      Tab(1).Control(2)=   "lvwDoctor"
      Tab(1).Control(3)=   "lblFind(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "���(&3)"
      TabPicture(2)   =   "frmRASetup.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFind(2)"
      Tab(2).Control(1)=   "vsfDiagnose"
      Tab(2).Control(2)=   "optDiagnose(1)"
      Tab(2).Control(3)=   "optDiagnose(0)"
      Tab(2).Control(4)=   "lblFind(2)"
      Tab(2).Control(5)=   "lblDiagnose"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "ҩƷ(&4)"
      TabPicture(3)   =   "frmRASetup.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picDrug"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "picSplit"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "tvwDrug"
      Tab(3).Control(3)=   "txtFind(3)"
      Tab(3).Control(4)=   "imgDrug"
      Tab(3).Control(5)=   "lblFind(3)"
      Tab(3).ControlCount=   6
      Begin VB.PictureBox picDrug 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5235
         Left            =   -70440
         ScaleHeight     =   5235
         ScaleWidth      =   4455
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   420
         Width           =   4455
         Begin VSFlex8Ctl.VSFlexGrid vsfDrug 
            Height          =   3555
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   3135
            _cx             =   5530
            _cy             =   6271
            Appearance      =   0
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
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
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
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   -70560
         ScaleHeight     =   4335
         ScaleWidth      =   75
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   75
      End
      Begin MSComctlLib.TreeView tvwDrug 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8705
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   -73200
         TabIndex        =   19
         Top             =   420
         Width           =   2535
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   -69480
         TabIndex        =   16
         Top             =   420
         Width           =   3495
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   -72720
         TabIndex        =   9
         Top             =   420
         Width           =   3495
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   1800
         TabIndex        =   5
         Top             =   420
         Width           =   3495
      End
      Begin MSComctlLib.ImageList imgDrug 
         Left            =   -72840
         Top             =   2880
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
               Picture         =   "frmRASetup.frx":04B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRASetup.frx":0A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRASetup.frx":0F56
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDiagnose 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   8895
         _cx             =   15690
         _cy             =   8705
         Appearance      =   0
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
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
      Begin VB.OptionButton optDiagnose 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   -72360
         TabIndex        =   14
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton optDiagnose 
         Caption         =   "���"
         Height          =   180
         Index           =   0
         Left            =   -73320
         TabIndex        =   13
         Top             =   450
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   180
         Index           =   1
         Left            =   -67080
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   180
         Index           =   0
         Left            =   7920
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwDoctor 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ұ��������(&F)"
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   18
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ұ��������(&F)"
         Height          =   180
         Index           =   2
         Left            =   -71160
         TabIndex        =   15
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ұ��롢����������(&F)"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   8
         Top             =   450
         Width           =   2070
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ұ��������(&F)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label lblDiagnose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ��������(&D)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   12
         Top             =   450
         Width           =   1350
      End
   End
   Begin VB.Label lblExplain 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRASetup.frx":14A8
      Height          =   585
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   6135
   End
End
Attribute VB_Name = "frmRASetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long
Private mstrPrivs As String
Private mblnEnter As Boolean        '�Ƿ���ɳ�ʼ�����̣�True��ɣ�False���ڳ�ʼ������
Private mblnMemory As Boolean       'ʹ�ø��Ի����True���ã�Falseδ����
Private mblnOutPatient As Boolean   '���ﴦ����飻True������Falseδ��
Private mintRecno As Integer
Private mrsDrug As ADODB.Recordset

Private Const MSTR_VSFDIAGNOSE As String = "����,,3,2000,|����,,3,6000,|ID,,1,,n"
Private Const MSTR_VSFDRUG As String = "ID,,1,0|����,,3,1000|����,,3,2500|������λ,,3,1000|����,,3,1000|��������,,3,1000|��������,,3,1000|" & _
                                       "����,,3,1000|ԭ��ҩ,,3,1000|����ҩ,,3,1000|�����Ա�,,3,1000"

Private Sub Form_Load()
    Dim lngTmp As Long

    mlngModule = glngModule
    mstrPrivs = zlstr.FormatString(";[1];", GetPrivFunc(glngSys, mlngModule))
    mblnMemory = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys))
    mblnOutPatient = (lngTmp = 1 Or lngTmp = 3)

    mblnEnter = False

    InitLV lvwDept
    InitLV lvwDoctor
    InitVSF vsfDiagnose
    InitVSF vsfDrug: vsfDrug.AllowSelection = True: vsfDrug.SelectionMode = flexSelectionListBox
    InitTVWDrug
    InitOther
    
    SetVSFHead vsfDiagnose, MSTR_VSFDIAGNOSE
    vsfDiagnose.ColComboList(vsfDiagnose.ColIndex("����")) = "..."
    
    SetVSFHead vsfDrug, MSTR_VSFDRUG

    FillLV lvwDept
    FillLV lvwDoctor
    
    FillVSFDiagnose IIf(optDiagnose(0).Value, 0, 1)
    
    FillTVDrug
    
    RestoreWinState Me, App.ProductName
    
    mblnEnter = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FindData txtFind(sstMain.Tab), sstMain.Tab, True
    End If
End Sub

Private Sub chkAll_Click(Index As Integer)
    Dim i As Integer
    Dim lvwTmp As ListView

    Select Case Index
        Case 0      '����
            Set lvwTmp = lvwDept
        Case 1      'ҽ��
            Set lvwTmp = lvwDoctor
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To lvwTmp.ListItems.Count
        lvwTmp.ListItems(i).Checked = chkAll(Index).Value = 1
    Next
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, intSN As Long
    Dim strIDs As String
    Dim colSQL As New Collection
    
    intSN = 1
    
    '����
    With lvwDept
        strIDs = ""
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                strIDs = strIDs & Trim(Mid(.ListItems(i).Key, 2)) & IIf(i >= .ListItems.Count, "", ",")
            End If
            '����ID����Ҫ��ֺ��ύSQL
            If Len(strIDs) > 3900 Or i >= .ListItems.Count Then
                If Right(strIDs, 1) = "," Then
                    strIDs = Left(strIDs, Len(strIDs) - 1)
                End If
                
                gstrSQL = "ZL_�����������_UPDATE"
                gstrSQL = gstrSQL & "(1,"                                   '���1-����
                gstrSQL = gstrSQL & intSN & ","                             '���
                gstrSQL = gstrSQL & zlstr.FormatString("'[1]')", strIDs)              'ID��
                
                'SQL���뼯�϶���
                AddArray colSQL, gstrSQL
                
                intSN = intSN + 1
                strIDs = ""
            End If
        Next
        
    End With
    
    'ҽ��
    With lvwDoctor
        strIDs = ""
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                strIDs = strIDs & Trim(Mid(.ListItems(i).Key, 2)) & IIf(i >= .ListItems.Count, "", ",")
            End If
            '����ID����Ҫ��ֺ��ύSQL
            If Len(strIDs) > 3900 Or i >= .ListItems.Count Then
                If Right(strIDs, 1) = "," Then
                    strIDs = Left(strIDs, Len(strIDs) - 1)
                End If
                
                gstrSQL = "ZL_�����������_UPDATE"
                gstrSQL = gstrSQL & "(2,"                                   '���2-ҽ��
                gstrSQL = gstrSQL & intSN & ","                             '���
                gstrSQL = gstrSQL & zlstr.FormatString("'[1]')", strIDs)              'ID��
                
                'SQL���뼯�϶���
                AddArray colSQL, gstrSQL
                
                intSN = intSN + 1
                strIDs = ""
            End If
        Next
    End With
    
    '���
    With vsfDiagnose
        strIDs = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ID"))) > 0 Then
                strIDs = strIDs & Trim(.TextMatrix(i, .ColIndex("ID"))) & ","
            End If
            '����ID����Ҫ��ֺ��ύSQL
            If Len(strIDs) > 3900 Or i >= .Rows - 1 Then
                If Right(strIDs, 1) = "," Then
                    strIDs = Left(strIDs, Len(strIDs) - 1)
                End If
                
                gstrSQL = "ZL_�����������_UPDATE"
                gstrSQL = gstrSQL & zlstr.FormatString("([1],", IIf(optDiagnose(0).Value, "3", "4"))  '���3-��ϣ�4-����
                gstrSQL = gstrSQL & intSN & ","                                             '���
                gstrSQL = gstrSQL & zlstr.FormatString("'[1]')", strIDs)                              'ID��
                
                'SQL���뼯�϶���
                AddArray colSQL, gstrSQL
                
                intSN = intSN + 1
                strIDs = ""
            End If
        Next
    End With
    
    'ҩƷ
    With vsfDrug
        strIDs = ""
        For i = 1 To .Rows - 1
            If Val(Mid(.TextMatrix(i, .ColIndex("ID")), 3)) > 0 Then
                strIDs = strIDs & Trim(Mid(.TextMatrix(i, .ColIndex("ID")), 3)) & ","
            End If
            '����ID����Ҫ��ֺ��ύSQL
            If Len(strIDs) > 3900 Or i >= .Rows - 1 Then
                If Right(strIDs, 1) = "," Then
                    strIDs = Left(strIDs, Len(strIDs) - 1)
                End If
                
                gstrSQL = "ZL_�����������_UPDATE"
                gstrSQL = gstrSQL & "(5,"                                   '���5-ҩƷ
                gstrSQL = gstrSQL & intSN & ","                             '���
                gstrSQL = gstrSQL & zlstr.FormatString("'[1]')", strIDs)              'ID��
                
                'SQL���뼯�϶���
                AddArray colSQL, gstrSQL
                
                intSN = intSN + 1
                strIDs = ""
            End If
        Next
    End With
    
    'ִ�д洢����
    Err = 0: On Error GoTo errHandle
    ExecuteProcedureArray colSQL, Me.Caption
    
    Unload Me
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    Else
        gcnOracle.RollbackTrans
    End If
End Sub

Private Sub InitOther()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    sstMain.Tab = 0
    
    On Error GoTo errHandle
    
    For i = txtFind.LBound To txtFind.UBound
        txtFind(i).ToolTipText = "��F3����������"
    Next
    
    lblDiagnose.Tag = "0"
    
    '��ʼ����ϵ�ѡ�����ͨ���������������.��𡱵��ڡ�3-��ϡ���4-�������ļ�¼��ȷ�������û�м�¼��ȱʡΪ��3-��ϡ�
    gstrSQL = "Select ��� From ����������� Where ��� In (3, 4) And Rownum < 2 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ϵ����")
    If rsTmp.RecordCount = 1 Then
        If rsTmp!��� = 4 Then
            lblDiagnose.Tag = "1"
            optDiagnose(1).Value = True
        Else
            lblDiagnose.Tag = "0"
            optDiagnose(0).Value = True
        End If
    Else
        'Ĭ��
        lblDiagnose.Tag = "1"
        optDiagnose(1).Value = True
    End If
    rsTmp.Close
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitLV(ByRef lvwVar As ListView)
'���ܣ���ʼ�������ListView�ؼ��ķ��
'������
'  lvwVar��Ҫ��ʼ����ListView�ؼ�

    'ListView���
    With lvwVar
        .Appearance = ccFlat
        .Checkboxes = UCase(lvwVar.Name) <> "LVWDRUG"   'ҩƷ����ҪCheckboxs����
        .HideSelection = False
        .HideColumnHeaders = False
        .View = lvwReport
    End With
    
End Sub

Private Sub InitTVWDrug()
    With tvwDrug
        Set .ImageList = Me.imgDrug
        .Indentation = 200
    End With
End Sub

Private Sub InitVSF(ByVal vsfVar As VSFlexGrid)
'���ܣ���ʼ�������VSFlexGrid�ؼ��ķ��
'������
'  vsfVar��Ҫ��ʼ����VSFlexGrid�ؼ�

    With vsfVar
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionListBox
        .SheetBorder = .BackColor
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub FillLV(ByRef lvwVar As ListView)
'���ܣ���ListView�������
'������
'  lvwVar��Ҫ������ݵ�ListView�ؼ�

    Dim rsTmp As ADODB.Recordset

    On Error GoTo errHandle
    Select Case UCase(lvwVar.Name)
        Case "LVWDEPT"      '����
            If lvwVar.ColumnHeaders.Count <= 0 Then
                '����ListView��
                SetLVColumnHeaders lvwDept, "ѡ��,,550,,|����,,1500,,|����,,6500"
            End If
            
            gstrSQL = "Select b.Id, b.����, b.����, Decode(Nvl(c.����id, 0), 0, 0, 1) ѡ�� " & vbCr & _
                      "From ��������˵�� A, ���ű� B, ����������� C " & vbCr & _
                      "Where a.����id = b.Id And b.Id = c.����id(+) And a.�������� = '�ٴ�' And a.������� In (1, 3) " & vbCr & _
                      "  And (b.����ʱ�� Is Null Or To_Char(b.����ʱ��, 'yyyy') = '3000') And c.���(+) = 1 " & vbCr & _
                      "Order By b.���� "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������������Ŀ���")
            
            '�������
            FillLVData rsTmp, lvwDept, "ѡ��", "ID"
            
        Case "LVWDOCTOR"    'ҽ��
            If lvwVar.ColumnHeaders.Count <= 0 Then
                '����ListView��
                SetLVColumnHeaders lvwDoctor, "ѡ��,,550,,|���,,1500,,|����,,1500,,|��������,,5000"
            End If
            
            gstrSQL = "Select b.Id, b.���, b.����, f_List2Str(Cast(Collect(d.����) as t_StrList), ',') ��������, Decode(Nvl(f.ҽ��id, 0), 0, 0, 1) ѡ�� " & vbCr & _
                      "From ��Ա����˵�� A, ��Ա�� B, ������Ա C, ���ű� D, ��������˵�� E, ����������� F " & vbCr & _
                      "Where a.��Աid = b.Id And b.Id = c.��Աid And c.����id = d.Id And d.Id = e.����id And b.Id = f.ҽ��id(+)  " & vbCr & _
                      "  And a.��Ա���� = 'ҽ��' And (d.����ʱ�� Is Null Or To_Char(d.����ʱ��, 'yyyy') = '3000') " & vbCr & _
                      "  And e.�������� = '�ٴ�' And e.������� In (1, 3) And f.���(+) = 2 " & vbCr & _
                      "Group By b.Id, b.���, b.����, Decode(Nvl(f.ҽ��id, 0), 0, 0, 1) " & vbCr & _
                      "Order By b.��� "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������������ҽ��")
            
            '�������
            FillLVData rsTmp, lvwDoctor, "ѡ��", "ID"
            
    End Select
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillTVDrug()
'���ܣ���tvwDrug�������
    
    Dim rsTmp As ADODB.Recordset
    Dim nodTmp As Node
    Dim intClass As Integer
    Dim strTmp As String
    
    '��������ҩ���г�ҩ���в�ҩ���
    With tvwDrug
        .Nodes.Clear
        Set nodTmp = .Nodes.Add(, tvwChild, "R_1", "����ҩ", 1) ': nodTmp.Expanded = True
        Set nodTmp = .Nodes.Add(, tvwChild, "R_2", "�г�ҩ", 1) ': nodTmp.Expanded = True
        Set nodTmp = .Nodes.Add(, tvwChild, "R_3", "�в�ҩ", 1) ': nodTmp.Expanded = True
    End With
    
    On Error GoTo errHandle
    gstrSQL = "Select * From (" & vbCr & _
              "    Select 'P_' || ID ID, Decode(�ϼ�id, Null, Null, 'P_' || �ϼ�id) �ϼ�id, ����, ����, ����, Null ��ѡ, " & vbCr & _
              "       Null ������λ, Null ����, Null ��������, Null ��������, Null ����, Null ԭ��ҩ, Null ����ҩ, Null �����Ա� " & vbCr & _
              "    From ���Ʒ���Ŀ¼ " & vbCr & _
              "    Where ���� In (1, 2, 3) And (����ʱ�� Is Null Or To_Char(����ʱ��, 'yyyy') = '3000') " & vbCr & _
              "    Order by ����)" & vbCr & _
              "Union All " & vbCr & _
              "Select * From (" & vbCr & _
              "    Select 'C_' || a.ID ID, Decode(a.����id, Null, Null, 'P_' || a.����id) �ϼ�id, a.����, a.����, Null, decode(nvl(b.ҩ��ID, 0), 0, 0, 1), " & vbCr & _
              "      a.���㵥λ ������λ, c.ҩƷ���� ����, c.��������, Decode(c.�Ƿ�Ƥ��, 1, '��Ҫ', Null) ��������, c.������� ����, " & vbCr & _
              "      Decode(c.�Ƿ�ԭ��, 1, '��', Null) ԭ��ҩ, Decode(c.����ҩ��, 1, '��', Null) ����ҩ, " & vbCr & _
              "      Decode(a.�����Ա�, 1, '����', 2, 'Ů��', '������') �����Ա� " & vbCr & _
              "    From ������ĿĿ¼ A, ����������� B, ҩƷ���� C " & vbCr & _
              "    Where a.ID = b.ҩ��ID(+) And a.Id = c.ҩ��ID " & vbCr & _
              "      And a.��� In ('5', '6', '7') And (a.����ʱ�� Is Null Or To_Char(a.����ʱ��, 'yyyy') = '3000') " & vbCr & _
              "      And b.���(+) = 5 " & vbCr & _
              "    Order by ����) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ������Ʒ������")
    
    On Error GoTo hErr
    With rsTmp
    
        '���ƹ���
        Set mrsDrug = rsTmp.Clone
    
        Do While .EOF = False
            intClass = zlCommFun.NVL(!����, 0)
            strTmp = zlstr.FormatString("��[1]��[2]", !����, !����)
            If zlCommFun.NVL(!�ϼ�ID, 0) = 0 Then
                Select Case intClass
                    Case 1  '����ҩ
                        Set nodTmp = tvwDrug.Nodes.Add("R_1", tvwChild, !ID, strTmp, 1)
                    Case 2  '�г�ҩ
                        Set nodTmp = tvwDrug.Nodes.Add("R_2", tvwChild, !ID, strTmp, 1)
                    Case 3  '�в�ҩ
                        Set nodTmp = tvwDrug.Nodes.Add("R_3", tvwChild, !ID, strTmp, 1)
                    Case Else
                        Set nodTmp = Nothing
                End Select
            Else
                Set nodTmp = tvwDrug.Nodes.Add(CStr(!�ϼ�ID), tvwChild, !ID, strTmp, IIf(IsNull(!����), 3, 1))
                If Val(zlCommFun.NVL(!��ѡ)) = 1 Then
                    nodTmp.Checked = True
                    Call tvwDrug_NodeCheck(nodTmp)
                End If
            End If
            
            .MoveNext
        Loop
        
    End With
    
    '����չ������ͼ��
    For Each nodTmp In tvwDrug.Nodes
        nodTmp.ExpandedImage = 2
    Next

    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
    Exit Sub
    
hErr:
    MsgBox "������" & strTmp & "���쳣��", vbInformation, gstrSysName
End Sub

Private Sub FillVSFDiagnose(ByVal bytClass As Byte)
'���ܣ���vsfDiagnose�������
'������
'  bytClass������������ݣ�0-��ϣ�1-����

    Dim rsTmp As ADODB.Recordset

    On Error GoTo errHandle
    If bytClass = 1 Then
        gstrSQL = "Select a.Id, a.����, a.���� From ��������Ŀ¼ A, ����������� B Where a.Id = b.����id " & vbCr & _
                  "    And (a.����ʱ�� Is Null Or a.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And b.��� = 4 " & vbCr & _
                  "Union All " & vbCr & _
                  "Select Null, Null, Null From Dual "  'Ϊ��������Ԥ������
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������Ϣ")
    Else
        gstrSQL = "Select a.Id, a.����, a.���� From �������Ŀ¼ A, ����������� B Where a.Id = b.���id " & vbCr & _
                  "    And (a.����ʱ�� Is Null Or a.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) And b.��� = 3 " & vbCr & _
                  "Union All " & vbCr & _
                  "Select Null, Null, Null From Dual "  'Ϊ��������Ԥ������
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ϣ")
    End If
    
    FillVSFData vsfDiagnose, rsTmp
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl9ComLib.SaveWinState Me, App.ProductName
End Sub

Private Sub optDiagnose_Click(Index As Integer)
    '��ֹ���δ����¼�
    If Val(lblDiagnose.Tag) = Index Or mblnEnter = False Then Exit Sub
    
    If vsfDiagnose.Rows < 2 Then ' Or vsfDiagnose.Rows = 2 And vsfDiagnose.TextMatrix(1, 0) = "" Then
        FillVSFDiagnose Index
        lblDiagnose.Tag = CStr(Index)
        Exit Sub
    End If
    
    If MsgBox("�л�������Ὣԭ�����������������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        optDiagnose(Val(lblDiagnose.Tag)).Value = 1
        optDiagnose(Val(lblDiagnose.Tag)).SetFocus
        Exit Sub
    End If
    
    'ˢ��vsfDiagnose�ؼ�
    FillVSFDiagnose Index
    
    lblDiagnose.Tag = CStr(Index)
End Sub

Private Sub picDrug_Resize()
    On Error Resume Next
    
    With vsfDrug
        .Left = 0
        .Top = 0
        .Width = picDrug.ScaleWidth
        .Height = picDrug.ScaleHeight
    End With
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picSplit.Left = picSplit.Left + X
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picSplit.Left < 2500 Then picSplit.Left = 2500
        If picSplit.Left > 6000 Then picSplit.Left = 6000
        txtFind(3).Width = picSplit.Left - txtFind(3).Left - 15
        tvwDrug.Width = picSplit.Left - tvwDrug.Left - 15
        picDrug.Left = picSplit.Left + picSplit.Width + 15
        picDrug.Width = sstMain.Width - picDrug.Left - tvwDrug.Left - 15
    End If
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    Select Case sstMain.Tab
        Case 0  '����
            If lvwDept.Visible Then lvwDept.SetFocus
        Case 1  'ҽ��
            If lvwDoctor.Visible Then lvwDoctor.SetFocus
        Case 2  '���
            If vsfDiagnose.Visible Then vsfDiagnose.SetFocus
        Case 3  'ҩ��
            '�����ؼ��������С��Form_Loadִ�����´��������sstab��bug��ˢ��ҳ�治������
            If picSplit.Tag <> "1" Then
                With picSplit
                    .MousePointer = 9
                    .Move tvwDrug.Left + tvwDrug.Width + 15, picDrug.Top, 30, picDrug.Height
                End With
                With txtFind(3)
                    .Width = picSplit.Left - .Left - 15
                End With
                With picDrug
                    .Left = picSplit.Left + picSplit.Width + 15
                    .Width = sstMain.Width - .Left - tvwDrug.Left
                End With
                picSplit.Tag = "1"
            End If
            
            If tvwDrug.Visible Then tvwDrug.SetFocus
    End Select
End Sub

Private Sub tvwDrug_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim l As Long
    Dim blnFind As Boolean
    
    On Error GoTo errHandle
    Screen.MousePointer = vbHourglass
    vsfDrug.Redraw = flexRDNone
    If Node.Checked Then
        '����
        NodeChecked Node, True
        NodeChecked Node, True, False
    Else
        'ɾ��
        NodeChecked Node, False
        NodeChecked Node, False, False
    End If
    vsfDrug.Redraw = flexRDDirect
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
errHandle:
    Screen.MousePointer = vbDefault
    Call ErrCenter
End Sub

Private Sub NodeChecked(ByVal nodVar As MSComctlLib.Node, ByVal blnVar As Boolean, Optional ByVal blnDown As Boolean = True)
'���ܣ��ݹ����µ������ӽ��
'������
'  nodVar��������
'  blnVar��True��ѡ��Falseȡ����ѡ

    If nodVar Is Nothing Then Exit Sub
    
    Dim nodTmp As MSComctlLib.Node
    Dim blnFind As Boolean

    If blnDown Then
        If nodVar.Child Is Nothing And nodVar.Image = 3 Then
            Dim lngRow As Long
            
            If blnVar Then
                With vsfDrug
                    '����ظ�
                    For lngRow = 0 To .Rows - 1
                        If .TextMatrix(lngRow, .ColIndex("ID")) = nodVar.Key Then
                            Exit Sub
                        End If
                    Next
                    
                    '����
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    
                    mrsDrug.Filter = zlstr.FormatString("ID='[1]'", nodVar.Key)
                    If mrsDrug.RecordCount > 0 Then
                        .TextMatrix(lngRow, .ColIndex("ID")) = nodVar.Key
                        .TextMatrix(lngRow, .ColIndex("����")) = mrsDrug!����
                        .TextMatrix(lngRow, .ColIndex("����")) = mrsDrug!����
                        .TextMatrix(lngRow, .ColIndex("������λ")) = zlCommFun.NVL(mrsDrug!������λ)
                        .TextMatrix(lngRow, .ColIndex("����")) = zlCommFun.NVL(mrsDrug!����)
                        .TextMatrix(lngRow, .ColIndex("��������")) = zlCommFun.NVL(mrsDrug!��������)
                        .TextMatrix(lngRow, .ColIndex("��������")) = zlCommFun.NVL(mrsDrug!��������)
                        .TextMatrix(lngRow, .ColIndex("����")) = zlCommFun.NVL(mrsDrug!����)
                        .TextMatrix(lngRow, .ColIndex("ԭ��ҩ")) = zlCommFun.NVL(mrsDrug!ԭ��ҩ)
                        .TextMatrix(lngRow, .ColIndex("����ҩ")) = zlCommFun.NVL(mrsDrug!����ҩ)
                        .TextMatrix(lngRow, .ColIndex("�����Ա�")) = zlCommFun.NVL(mrsDrug!�����Ա�)
                    End If
                End With
            Else
                'ɾ��
                With vsfDrug
                    For lngRow = 0 To .Rows - 1
                        If .TextMatrix(lngRow, .ColIndex("ID")) = nodVar.Key Then
                            .RemoveItem lngRow
                            Exit For
                        End If
                    Next
                End With
            End If
            
            nodVar.Checked = blnVar
        Else
            '�ݹ�
            Set nodVar = nodVar.Child
            Do While Not nodVar Is Nothing
                NodeChecked nodVar, blnVar
                nodVar.Checked = blnVar
                Set nodVar = nodVar.Next
            Loop
        End If
    Else
        nodVar.Checked = blnVar
        If Not nodVar.Parent Is Nothing Then 'And nodVar.Parent.Checked <> blnVar Then
            '���ͬ������Ƿ�δ��ѡ
            blnFind = False
            Set nodTmp = nodVar.FirstSibling
            Do While Not nodTmp Is Nothing
                If nodTmp.Checked <> blnVar Then
                    blnFind = True
                    Exit Do
                End If
                Set nodTmp = nodTmp.Next
            Loop
            '�ҵ��빴ѡֵ�����Ľ�㣬�����и����ȡ����ѡ
            If blnFind Then
                NodeChecked nodVar.Parent, False, False
            Else
                NodeChecked nodVar.Parent, blnVar, False
            End If
        End If
    End If
    
End Sub

Private Sub txtFind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindData(txtFind(Index).Text, Index, False)
    ElseIf KeyCode = vbKeyF3 Then
    End If
End Sub

Private Sub FindData(ByVal strText As String, ByVal intObject As Integer, Optional ByVal blnNext As Boolean = False)
'���ܣ�����intObject��������ƥ�������
'������
'  strText��Ҫ���ҵ��ı�
'  intObject��������
'  blnNext��������һ��ƥ������

    Dim i As Long, j As Long
    Dim blnFind As Boolean
    
    If Trim(strText) = "" Then Exit Sub
    
    If blnNext = False Or mintRecno <= 0 Then
        mintRecno = 1
    End If

    strText = Trim(strText)
    
    Select Case intObject
        Case 0, 1   '0-���ң�1-ҽ��
            Dim lvwVar As ListView
            Dim intTmp As Integer
        
            If intObject = 1 Then
                Set lvwVar = lvwDoctor
            Else
                Set lvwVar = lvwDept
            End If
            With lvwVar
                If .ListItems.Count <= 1 Then Exit Sub
                If mintRecno > .ListItems.Count Then
                    If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call FindData(strText, intObject)
                    End If
                    Exit Sub
                End If
                For i = mintRecno To .ListItems.Count
                    For j = 1 To .ColumnHeaders.Count - 1
                        If .ListItems(i).SubItems(j) Like "*" & strText & "*" Then
                            .ListItems(i).Selected = True
                            .ListItems(i).EnsureVisible         'ʹѡ���Ҳ��ɼ�����Ŀ�ɼ�
                            mintRecno = i + 1
                            blnFind = True
                            .SetFocus
                            Exit For
                        End If
                    Next
                    If blnFind Then Exit For
                Next
                If blnFind = False Then
                    If mintRecno > 1 Then
                        If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                            Call FindData(strText, intObject)
                            Exit Sub
                        End If
                    ElseIf mintRecno = 1 Then
                        Call MsgBox("δ���ҵ�ƥ������ݣ�", vbInformation, gstrSysName)
                    End If
                End If
            End With
        Case 2      '���
            With vsfDiagnose
                If .Rows <= 2 Then Exit Sub
                If mintRecno > .Rows - 1 Then
                    If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call FindData(strText, intObject)
                    End If
                    Exit Sub
                End If
                For i = mintRecno To .Rows - 1
                    For j = 0 To .Cols - 1
                        If .ColHidden(j) = False Or .ColWidth(j) > 0 Then
                            If .TextMatrix(i, j) Like "*" & strText & "*" Then
                                .Row = i
                                .TopRow = i
                                mintRecno = i + 1
                                blnFind = True
                                .SetFocus
                                Exit For
                            End If
                        End If
                    Next
                    If blnFind Then Exit For
                Next
                If blnFind = False Then
                    If mintRecno > 1 Then
                        If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                            Call FindData(strText, intObject)
                            Exit Sub
                        End If
                    ElseIf mintRecno = 1 Then
                        Call MsgBox("δ���ҵ�ƥ������ݣ�", vbInformation, gstrSysName)
                    End If
                End If
            End With
        
            vsfDiagnose.SetFocus
        Case 3      'ҩƷ
            With tvwDrug
                If .Nodes.Count <= 1 Then Exit Sub
                If mintRecno > .Nodes.Count Then
                    If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call FindData(strText, intObject)
                    End If
                    Exit Sub
                End If
                For i = mintRecno To .Nodes.Count
                     If .Nodes(i).Text Like "*" & strText & "*" Then
                        .Nodes(i).Selected = True
                        mintRecno = i + 1
                        blnFind = True
                        .SetFocus
                        Exit For
                     End If
                Next
                If blnFind = False Then
                    If mintRecno > 1 Then
                        If MsgBox("�Ѳ��ҵ��ײ����Ƿ��ͷ�������ң�", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
                            Call FindData(strText, intObject)
                            Exit Sub
                        End If
                    ElseIf mintRecno = 1 Then
                        Call MsgBox("δ���ҵ�ƥ������ݣ�", vbInformation, gstrSysName)
                    End If
                End If
            End With
    End Select
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'""", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsfDiagnose_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> vsfDiagnose.ColIndex("����")
End Sub

Private Sub vsfDiagnose_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsDiagnose As ADODB.Recordset
    Dim lngRow As Long
    Dim strSel As String
    Dim blnFind As Boolean
    
    With vsfDiagnose
        '��ȡ�Ѵ��ڵı��룬��ѡ����ʹ��
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                strSel = strSel & .TextMatrix(lngRow, .ColIndex("����")) & ","
            End If
        Next
        If strSel <> "" Then strSel = "," & strSel
    End With
    
    '����ͳһ�ļ���ѡ����
    Set rsDiagnose = FS.ShowILLSelect(Me, IIf(optDiagnose(1).Value = False, "1,2", "D,Y,M,B"), 0, , True, optDiagnose(1).Value, strSel)
    
    If Not rsDiagnose Is Nothing Then
        With rsDiagnose
            vsfDiagnose.Redraw = False
            
            Do While .EOF = False
                '����Ƿ����
                blnFind = False
                For lngRow = 1 To vsfDiagnose.Rows - 1
                    If Val(vsfDiagnose.TextMatrix(lngRow, vsfDiagnose.ColIndex("ID"))) = Val(zlCommFun.NVL(.Fields("��ĿID").Value)) _
                        And Val(zlCommFun.NVL(.Fields("��ĿID").Value)) > 0 Then
                        blnFind = True
                        Exit For
                    End If
                Next
                                
                If blnFind = False Then
                    '׷��
                    lngRow = vsfDiagnose.Rows - 1
                    If vsfDiagnose.TextMatrix(lngRow, vsfDiagnose.ColIndex("����")) <> "" Then
                        vsfDiagnose.Rows = vsfDiagnose.Rows + 1
                        lngRow = vsfDiagnose.Rows - 1
                    End If
                    
                    vsfDiagnose.TextMatrix(lngRow, vsfDiagnose.ColIndex("ID")) = zlCommFun.NVL(.Fields("��ĿID").Value)
                    vsfDiagnose.TextMatrix(lngRow, vsfDiagnose.ColIndex("����")) = zlCommFun.NVL(.Fields("����").Value)
                    vsfDiagnose.TextMatrix(lngRow, vsfDiagnose.ColIndex("����")) = zlCommFun.NVL(.Fields("����").Value)
                End If
                
                .MoveNext
            Loop
            If blnFind = False Then vsfDiagnose.Rows = vsfDiagnose.Rows + 1
            
            vsfDiagnose.Redraw = True
        End With
    End If
End Sub

Private Sub vsfDiagnose_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With vsfDiagnose
            If .Rows <= 2 Then
                .Rows = 2
                .Clear 1
            Else
                .RemoveItem .Row
            End If
        End With
    End If
End Sub

Private Sub vsfDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    Dim strID As String
    Dim nodTmp As Node
    
    If KeyCode = vbKeyDelete Then
        'ɾ�����������ѡɾ��
        With vsfDrug
            For lngRow = .SelectedRows - 1 To 0 Step -1
                If .SelectedRow(lngRow) >= 0 Then
                    strID = .TextMatrix(.SelectedRow(lngRow), .ColIndex("ID"))
                    'ȡ����㹴ѡ
                    For Each nodTmp In tvwDrug.Nodes
                       If nodTmp.Key = strID Then
                           nodTmp.Checked = False
                           '����NodeCheck�¼�
                           tvwDrug_NodeCheck nodTmp
                           Exit For
                       End If
                    Next
                 End If
            Next
        End With
    End If
End Sub

