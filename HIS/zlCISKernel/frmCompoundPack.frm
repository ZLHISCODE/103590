VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCompoundPack 
   Caption         =   "��Һ�������"
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15210
   Icon            =   "frmCompoundPack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15210
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picExecuted 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   4560
      ScaleHeight     =   5415
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   4440
      Width           =   13095
      Begin VSFlex8Ctl.VSFlexGrid vsgExecUnpack 
         Height          =   4860
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   8505
         _cx             =   15002
         _cy             =   8572
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCompoundPack.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
      Begin MSComCtl2.DTPicker dpkExecuted 
         Height          =   300
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   126222339
         CurrentDate     =   40945.9999884259
      End
      Begin MSComCtl2.DTPicker dpkExecuted 
         Height          =   300
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   126222339
         CurrentDate     =   40945
      End
      Begin VB.Label lblInfo 
         Caption         =   "~"
         Height          =   135
         Index           =   10
         Left            =   2880
         TabIndex        =   5
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Caption         =   "ִ��ʱ��"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.PictureBox picWaitExecute 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   360
      ScaleHeight     =   9135
      ScaleWidth      =   14655
      TabIndex        =   6
      Top             =   600
      Width           =   14655
      Begin VB.PictureBox picWaitExecAdvice 
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   3000
         ScaleHeight     =   8895
         ScaleWidth      =   11535
         TabIndex        =   7
         Top             =   0
         Width           =   11535
         Begin VSFlex8Ctl.VSFlexGrid vsgWaitUnpack 
            Height          =   4860
            Left            =   0
            TabIndex        =   8
            Top             =   720
            Width           =   8505
            _cx             =   15002
            _cy             =   8572
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
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCompoundPack.frx":68ED
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
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
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
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
      Begin VB.Frame fraPatiInfo 
         Height          =   9015
         Left            =   0
         TabIndex        =   9
         Top             =   -80
         Width           =   2895
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   6495
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   2655
            _Version        =   589884
            _ExtentX        =   4683
            _ExtentY        =   11456
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picFitter 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   120
            ScaleHeight     =   1455
            ScaleWidth      =   2655
            TabIndex        =   15
            Top             =   7440
            Width           =   2655
            Begin VB.CheckBox chk��Ч 
               Caption         =   "����"
               Height          =   180
               Index           =   1
               Left            =   1800
               TabIndex        =   17
               Top             =   1117
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.CheckBox chk��Ч 
               Caption         =   "����"
               Height          =   180
               Index           =   0
               Left            =   960
               TabIndex        =   16
               Top             =   1117
               Value           =   1  'Checked
               Width           =   735
            End
            Begin MSComCtl2.DTPicker dpkReqTime 
               Height          =   300
               Index           =   0
               Left            =   720
               TabIndex        =   18
               Top             =   330
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   126222339
               CurrentDate     =   40945
            End
            Begin MSComCtl2.DTPicker dpkReqTime 
               Height          =   300
               Index           =   1
               Left            =   720
               TabIndex        =   19
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd HH:mm"
               Format          =   126222339
               CurrentDate     =   40945.9999884259
            End
            Begin VB.Label lblInfo 
               Caption         =   "��Ч"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   23
               Top             =   1117
               Width           =   495
            End
            Begin VB.Label lblInfo 
               Caption         =   "��"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   22
               Top             =   750
               Width           =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "��"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   21
               Top             =   360
               Width           =   255
            End
            Begin VB.Label lblInfo 
               Caption         =   "ִ��ʱ��"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   50
               Width           =   975
            End
         End
         Begin VB.Frame fraBaby 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   240
            TabIndex        =   11
            Top             =   7200
            Visible         =   0   'False
            Width           =   2600
            Begin VB.OptionButton optBaby 
               Caption         =   "����"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   14
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "����ҽ��"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "Ӥ��"
               Height          =   180
               Index           =   2
               Left            =   1815
               TabIndex        =   12
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.Label lblInfo 
            Caption         =   "��ǰ������"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   2500
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":6988
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":6F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":74BC
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":780E
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":E070
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":148D2
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompoundPack.frx":14E6C
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   9975
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   15015
      _Version        =   589884
      _ExtentX        =   26485
      _ExtentY        =   17595
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   9870
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   635
      SimpleText      =   $"frmCompoundPack.frx":15406
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCompoundPack.frx":1544D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21749
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCompoundPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PatiCol
    COL_����ID = 0
    COL_��ҳID = 1
    COL_ѡ�� = 2
    COL_���� = 3
    COL_���� = 4
    COL_�Ա� = 5
    COL_סԺ�� = 6
End Enum

Private Enum AdviceCol
    colѡ�� = 0
    COL��λ = 1
    col���� = 2
    col�Ա� = 3
    col��Ч = 4
    colҽ������ = 5
    col���� = 6
    col���� = 7
    COL��ҩ;�� = 8
    
    colִ��ʱ�� = 9
    col��ҩ���� = 10
    col��ҩ����ʱ�� = 11
    colƿǩ�� = 12
    col״̬ = 13
    Col����ʱ�� = 14
    col���������� = 15
    
    colҽ��ID = 16
    col���ID = 17
    col������� = 18
    Col����ID = 19
    COL��ҳID = 20
    COLƵ�� = 21
    col���ͺ� = 22
    col��Ժ = 23
    col��ҩID = 24
End Enum

Private mlng����ID As Long
Private mlng����ID As Long
Private mintҽ������Χ As Integer    'ҽ������Χ   0-����ҽ��,1-����ҽ��,2-Ӥ��ҽ��
Private mlngҽ������ID As Long
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long
Private mrsDefine As New Recordset
Private mbln��ҩ���ܸ�״̬ As Boolean

Public Sub ShowMe(ByVal intType As Integer, ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, Optional ByVal lngҽ������ID As Long, _
    Optional ByVal lngӤ������ID As Long, Optional ByVal lngӤ������ID As Long)
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlngҽ������ID = lngҽ������ID
    mlngӤ������ID = lngӤ������ID
    mlngӤ������ID = lngӤ������ID
    Me.lblInfo(0).Caption = "��ǰ������" & Sys.RowValue("���ű�", IIF(mlngӤ������ID <> 0 And mlngӤ������ID = mlngҽ������ID, lngӤ������ID, lng����ID), "����")
    
    Me.Show intType, frmParent
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    tbcSub.SetFocus
    Select Case Control.ID
    
    Case conMenu_View_Refresh
        If tbcSub.Selected.Tag = "�����ҽ��" Then
            Call LoadAdvice
        Else
            Call LoadAdvice(True)
        End If
    Case conMenu_Edit_Save
        Call FuncUnpack
    Case conMenu_Manage_Undone
        Call FuncCancleUnpack
    Case conMenu_File_Exit
        Unload Me
    
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
        
    'TabControl
    tbcSub.Left = lngLeft
    tbcSub.Top = lngTop
    tbcSub.Width = Me.Width
    tbcSub.Height = Me.Height - stbThis.Height - 560 - lngTop
    
    
       
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_Manage_Undone
        If tbcSub.Selected.Tag = "�����ҽ��" Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
    Case conMenu_Edit_Save
        If tbcSub.Selected.Tag = "�����ҽ��" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    End Select
End Sub

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч(0).value = 0 And chk��Ч(1).value = 0 Then
        chk��Ч(Index).value = 1
    End If
End Sub

Private Sub FuncCancleUnpack()
'���ܣ�ȡ��ִ��
    Dim arrSQL() As Variant
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim blnTrans As Boolean
    Dim strCurDate As String
    Dim strIDs As String, rsTmp As Recordset, strSQL As String
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    arrSQL = Array()
    With vsgExecUnpack
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, colѡ��) = "1" Then
                
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_��Һ��ҩ��¼_Update(" & .TextMatrix(i, col��ҩID) & ",Null," & _
                    ZVal(.Cell(flexcpData, i, col��ҩ����)) & ",'" & UserInfo.���� & "'," & strCurDate & ")"
                    strIDs = strIDs & "," & .TextMatrix(i, col��ҩID)
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = "select ID from ��Һ��ҩ��¼ where �Ƿ�����=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ��������ҩ��¼�Ѿ�����Һ��ҩ������������ʱ���������ȡ�����,�ѽ���Щ��¼ȡ����ѡ��", vbInformation, "��Һ��Һ��¼"
            Screen.MousePointer = 0
            For i = 1 To vsgExecUnpack.Rows - 1
                rsTmp.Filter = "ID=" & Val(vsgExecUnpack.TextMatrix(i, col��ҩID))
                If rsTmp.RecordCount > 0 Then
                    vsgExecUnpack.Row = i
                    Call ExecCheck(vsgExecUnpack)
                End If
            Next
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    If vsgExecUnpack.TextMatrix(1, colҽ��ID) = "" Then
        stbThis.Panels(2).Text = "û�п�ȡ�������ҽ����"
    Else
        If UBound(arrSQL) = -1 Then
            MsgBox "�빴ѡ����Ҫȡ�������ҽ����", vbInformation, Me.Caption
            Exit Sub
        End If
        stbThis.Panels(2).Text = "ȡ���ɹ������ι�ȡ������� " & UBound(arrSQL) + 1 & " ��ҽ����"
    End If
    Call LoadAdvice(True)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncUnpack()
'���ܣ����
    Dim arrSQL() As Variant
    Dim i As Long
    Dim blnTrans As Boolean
    Dim strCurDate As String
    Dim strIDs As String, rsTmp As Recordset, strSQL As String
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrSQL = Array()
    With vsgWaitUnpack
        For i = 1 To .Rows - 1
            If .TextMatrix(i, colҽ��ID) <> "" And .RowData(i) = "Begin" And .Cell(flexcpData, i, colѡ��) = "1" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_��Һ��ҩ��¼_Update(" & .TextMatrix(i, col��ҩID) & "," & 1 & "," & _
                    ZVal(.Cell(flexcpData, i, col��ҩ����)) & ",'" & UserInfo.���� & "'," & strCurDate & ")"
                strIDs = strIDs & "," & .TextMatrix(i, col��ҩID)
            End If
        Next
    End With
    
    Screen.MousePointer = 11
    On Error GoTo errH
    If strIDs <> "" Then
        strSQL = "select ID from ��Һ��ҩ��¼ where �Ƿ�����=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ��������ҩ��¼�Ѿ�����Һ��ҩ������������ʱ��������д��,�ѽ���Щ��¼ȡ����ѡ��", vbInformation, "��Һ��Һ��¼"
            For i = 1 To vsgWaitUnpack.Rows - 1
                rsTmp.Filter = "ID=" & Val(vsgWaitUnpack.TextMatrix(i, col��ҩID))
                If rsTmp.RecordCount > 0 Then
                    vsgWaitUnpack.Row = i
                    Call ExecCheck(vsgWaitUnpack)
                End If
            Next
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    On Error GoTo 0
    Screen.MousePointer = 0
    If vsgWaitUnpack.TextMatrix(1, colҽ��ID) = "" Then
        stbThis.Panels(2).Text = "û�пɴ����ҩƷ��"
    Else
        stbThis.Panels(2).Text = "����ɹ������ι������ " & UBound(arrSQL) + 1 & " ��ҩƷҽ����"
    End If
    Call LoadAdvice
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        If tbcSub.Selected.Tag = "�����ҽ��" Then
            LoadAdvice
        Else
            LoadAdvice True
        End If
    ElseIf KeyCode = vbKey1 And Shift = 4 Then
        tbcSub.Item(0).Selected = True
    ElseIf KeyCode = vbKey2 And Shift = 4 Then
        tbcSub.Item(1).Selected = True
    End If
End Sub

Private Sub Form_Load()
    Dim strHead As String
    Dim strTbc As String
    
    mbln��ҩ���ܸ�״̬ = Val(zlDatabase.GetPara("��Һ����ҩ���ٴ�������ı���״̬", glngSys, 1345, 0)) = 1
    
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        strTbc = "���"
        .InsertItem(0, "��" & strTbc & "ҽ��(&1)", picWaitExecute.hwnd, 0).Tag = "�����ҽ��"
        .InsertItem(1, "��" & strTbc & "ҽ��(&2)", picExecuted.hwnd, 0).Tag = "�Ѵ��ҽ��"
        
        .Item(0).Selected = True
    End With
    'commandbar
    '-----------------------------------------------------
    Call InitCommandBar
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    'VSFlexGrid
    '-----------------------------------------------------
    strHead = ",400,1;��λ,750,1;����,850,1;�Ա�,450,1;��Ч,450,1;ҽ������,2500,1;����;����,700,1;��ҩ;��;ִ��ʱ��,1550,1;��ҩ����,1470,1;��ҩ����ʱ��,1550,1;ƿǩ��,1980,1;״̬,900,1;����ʱ��,1550,1;����������;ҽ��ID;���ID;�������;����ID;��ҳID;Ƶ��;���ͺ�;��Ժ;��ҩID"

    Call InitTable(vsgWaitUnpack, strHead)
    
    Call InitTable(vsgExecUnpack, strHead)

    Set mrsDefine = InitAdviceDefine
    Call InitPageData
    Call LoadPatiInfo
    
    Call RestoreWinState(Me, App.ProductName)
    
    If DeptIsWoman(0, Get����IDs(IIF(mlngӤ������ID <> 0 And mlngӤ������ID = mlngҽ������ID, mlngӤ������ID, mlng����ID))) Then
        fraBaby.Visible = True
        'ҽ������Χ
        mintҽ������Χ = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
        optBaby(mintҽ������Χ).value = True
    End If
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
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
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " ��ѯ(&Q)"): objControl.BeginGroup = True
        objControl.ToolTipText = "��ȡ�����/�Ѵ��������"

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, " ���(&S)")
        objControl.BeginGroup = True
        objControl.ToolTipText = "���Ѿ���ѡ��ҽ�����д����"
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&C)")
        objControl.ToolTipText = "���Ѿ���ѡ��ҽ������ȡ������Ĳ�����"
        objControl.IconId = 3651
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&E)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyEscape, conMenu_File_Exit
    End With

End Sub

Private Sub LoadAdvice(Optional ByVal blnIsUnpack As Boolean)
'���ܣ�����ҽ��
'������blnIsUnpack=true�����Ѵ��ҽ��,falseΪ���ش����ҽ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim lngID As Long       '���ڶ�λ
    Dim strFormat As String
    Dim strTmp As String
    Dim strFitter As String
    Dim strPatis As Variant
    Dim blnDo As Boolean, blnSetup As Boolean
    Dim lngCount As Long   '��Ҫִ�е�ҽ����(һ����ҩ��1��)
    Dim rsState As Recordset  '��Һ��ҩ״̬
    Dim strIDs As String
    Dim arrIDs() As Variant '����һ���ַ�������

    
    strSQL = "Select " & _
            " a.Id, b.���ͺ�,b.id as ��ҩID,a.���id, a.�������,a.��ʼִ��ʱ��, B.����, p.��ǰ���� As ����, B.�Ա�, Decode(Nvl(a.ҽ����Ч, 0), 0, '����', '����') As ��Ч,a.ҽ��״̬,p.��Ժ," & vbNewLine & _
            "       Decode(a.��������, Null, Null, decode(sign(1-A.��������),1,'0'||A.��������,A.��������) || c.���㵥λ) As ����,  Decode(a.���id,Null,a.ҽ������ || ' ' || a.ִ��Ƶ��  ,a.ҽ������) as ҽ������, to_char(b.ִ��ʱ��,'YYYY-MM-DD HH24:MI') as ִ��ʱ��,p.��ǰ����ID" & _
            ", a.ִ��Ƶ�� As Ƶ��, a.����id, a.��ҳid, a.������Ŀid,c.��������,c.ִ�з���,Decode(a.�ܸ�����, Null, Null," & _
            "  Round(a.�ܸ����� / Decode(a.������Դ, 2, d.סԺ��װ, d.�����װ), 5) || Decode(a.������Դ, 2, d.סԺ��λ, d.���ﵥλ)) As ����,to_char(E.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��," & _
            "b.��ҩ����,g.��ҩʱ�� as ��ҩ����ʱ��,b.ƿǩ��,Decode(b.����״̬,1, '����ҩ',2, '����ҩ', 3,'����ҩ', 4,'����ҩ', '�ѷ���') As ״̬,'' AS ����������" & vbNewLine & _
            " From ��Һ��ҩ��¼ B,����ҽ������ E, ����ҽ����¼ A,������ҳ F, ������Ϣ P, ������ĿĿ¼ C, ҩƷ��� D,��ҩ�������� G" & vbNewLine & _
            " Where F.����ID=P.����ID And F.��ҳID = P.��ҳID And b.ҽ��ID=e.ҽ��ID And b.���ͺ�=e.���ͺ� And g.����(+)=b.��ҩ���� and g.��������id(+)=b.����id And (a.Id = b.ҽ��id Or a.���id = b.ҽ��id) And p.����id = a.����id And a.������Ŀid = c.Id And a.�շ�ϸĿid = d.ҩƷid(+) And a.������� Not In('C','7') And Not (a.�������='E' And c.��������='3') " & _
            Decode(mintҽ������Χ, 1, " And nvl(a.Ӥ��,0) = 0 ", 2, " And nvl(a.Ӥ��,0) <> 0 ", "") & _
            " And (F.Ӥ������ID is null or F.Ӥ������ID is not null and (F.Ӥ������ID=[5] or F.Ӥ������ID=[5]) and NVL(A.Ӥ��,0)<>0 or F.Ӥ������ID is not null and (F.Ӥ������ID<>[5] and f.Ӥ������ID<>[5]) and NVL(A.Ӥ��,0)=0) "
            
    '����
    If Not blnIsUnpack Then
        For i = 0 To rptPati.Rows.Count - 1
            If Not rptPati.Rows(i).GroupRow Then
                If rptPati.Rows(i).Record.Tag = "1" Then
                    strSQL = strSQL & IIF(strPatis = "", " And(", " Or") & " a.����ID =" & rptPati.Rows(i).Record(COL_����ID).value
                    strPatis = strPatis & "," & rptPati.Rows(i).Record(COL_����ID).value
                End If
            End If
        Next
        strPatis = Mid(strPatis, 2)
        If strPatis = "" Then
            MsgBox "��ѡ����Ҫ��ѯ�Ĳ��ˡ�", vbInformation, Me.Caption
            Exit Sub
        End If
        strSQL = strSQL & " )"
        strSQL = strSQL & " And nvl(b.�Ƿ���,0)<>1  And b.ִ��ʱ�� between [1] and [2] "
        strSQL = strSQL & " And (" & IIF(chk��Ч(0).value, "Nvl(a.ҽ����Ч, 0)=0" & IIF(chk��Ч(1).value, " Or Nvl(a.ҽ����Ч, 0)=1", ""), IIF(chk��Ч(1).value, "Nvl(a.ҽ����Ч, 0)=1", "")) & ")"
        '�Ѿ����ʺ��Ѿ���ҩ�Ĳ������������� ��ҩ���ܴ��
        If mbln��ҩ���ܸ�״̬ Then
            strSQL = strSQL & " And b.����״̬=1 "
        Else
            strSQL = strSQL & " And b.����״̬ in(1,2,3) "
        End If
        vsgWaitUnpack.Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("AllCheck").Picture
        vsgWaitUnpack.Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
        vsgWaitUnpack.ColData(colѡ��) = "Check"
        strSQL = strSQL & " Order By p.��ǰ����,B.����, b.ִ��ʱ��,Nvl(a.���id, a.Id),a.id,a.���"
    Else
        '�Ѿ������
        strSQL = strSQL & " and f.��ǰ����id+0=[5] And nvl(b.�Ƿ���,0)=1 And b.ִ��ʱ�� between [3] and [4]  "
        vsgExecUnpack.Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("UnCheck").Picture
        vsgExecUnpack.Cell(flexcpPictureAlignment, 0, colѡ��) = flexPicAlignCenterCenter
        vsgExecUnpack.ColData(colѡ��) = ""
        strSQL = strSQL & " Order By B.���ʱ��,p.��ǰ����,B.����, b.ִ��ʱ��,Nvl(a.���id, a.Id),a.id,a.���"
    End If
    
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(dpkReqTime(0).value), CDate(dpkReqTime(1).value), CDate(dpkExecuted(0).value), CDate(dpkExecuted(1).value), mlngҽ������ID)
    
    i = 0
    strSQL = ""
    ReDim Preserve arrIDs(i)
    Do While Not rsTmp.EOF
        If Len(arrIDs(i) & "," & rsTmp!ID) >= 4000 Then
            i = i + 1
            ReDim Preserve arrIDs(i)
        End If
        arrIDs(i) = arrIDs(i) & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop

    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For j = 0 To i
        arrIDs(j) = Mid(arrIDs(j), 2)
        If arrIDs(j) <> "" Then
            strSQL = strSQL & "Select ��ҩID,��������,������Ա,����ʱ�� from ��Һ��ҩ״̬ Where ��ҩID in(select Column_Value From Table(Cast(f_num2list([" & j + 1 & "]) As ZLTOOLS.t_numlist)))"
        End If
        If j < i Then
            strSQL = strSQL & " union all "
        End If
    Next
    If strSQL <> "" Then
        Set rsState = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, arrIDs)
    End If

    With IIF(blnIsUnpack, vsgExecUnpack, vsgWaitUnpack)
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                If .ColData(colѡ��) = "Check" Then
                    .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, colѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
                End If
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, col��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, colҽ��ID) = rsTmp!ID & ""
                .TextMatrix(i, col���ID) = rsTmp!���ID & ""
                .TextMatrix(i, col�Ա�) = rsTmp!�Ա� & ""
                .TextMatrix(i, COL��λ) = rsTmp!���� & ""
                .TextMatrix(i, Col����ID) = rsTmp!����ID & ""
                .TextMatrix(i, COL��ҳID) = rsTmp!��ҳID & ""
                .TextMatrix(i, col�������) = rsTmp!������� & ""
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, col���ͺ�) = rsTmp!���ͺ� & ""
                .TextMatrix(i, COLƵ��) = rsTmp!Ƶ�� & ""
                .TextMatrix(i, Col����ʱ��) = rsTmp!����ʱ�� & ""
                .TextMatrix(i, colִ��ʱ��) = rsTmp!ִ��ʱ�� & ""
                .TextMatrix(i, col��ҩ����) = "��" & rsTmp!��ҩ���� & "��"
                .Cell(flexcpData, i, col��ҩ����) = Val(rsTmp!��ҩ���� & "")
                .TextMatrix(i, col��ҩ����ʱ��) = rsTmp!��ҩ����ʱ�� & ""
                .TextMatrix(i, col״̬) = rsTmp!״̬ & ""
                .TextMatrix(i, colƿǩ��) = rsTmp!ƿǩ�� & ""
                rsState.Filter = "��ҩID=" & rsTmp!ID & " And ��������=9"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, col����������) = rsState!������Ա & ""
                End If

                .TextMatrix(i, col��Ժ) = rsTmp!��Ժ & ""
                .TextMatrix(i, col��ҩID) = rsTmp!��ҩID & ""

                .RowData(i) = IIF(.TextMatrix(i, col���ID) = "", "Begin", "")
                '��ʾ���ģʽ�µ�ҽ������
                strFormat = rsTmp!ҽ������
                If .TextMatrix(i, COLƵ��) <> "һ����" Then
                    blnDo = True
                    If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                    If blnDo Then
                        strTmp = .TextMatrix(i, col����)
                        If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                    End If
                End If
                .TextMatrix(i, colҽ������) = strFormat
                '�ɱ༭����ɫ
                .Cell(flexcpBackColor, i, colѡ��, i, colѡ��) = COLEditBackColor
                
                '��Ҫִ�е�ҽ������
                If rsTmp!���ID & "" = "" Then lngCount = lngCount + 1
                
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
        If blnIsUnpack Then
            stbThis.Panels(2).Text = "���� " & lngCount & " ��ҽ���Ѿ������"
        Else
            stbThis.Panels(2).Text = "���� " & lngCount & " ��ҽ�����Դ����"
        End If
        '�Զ������и�
        .AutoSize colҽ������
        .Redraw = flexRDDirect
        '�ָ�ǰ��ɫ
        .Cell(flexcpForeColor, 1, colѡ��, .Rows - 1, colѡ��) = vbBlack
        If blnIsUnpack Then
            dpkReqTime(0).value = dpkExecuted(0).value
            dpkReqTime(1).value = dpkExecuted(1).value
        Else
            dpkExecuted(0).value = dpkReqTime(0).value
            dpkExecuted(1).value = dpkReqTime(1).value
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatiInfo()
'���ܣ����ز����б�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long, lngUnitID As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngSelectRow As Long
        
    On Error GoTo errH
    lngUnitID = mlng����ID
    If mlngӤ������ID <> 0 Then
        If mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID Then
            lngUnitID = mlngӤ������ID
        End If
    End If
    
    str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
            
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, False, False, False)
    With rptPati
        For i = 1 To rsTmp.RecordCount
            If Val(rsTmp!��˱�־ & "") < 1 Or gbyt������˷�ʽ <> 1 Then
                Set objRecord = .Records.Add()
                objRecord.Tag = "0"
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
                Set objItem = objRecord.AddItem(rsTmp!��ҳID & "")
                Set objItem = objRecord.AddItem("")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                Set objItem = objRecord.AddItem(rsTmp!���� & "")
                    objItem.Icon = img16.ListImages.Item(IIF(rsTmp!�Ա� & "" = "��", "Man", "Woman")).Index - 1
                Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
                Set objItem = objRecord.AddItem(rsTmp!סԺ�� & "")
                
                
                '������ɫ
                objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
                For j = 1 To objRecord.Childs.Count - 1
                    objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
                Next
                
                '�ϴ��Ƿ�ѡ��
                If lngUnitID = lng����ID And str����IDs <> "" Then
                    If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 Or str����IDs = "ALL" Then
                        objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
                        objRecord.Tag = "1"
                        lngSelectRow = i
                    End If
                ElseIf rsTmp!����ID = mlng����ID Then
                    objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
                    objRecord.Tag = "1"
                    lngSelectRow = i
                End If
            End If
            rsTmp.MoveNext
        Next
        .Populate
        If lngSelectRow > 0 Then
            Set .FocusedRow = .Rows(lngSelectRow - 1)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False)
        Set objCol = .Columns.Add(COL_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 60, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub InitPageData()
'���ܣ���ʼ������
    Dim curDate As Date
    Dim strSQL As String
    Dim rsTmp As Recordset
    Dim i As Long
    Dim strTmp As String
    On Error GoTo errH
    curDate = zlDatabase.Currentdate
    
    dpkExecuted(0).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkExecuted(1).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    dpkReqTime(0).value = Format(curDate, "yyyy-MM-dd 00:00:00")
    dpkReqTime(1).value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim str����IDs As String
    
    '���汨��������
    str����IDs = ""
    For i = 0 To rptPati.Rows.Count - 1
        If rptPati.Rows(i).Record.Tag = "1" Then
            str����IDs = str����IDs & "," & rptPati.Rows(i).Record(COL_����ID).value
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    If str����IDs <> "" Then
        If UBound(Split(str����IDs, ",")) = 0 And Val(str����IDs) = mlng����ID Then
            Call zlDatabase.SetPara("���Ͳ���", "", glngSys, pסԺҽ������)
        Else
            Call zlDatabase.SetPara("���Ͳ���", mlng����ID & ":" & str����IDs, glngSys, pסԺҽ������)
        End If
    End If

    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optBaby_Click(Index As Integer)
    mintҽ������Χ = Index
End Sub

Private Sub picExecuted_Resize()
    On Error Resume Next
    vsgExecUnpack.Width = picExecuted.Width - 200
    vsgExecUnpack.Height = picExecuted.Height - vsgExecUnpack.Top
    
End Sub

Private Sub picWaitExecAdvice_Resize()
    On Error Resume Next
    vsgWaitUnpack.Top = 0
    vsgWaitUnpack.Height = picWaitExecAdvice.Height - vsgWaitUnpack.Top
    vsgWaitUnpack.Width = picWaitExecAdvice.Width
    
End Sub

Private Sub picWaitExecute_Resize()
    On Error Resume Next
    fraPatiInfo.Height = picWaitExecute.Height + 80
    rptPati.Height = fraPatiInfo.Height - rptPati.Top - picFitter.Height - 100
    fraBaby.Top = rptPati.Top + rptPati.Height + 50
    picFitter.Top = IIF(fraBaby.Visible, fraBaby.Top, rptPati.Top) + IIF(fraBaby.Visible, fraBaby.Height, rptPati.Height) + 50
    picWaitExecAdvice.Height = picWaitExecute.Height
    picWaitExecAdvice.Width = picWaitExecute.Width - picWaitExecAdvice.Left - 300
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptPati.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptPati_RowDblClick(rptPati.SelectedRows(0), rptPati.SelectedRows(0).Record.Item(COL_ѡ��))
        End If
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, x As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptPati.HitTest(x, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(x, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
                            rptPati.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptPati.Columns(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.Count - 1
                            rptPati.Records(i)(COL_ѡ��).Icon = -1
                            rptPati.Rows(i).Record.Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COL_ѡ��).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptPati.Populate
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If .TextMatrix(lngRow, col�������) = "�������" Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(lngRow - 1, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow - 1, colҽ��ID)) <> 0 Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow + 1, col���ID)) <> 0 Or Val(.TextMatrix(lngRow + 1, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.TextMatrix(i, colҽ��ID)) <> Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(i, colҽ��ID)) <> 0 Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.TextMatrix(i, colҽ��ID)) <> Val(.TextMatrix(lngRow, colҽ��ID)) Or Val(.TextMatrix(i, colҽ��ID)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, colҽ��ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        Else
            .RowData(lngRow) = "Begin"
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub
    If Item.Tag = "�����ҽ��" Then
        'Call LoadAdvice
    Else
        Call LoadAdvice(True)
    End If
End Sub

Private Sub vsgExecUnpack_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgExecUnpack.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgExecUnpack.RowData(NewRow) = "Begin" Then
        vsgExecUnpack.Editable = flexEDNone
    Else
        vsgExecUnpack.FocusRect = flexFocusNone
        vsgExecUnpack.Editable = flexEDNone
        vsgExecUnpack.ComboList = ""
    End If
End Sub

Private Sub vsgExecUnpack_Click()
    Dim i As Long
    
    With vsgExecUnpack
        If .MouseCol = colѡ�� And .MouseRow = .FixedRows - 1 Then
            If .TextMatrix(1, colҽ��ID) = "" Then Exit Sub
            If .ColData(colѡ��) = "Check" Then
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("UnCheck").Picture
                .ColData(colѡ��) = ""
            Else
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("AllCheck").Picture
                .ColData(colѡ��) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, colҽ��ID) = "" Then Exit For
                If .ColData(colѡ��) = "Check" And .TextMatrix(i, col����������) = "" And (.TextMatrix(i, col״̬) = "����ҩ" Or .TextMatrix(i, col״̬) = "����ҩ") Then
                    .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, colѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, colѡ��) = Nothing
                    .Cell(flexcpData, i, colѡ��) = 0
                End If
                
            Next
        End If
    End With
End Sub

Private Sub vsgExecUnpack_DblClick()
    With vsgExecUnpack
        If .MouseCol = colѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgExecUnpack_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgExecUnpack_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgExecUnpack
        lngLeft = colѡ��: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsgExecUnpack) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If .TextMatrix(Row, col���ID) = "" Then
            vRect.Top = Bottom - 1 '���IDΪ�յ������ֱ���
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, colѡ��, colѡ��) Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgExecUnpack_KeyPress(KeyAscii As Integer)
    With vsgExecUnpack
        If .Col = colѡ�� And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgExecUnpack)
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsgWaitUnpack.Cell(flexcpBackColor, NewRow, NewCol) = COLEditBackColor And vsgWaitUnpack.RowData(NewRow) = "Begin" Then
        vsgWaitUnpack.FocusRect = flexFocusHeavy
        vsgWaitUnpack.Editable = flexEDNone
    Else
        vsgWaitUnpack.FocusRect = flexFocusNone
        vsgWaitUnpack.Editable = flexEDNone
        vsgWaitUnpack.ComboList = ""
    End If
End Sub

Private Sub vsgWaitUnpack_Click()
    Dim i As Long
    
    With vsgWaitUnpack
        If .MouseCol = colѡ�� And .MouseRow = .FixedRows - 1 Then
            If .TextMatrix(1, colҽ��ID) = "" Then Exit Sub
            If .ColData(colѡ��) = "Check" Then
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("UnCheck").Picture
                .ColData(colѡ��) = ""
            Else
                .Cell(flexcpPicture, 0, colѡ��) = img16.ListImages("AllCheck").Picture
                .ColData(colѡ��) = "Check"
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, colҽ��ID) = "" Then Exit For
                If .ColData(colѡ��) = "Check" Then
                    .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, colѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
                Else
                    Set .Cell(flexcpPicture, i, colѡ��) = Nothing
                    .Cell(flexcpData, i, colѡ��) = 0
                End If
                
            Next
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_DblClick()
    With vsgWaitUnpack
        If .MouseCol = colѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsgWaitUnpack_KeyPress(vbKeySpace)
        End If
    End With
End Sub

Private Sub vsgWaitUnpack_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsgWaitUnpack
        lngLeft = colѡ��: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsgWaitUnpack) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If .TextMatrix(Row, col���ID) = "" Then
            vRect.Top = Bottom - 1 '���IDΪ�յ������ֱ���
            vRect.Bottom = Bottom - 1
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Between(Col, colѡ��, colѡ��) Then
                SetBkColor hDC, OS.SysColor2RGB(COLEditBackColor)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsgWaitUnpack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsgWaitUnpack_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsgWaitUnpack_KeyPress(KeyAscii As Integer)
    With vsgWaitUnpack
        If .Col = colѡ�� And KeyAscii = vbKeySpace Then
            Call ExecCheck(vsgWaitUnpack)
        End If
    End With
End Sub

Private Sub ExecCheck(ByRef objVsg As VSFlexGrid)
'���ܣ�ͬ��ѡ��һ��ҽ��
'���������
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    
    With objVsg
        If .TextMatrix(.Row, colҽ��ID) = "" Then Exit Sub
        If Not RowInһ����ҩ(.Row, lngBegin, lngEnd, objVsg) Then
            lngBegin = .Row: lngEnd = .Row
        End If
        
        For i = lngBegin To lngEnd
            If .Cell(flexcpData, i, colѡ��) = 1 Then
                Set .Cell(flexcpPicture, i, colѡ��) = Nothing
                .Cell(flexcpData, i, colѡ��) = 0
            Else
                If objVsg.Name = "vsgExecUnpack" Then
                    '����Ƿ��Ժ
                    If .TextMatrix(i, col��Ժ) <> "1" Then
                        MsgBox "�ò����Ѿ���Ժ������ȡ�������", vbInformation, Me.Caption
                        Exit Sub
                    ElseIf .TextMatrix(i, col����������) <> "" Then
                        MsgBox "��ҽ���Ѿ����ʣ�����ȡ�������", vbInformation, Me.Caption
                        Exit Sub
                    ElseIf .TextMatrix(i, col״̬) <> "����ҩ" And .TextMatrix(i, col״̬) <> "����ҩ" Then
                        MsgBox "��ҽ���Ѿ���ҩ������ȡ�������", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                .Cell(flexcpPicture, i, colѡ��) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, colѡ��) = 1
                .Cell(flexcpPictureAlignment, i, colѡ��) = flexPicAlignCenterCenter
            End If
        Next
    End With
End Sub



