VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRModelMan 
   Caption         =   "�������Ĺ���"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmEPRModelMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicFile 
      BorderStyle     =   0  'None
      Height          =   4785
      Left            =   75
      ScaleHeight     =   4785
      ScaleWidth      =   2565
      TabIndex        =   8
      Top             =   705
      Width           =   2565
      Begin XtremeReportControl.ReportControl rptFile 
         Height          =   4800
         Left            =   0
         TabIndex        =   11
         Top             =   405
         Width           =   2445
         _Version        =   589884
         _ExtentX        =   4313
         _ExtentY        =   8467
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   645
         TabIndex        =   9
         Top             =   15
         Width           =   1725
      End
      Begin VB.Label lblFind 
         Caption         =   "����(&V)"
         Height          =   405
         Left            =   0
         TabIndex        =   10
         Top             =   30
         Width           =   945
      End
   End
   Begin VB.PictureBox picNote 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2745
      ScaleHeight     =   345
      ScaleWidth      =   7515
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   7515
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��: "
         Height          =   180
         Left            =   90
         TabIndex        =   6
         Top             =   75
         Width           =   540
      End
   End
   Begin VB.PictureBox picTerm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   7815
      ScaleHeight     =   3150
      ScaleWidth      =   2445
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2445
      Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
         Height          =   2895
         Left            =   15
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   2340
         _cx             =   4128
         _cy             =   5106
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         OwnerDraw       =   0
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   2715
      ScaleHeight     =   3555
      ScaleWidth      =   4770
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4770
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4215
         Left            =   915
         TabIndex        =   7
         Top             =   195
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7435
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   240
         Top             =   2955
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":0B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":10BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":1458
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":1D32
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRModelMan.frx":20CC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vgdList 
         Height          =   900
         Left            =   3930
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2715
         Visible         =   0   'False
         Width           =   1080
         _cx             =   1905
         _cy             =   1587
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRModelMan.frx":2466
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15346
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
   Begin MSComctlLib.ImageList imgFile 
      Left            =   360
      Top             =   5655
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
            Picture         =   "frmEPRModelMan.frx":2CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":3292
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":382C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":3DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":4360
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":48FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelMan.frx":4E94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1395
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   315
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRModelMan.frx":542E
      Left            =   930
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRModelMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const con_UnDefine = -999
Private Enum mPan
    File = 201
    Note = 202
    List = 203
    Term = 204
    View = 205
End Enum
Private Enum mFCol
    ͼ�� = 0: ID: ����: ���: ����: ����: ����
End Enum
Private Enum mLCol
    ͼ�� = 0: ����: ID: ����: ���: ����: ����: ˵��: ����: ��Ա
End Enum

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mstrKinds As String     '��ǰ������Ĳ������ʹ�
Private mintPower As Integer    'ʾ������Ȩ��Χ
'    mintPower=con_UnDefine��δ����;
'    mintPower=-1�����߱�����Ȩ;
'    mintPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    mintPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    mintPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���

Private mlngFileID As Long      '��ǰ�ļ�ID
Private mblnShowAll As Boolean
Private WithEvents mfrmContent As frmEPRFileContent     '������ٴ���
Attribute mfrmContent.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR
'-----------------------------------------------------
'������������ڷ�����ģ��ı���ʹ��
'-----------------------------------------------------
Private mlng��� As Long
Private mstr���� As String

'-----------------------------------------------------
'��������������ڿ���������λ����
'-----------------------------------------------------
Private mblnFindTag As Boolean      '�����򽹵��ж�
Private mintLastRows As Integer     '�������λ��λ��

'-----------------------------------------------------
'����Ϊ���幫������
'-----------------------------------------------------
Public Sub RefreshList()
    '���ܣ�ˢ�µ�ǰ�ĵ������ݣ������ĵ����󱣴�ʱִ�е�ˢ�´���
Dim lngItemID As Long
Dim lngCount As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
    Else
        lngItemID = Me.rptList.FocusedRow.Record(mLCol.ID).Value
    End If
    lngCount = zlRefresh(mlngFileID, lngItemID)
    Me.stbThis.Panels(2).Text = "�����ļ���" & lngCount & "��ʾ��"
End Sub

Public Function zlRefFile(Optional lngFileID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
Dim strGroups As String
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    mlng��� = 0: mstr���� = ""
    
    gstrSQL = "Select Id, ����, ���, ����, ˵��,����" & vbNewLine & _
            "From �����ļ��б�" & vbNewLine & _
            "Where ���� In (" & mstrKinds & ") And Nvl(����, 0) > = 0 And (���� = 7 And ͨ�� > 0 Or ���� <> 7" & IIf(mblnShowAll, "", " And ͨ�� > 0") & ")"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Me.rptFile.Tag = ""
    Me.rptFile.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            If InStr(1, strGroups, !����) = 0 Then strGroups = strGroups & "," & !����
            Set rptRcd = Me.rptFile.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!����)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            Select Case !����
            Case 1: rptRcd.AddItem CStr("1-���ﲡ��")
            Case 2: rptRcd.AddItem CStr("2-סԺ����")
            Case 3: rptRcd.AddItem CStr("3-�����¼")
            Case 4: rptRcd.AddItem CStr("4-������")
            Case 5: rptRcd.AddItem CStr("5-����֤������")
            Case 6: rptRcd.AddItem CStr("6-֪���ļ�")
            Case 7: rptRcd.AddItem CStr("7-���Ʊ���")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem Val(CStr(!���))
            rptRcd.AddItem CStr(!����)
            rptRcd.AddItem NVL(!����, 0)
            rptRcd.AddItem zl9ComLib.zlStr.PinYinCode(CStr(!����))
            rptRcd.Tag = CStr("" & !˵��)
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptFile
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mFCol.����)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    If lngFileID <> 0 Then
        For Each rptRow In Me.rptFile.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mFCol.ID).Value) = lngFileID Then
                    Set Me.rptFile.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptFile.Rows.Count > 0 Then
        If Me.rptFile.FocusedRow Is Nothing Then Set Me.rptFile.FocusedRow = Me.rptFile.Rows(0)
        If Me.rptFile.FocusedRow.GroupRow Then
            lngFileID = 0
        Else
            lngFileID = Me.rptFile.FocusedRow.Record.Item(mFCol.ID).Value
        End If
    Else
        lngFileID = 0
    End If
    
    zlRefFile = Me.rptFile.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefFile = Me.rptFile.Records.Count
    lngFileID = 0
End Function

Public Function zlRefresh(ByVal lngFileID As Long, Optional ByVal lngDemoId As Long) As Long
    '���ܣ�ˢ��װ��ָ���ļ���ʾ��Ŀ¼
    '������ lngFileId���ļ�ID
    '       lngDemoID����Ҫ��λ����ʾ��
    '���أ�ˢ��װ���ʾ����Ŀ
Dim rsTemp As New ADODB.Recordset
Dim objItem As ReportRecordItem
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    Me.Tag = "zlRefresh"
    Err = 0: On Error GoTo errHand
    Select Case mintPower
    Case 0
        gstrSQL = "Select l.Id, l.���, l.����, l.����, Nvl(l.����,'δ����') As ����,l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա,Decode(l.����,Null,1,2) As ���� " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id And l.��Աid = p.Id And l.�ļ�id =[1] " _
                & "Order By Decode(l.����,Null,1,2),l.����,l.���"
    Case 1
        gstrSQL = "Select l.Id, l.���, l.����, l.����, Nvl(l.����,'δ����') As ����,l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա,Decode(l.����,Null,1,2) As ���� " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id(+) And l.��Աid = p.Id(+) And l.�ļ�id =[1] And " _
                & "      (Nvl(l.ͨ�ü�, 0) = 0 Or " _
                & "      l.ͨ�ü� in (1,2) And l.����id In (Select r.����id From ������Ա r, �ϻ���Ա�� u Where r.��Աid = u.��Աid And u.�û��� = User)) " _
                & "Order By Decode(l.����,Null,1,2),l.����,l.���"
    Case Else
        gstrSQL = "Select l.Id, l.���, l.����, l.����, Nvl(l.����,'δ����') As ����,l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա,Decode(l.����,Null,1,2) As ���� " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id(+) And l.��Աid = p.Id(+) And l.�ļ�id =[1] And " _
                & "      (Nvl(l.ͨ�ü�, 0) = 0 Or " _
                & "      l.ͨ�ü� =1 And l.����id In (Select r.����id From ������Ա r, �ϻ���Ա�� u Where r.��Աid = u.��Աid And u.�û��� = User) Or " _
                & "      l.ͨ�ü� =2 And l.��Աid In (Select u.��Աid From �ϻ���Ա�� u Where u.�û��� = User)) " _
                & "Order By Decode(l.����,Null,1,2),l.����,l.���"
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    
    Me.rptList.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CInt(IIf(IsNull(rsTemp!ͨ�ü�), 0, rsTemp!ͨ�ü�))): rptItem.Icon = rptItem.Value
        Set rptItem = rptRcd.AddItem(CInt(Val("" & rsTemp!����))): rptItem.Icon = IIf(rptItem.Value = 0, 4, IIf(rptItem.Value = 1, 5, 3))
        rptRcd.AddItem CStr(rsTemp!ID)
                
        Set objItem = rptRcd.AddItem(Val(rsTemp!����) & CStr(rsTemp!����))
        objItem.Caption = CStr(rsTemp!����)
                        
        rptRcd.AddItem ZLCommFun.NVL(rsTemp!���)
        rptRcd.AddItem CStr(rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!˵��)
        rptRcd.AddItem CStr("" & rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!��Ա)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        For Each rptRow In Me.rptList.Rows
            If Not (rptRow.Record Is Nothing) Then

                If lngDemoId = rptRow.Record(mLCol.ID).Value Then Set Me.rptList.FocusedRow = rptRow: Exit For
            
            End If
        Next
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.Tag = ""
    Call rptList_SelectionChanged
    zlRefresh = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefresh = Me.rptList.Records.Count
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL

    If Me.rptList.Records.Count = 0 Then Exit Sub
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    If Me.rptFile.FocusedRow Is Nothing Then
        strSubhead = ""
    ElseIf Me.rptFile.FocusedRow.GroupRow Then
        strSubhead = ""
    Else
        strSubhead = Me.rptFile.FocusedRow.Record(mFCol.����).Value
    End If
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = strSubhead & "ʾ��Ŀ¼"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

'-----------------------------------------------------
'����Ϊ����ؼ�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngDemoId As Long
Dim cbrControl As CommandBarControl
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_ExportToXML:
        '������XML�ļ�
        If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
        If Me.rptFile.FocusedRow.GroupRow = True Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        
        Dim strF As String
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        '��ͨסԺ����
        dlgThis.Filename = "ʾ��_" & Me.rptFile.FocusedRow.Record.Item(mFCol.����).Value & "_" & Me.rptList.FocusedRow.Record.Item(mLCol.����).Value & ".xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        Err = 0: On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        Err = 0: On Error GoTo 0
        On Error GoTo errHand
        strF = dlgThis.Filename
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If rptList.FocusedRow.Record(mLCol.����).Value = 2 Then '���ʽ�༭��
            mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_ȫ��ʾ���༭, lngDemoId, False, 0
            If mObjTabEpr.zlExportXML(strF) Then
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument
            DocXML.InitEPRDoc cprEM_�޸�, cprET_ȫ��ʾ���༭, lngDemoId
            DocXML.KeepRTF = True
            DocXML.OpenEPRDoc DocXML.frmEditor.Editor1
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_File_ExportToXMLs
        frmModelExportOrImport.ShowMe Me, 1
    Case conMenu_File_ImportFromXMLs
        frmModelExportOrImport.ShowMe Me, 2
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_NewItem
        If mlngFileID = 0 Then Exit Sub
        lngDemoId = frmEPRModelEdit.ShowMe(Me, True, CByte(mintPower), mlngFileID, 0, rptFile.FocusedRow.Record.Item(mFCol.����).Value)
        If lngDemoId <> 0 Then
            Call Me.zlRefresh(mlngFileID, lngDemoId)
            Me.stbThis.Panels(2).Text = "�����ļ�����" & Me.rptList.Rows.Count & "��ʾ��"
        End If
    Case conMenu_Edit_Modify
        If mlngFileID = 0 Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.Record Is Nothing Then Exit Sub
        
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        lngDemoId = frmEPRModelEdit.ShowMe(Me, False, CByte(mintPower), mlngFileID, lngDemoId)
        If lngDemoId <> 0 Then Call Me.zlRefresh(mlngFileID, lngDemoId)
    Case conMenu_Edit_Delete
        Dim lngIndex As Long, strMsg As String
        With Me.rptList
            If .FocusedRow Is Nothing Then Exit Sub
            strMsg = "���ɾ����ʾ����" & vbCrLf & "����" & .FocusedRow.Record(mLCol.����).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_��������Ŀ¼_delete('" & .FocusedRow.Record(mLCol.ID).Value & "')"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Err = 0: On Error GoTo 0
            lngIndex = .FocusedRow.Record.Index
            Call .Records.RemoveAt(.FocusedRow.Record.Index)
            .Populate
            If .Records.Count <> 0 Then
                If lngIndex >= .Records.Count Then lngIndex = 0
                lngDemoId = .Records(lngIndex).Item(mLCol.ID).Value
            Else
                lngDemoId = 0
            End If
            Call Me.zlRefresh(mlngFileID, lngDemoId)
            Me.stbThis.Panels(2).Text = "�����ļ�ʣ��" & Me.rptList.Rows.Count & "��ʾ��"
        End With
    Case conMenu_Edit_Compend
        If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
        If Me.rptFile.FocusedRow.GroupRow = True Then Exit Sub
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        lngDemoId = Me.rptList.FocusedRow.Record(mLCol.ID).Value
        If rptList.FocusedRow.Record(mLCol.����).Value = 2 Then '���ʽ�༭��
            On Error GoTo errHand
            mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_ȫ��ʾ���༭, lngDemoId
        Else
            Dim Doc As New cEPRDocument
            Doc.InitEPRDoc cprEM_�޸�, cprET_ȫ��ʾ���༭, lngDemoId
            Doc.ShowEPREditor Me
        End If
    Case conMenu_Edit_Request
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        lngDemoId = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        If frmEPRModelRequest.ShowMe(Me, lngDemoId, mintPower) = True Then Call rptList_SelectionChanged
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Option
        mblnShowAll = Not mblnShowAll
        Control.Checked = mblnShowAll
        Call zlRefFile
    Case conMenu_View_LocationItem
        txtFind.SetFocus
    Case conMenu_View_Refresh
        Call zlRefFile(mlngFileID)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        'ִ�з�������ǰģ��ı���
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If mstr���� <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                    "�ļ�ID=" & mlngFileID, "���=" & mlng���, "����=" & mstr����)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    If mblnFindTag = True Then
        txtFind.ForeColor = vbBlack
        If txtFind.Text = "���������ƻ�ƴ������" Then txtFind.Text = ""
    Else
        If txtFind.Text = "" Then txtFind.ForeColor = vbGrayText: txtFind.Text = "���������ƻ�ƴ������"
    End If
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = (mintPower >= 0)
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptFile.Records.Count <> 0)
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0)
        Control.Enabled = (mlngFileID <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
    
        Control.Visible = (mintPower >= 0)
        
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mLCol.ͼ��).Value >= mintPower)
        End If

    Case conMenu_Edit_Compend
        Control.Visible = (mintPower >= 0)
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record.Item(mLCol.ͼ��).Value >= mintPower)
        End If
    Case conMenu_File_ExportToXML
        Control.Enabled = True
        If Me.rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
            Control.Enabled = False
        Else
            If Control.Enabled Then Control.Enabled = Not Me.rptFile.FocusedRow.GroupRow
            If Control.Enabled Then Control.Enabled = Not (Me.rptList.FocusedRow Is Nothing)
        End If
        
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Option: Control.Checked = mblnShowAll
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case mPan.File
        Item.Handle = Me.PicFile.hWnd
    Case mPan.Note
        Item.Handle = Me.picNote.hWnd
    Case mPan.List
        Item.Handle = Me.picList.hWnd
    Case mPan.Term
        Item.Handle = Me.picTerm.hWnd
    Case mPan.View
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    If InStr(1, mstrPrivs, "���ﲡ������") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "סԺ��������") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "����������") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "����֤�����淶��") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "֪���ļ�����") > 0 Then mstrKinds = mstrKinds & ",6"
    If InStr(1, mstrPrivs, "���Ʊ��淶��") > 0 Then mstrKinds = mstrKinds & ",7"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    mblnShowAll = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ShowAll", False)
    
    Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    If InStr(1, gstrPrivsEpr, "ȫԺ��������") <> 0 Then
        mintPower = 0
    ElseIf InStr(1, gstrPrivsEpr, "���Ҳ�������") <> 0 Then
        mintPower = 1
    ElseIf InStr(1, gstrPrivsEpr, "���˲�������") <> 0 Then
        mintPower = 2
    Else
        mintPower = -1
    End If
    
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = ZLCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXMLs, "�������������ļ�(&E)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportFromXMLs, "�������뷶���ļ�(&I)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "����(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����(&Q)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "����(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "��ʾδ�ò���(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("F"), conMenu_View_LocationItem
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("D"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("R"), conMenu_Edit_Request
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '---------------------------------------------------------------
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panThis As Pane, panNote As Pane, panView As Pane, panList As Pane, panTerm As Pane
    If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
    
    Set panThis = dkpMan.CreatePane(mPan.File, 150, 480, DockLeftOf, Nothing)
    panThis.Title = "�ļ��б�": panThis.Options = PaneNoCaption
    Set panNote = dkpMan.CreatePane(mPan.Note, 500, 25, DockRightOf, Nothing)
    panNote.Title = "�ļ�˵��": panNote.Options = PaneNoCaption
    Set panView = dkpMan.CreatePane(mPan.View, 500, 240, DockBottomOf, panNote)
    panView.Title = "��������": panView.Options = PaneNoCaption
    Set panList = dkpMan.CreatePane(mPan.List, 400, 240, DockTopOf, panView)
    panList.Title = "�����б�": panList.Options = PaneNoCaption
    Set panTerm = dkpMan.CreatePane(mPan.Term, 120, 240, DockRightOf, panList)
    panTerm.Title = "Ӧ������": panTerm.Options = PaneNoCaption
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptFile
        Set rptCol = .Columns.Add(mFCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mFCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mFCol.����, "����", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mFCol.���, "���", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mFCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mFCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mFCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgFile
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mLCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mLCol.����, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mLCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mLCol.����, "����", 0, False): rptCol.Editable = False: rptCol.Groupable = False:  rptCol.Visible = False: rptCol.Sortable = False
        Set rptCol = .Columns.Add(mLCol.���, "���", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mLCol.����, "����", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.����, "����", 60, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.˵��, "˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mLCol.����, "����", 70, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mLCol.��Ա, "������", 50, False): rptCol.Editable = False: rptCol.Groupable = True
        
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
        
        .SetImageList Me.imgList

        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mLCol.����)
        .GroupsOrder(0).SortAscending = True
        
        .SortOrder.Add .Columns.Find(mLCol.���)
        
    End With
    
    '-----------------------------------------------------
    '��ѯ���ʼ��
    mblnFindTag = False
    txtFind.ForeColor = vbGrayText
    txtFind.Text = "���������ƻ�ƴ������"
    mintLastRows = 0
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '����װ��
    If mstrKinds = "" Then
        DoEvents
        Me.stbThis.Panels(2).Text = "�㲻�߱��κ������ʾ������Ȩ��"
    Else
        lngCount = Me.zlRefFile()
        Me.stbThis.Panels(2).Text = "����" & lngCount & "���ļ�"
    End If
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Set panThis = Me.dkpMan.FindPane(mPan.File)
    panThis.MinTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, 0
    panThis.MaxTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, 0
    panThis.MaxTrackSize.SetSize 3300 / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    
    Set panThis = Me.dkpMan.FindPane(mPan.Note)
    panThis.MinTrackSize.SetSize 0, 345 / Screen.TwipsPerPixelY
    panThis.MaxTrackSize.SetSize panThis.MaxTrackSize.Width, 345 / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Set mfrmContent = Nothing
    Set mObjTabEpr = Nothing
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ShowAll", mblnShowAll
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
Dim cbrControl As CommandBarControl
    If mlngFileID = 0 Then Exit Sub
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Compend)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub PicFile_Resize()
    lblFind.Move 70, 90, lblFind.Width, lblFind.Height
    If PicFile.Width > 800 Then txtFind.Move 800, 50, PicFile.Width - 800, 300
    If PicFile.Height > 400 Then rptFile.Move 0, 400, PicFile.Width, PicFile.Height - 400
End Sub

Private Sub piclist_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = 0: .Width = Me.picList.ScaleWidth
        .Top = 0: .Height = Me.picList.ScaleHeight
    End With
End Sub

Private Sub picTerm_Resize()
    Err = 0: On Error Resume Next
    With Me.vfgTerm
        .Left = 0: .Width = Me.picTerm.ScaleWidth
        .Top = 0: .Height = Me.picTerm.ScaleHeight
        .AutoSize 0
    End With
End Sub

Private Sub rptFile_SelectionChanged()
Dim lngCount As Long
    With Me.rptFile
        If .FocusedRow Is Nothing Then
            mlngFileID = 0: Me.lblNote.Caption = "˵��:"
        ElseIf .FocusedRow.GroupRow = True Then
            mlngFileID = 0: Me.lblNote.Caption = "˵��:"
        Else
            mlngFileID = .FocusedRow.Record.Item(mFCol.ID).Value
            Me.lblNote.Caption = "˵��: " & .FocusedRow.Record.Tag
        End If
        mlng��� = 0: mstr���� = ""
    End With
    If Me.rptFile.FocusedRow Is Nothing Then Exit Sub
    If Me.rptFile.Tag = "" Or Val(Me.rptFile.Tag) <> Me.rptFile.FocusedRow.Index Then
        lngCount = zlRefresh(mlngFileID)
        Me.rptFile.Tag = Me.rptFile.FocusedRow.Index
        If lngCount = 0 Then
            mfrmContent.edtThis.ForceEdit = True
            mfrmContent.edtThis.ReadOnly = False
            mfrmContent.edtThis.NewDoc
            mfrmContent.edtThis.ReadOnly = True
            mfrmContent.edtThis.ForceEdit = False
        End If
        Me.stbThis.Panels(2).Text = "�����ļ���" & lngCount & "��ʾ��"
    End If
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mLCol.���))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup

    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub rptList_SelectionChanged()
    Dim lngItemID As Long
    Dim rsTemp As New ADODB.Recordset
    
    '������⴦����̵��µ�ѡ������仯��������ˢ�¹��̣���ֱ���˳�
    If Me.Tag <> "" Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then
        lngItemID = 0
        mlng��� = 0
        mstr���� = ""
        Call mfrmContent.zlRefresh(0, cprEmCPKModelEssay)
    ElseIf Me.rptList.FocusedRow.Record Is Nothing Then
    
        lngItemID = 0
        mlng��� = 0
        mstr���� = ""
        Call mfrmContent.zlRefresh(0, cprEmCPKModelEssay)
    Else
    
        lngItemID = Me.rptList.FocusedRow.Record.Item(mLCol.ID).Value
        mlng��� = Val(Me.rptList.FocusedRow.Record.Item(mLCol.���).Value)
        mstr���� = Me.rptList.FocusedRow.Record.Item(mLCol.����).Value
    End If
    
    'ˢ������
    Call mfrmContent.zlRefresh(lngItemID, cprEmCPKModelEssay)
    
    'ˢ������
    Err = 0: On Error GoTo errHand
    Me.vfgTerm.Clear: Me.vfgTerm.Rows = Me.vfgTerm.FixedRows
    Set Me.vfgTerm.Cell(flexcpPicture, Me.vfgTerm.FixedRows - 1, 0) = Me.imgList.ListImages(4).Picture
    gstrSQL = "Select ���� As ������, ���� As ����ֵ" & vbNewLine & _
            "From Table(Cast(f_Segment_������([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
            "Where ���� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngItemID)
    With rsTemp
        If .RecordCount <= 0 Then
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "��ʹ������������"
        Else
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "��������������ʱ����ʹ�ã�"
        End If
        Do While Not .EOF
            Me.vfgTerm.Rows = Me.vfgTerm.Rows + 1
            Me.vfgTerm.TextMatrix(Me.vfgTerm.Rows - 1, 0) = Space(2) & Me.vfgTerm.Rows - 1 & ")" & !������ & "Ϊ'" & Replace(!����ֵ, vbTab, "'��'") & "'"
            .MoveNext
        Loop
    End With
    Me.vfgTerm.AutoSize 0
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_Change()
    mintLastRows = 0
End Sub

Private Sub txtFind_GotFocus()
    mblnFindTag = True
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intCount As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        For intCount = mintLastRows + 1 To Me.rptFile.Rows.Count - 1
            If Me.rptFile.Rows(intCount).GroupRow = False Then
                If InStr(Me.rptFile.Rows(intCount).Record(mFCol.����).Value, txtFind.Text) Or InStr(Me.rptFile.Rows(intCount).Record(mFCol.����).Value, UCase(txtFind.Text)) Then
                    Set Me.rptFile.FocusedRow = Me.rptFile.Rows(intCount)
                    mintLastRows = intCount
                    Exit For
                End If
            End If
        Next
        If intCount = Me.rptFile.Rows.Count And mintLastRows = 0 Then
            Call MsgBox("δ�ҵ��롰" & txtFind.Text & "��ƥ��ķ��ģ��������������ƻ���롣", vbInformation, gstrSysName)
            txtFind.Text = ""
        End If
    End If
    txtFind.SetFocus
End Sub

Private Sub txtFind_LostFocus()
    mblnFindTag = False
End Sub
