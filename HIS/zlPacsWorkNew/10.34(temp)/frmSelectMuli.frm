VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmSelectMuli 
   Appearance      =   0  'Flat
   Caption         =   "ѡ������"
   ClientHeight    =   6795
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   14985
   Icon            =   "frmSelectMuli.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   14985
   StartUpPosition =   1  '����������
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picList 
      Height          =   5175
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   7635
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      Begin VB.PictureBox picCommand 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   6135
         TabIndex        =   10
         Top             =   4200
         Width           =   6135
         Begin VB.CommandButton cmdDel 
            Caption         =   "ɾ ��(&D)"
            Height          =   400
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ȡ ��(&C)"
            Height          =   400
            Left            =   3600
            TabIndex        =   12
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ȷ ��(&S)"
            Height          =   400
            Left            =   2400
            TabIndex        =   11
            Top             =   120
            Width           =   1095
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTree 
         Height          =   2055
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   3855
         _cx             =   6800
         _cy             =   3625
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VB.Frame frmFilter 
         Caption         =   "��������"
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   7455
         Begin VB.Frame frmTime 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   7215
            Begin VB.OptionButton optDays 
               Caption         =   "2��"
               Height          =   180
               Index           =   1
               Left            =   840
               TabIndex        =   24
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "����"
               Height          =   180
               Index           =   5
               Left            =   3240
               TabIndex        =   23
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optDays 
               Caption         =   "7��"
               Height          =   180
               Index           =   4
               Left            =   2640
               TabIndex        =   22
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "5��"
               Height          =   180
               Index           =   3
               Left            =   2040
               TabIndex        =   21
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "3��"
               Height          =   180
               Index           =   2
               Left            =   1440
               TabIndex        =   20
               Top             =   240
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optDays 
               Caption         =   "1��"
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   19
               Top             =   240
               Width           =   615
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               Height          =   300
               Left            =   5760
               TabIndex        =   25
               Top             =   195
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   234881027
               CurrentDate     =   40833
            End
            Begin MSComCtl2.DTPicker dtpStart 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "yyyy-MM-dd"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   0
               EndProperty
               Height          =   300
               Left            =   4080
               TabIndex        =   26
               Top             =   195
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   234881027
               CurrentDate     =   40833
            End
            Begin VB.Label Label3 
               Caption         =   "��"
               Height          =   255
               Left            =   5520
               TabIndex        =   27
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Left            =   5880
            TabIndex        =   15
            Top             =   930
            Width           =   1455
         End
         Begin VB.TextBox txtStudyNo 
            Height          =   300
            Left            =   3360
            TabIndex        =   14
            Top             =   930
            Width           =   1455
         End
         Begin VB.ComboBox cboModality 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "��  ����"
            Height          =   255
            Left            =   5040
            TabIndex        =   16
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "�� �� �ţ�"
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Ӱ�����"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picImage 
      Height          =   3255
      Left            =   6840
      ScaleHeight     =   3195
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   2640
      Width           =   7215
      Begin VB.CheckBox chkViewImage 
         Caption         =   "Ԥ��ͼ��"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin zl9PacsControl.ucSplitPage ucPage 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   6210
         _extentx        =   10504
         _extenty        =   582
         pagecount       =   0
         pagerecord      =   9
      End
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2295
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "˫��������ʾ��ͼ"
         Top             =   240
         Width           =   5895
         _Version        =   262147
         _ExtentX        =   10398
         _ExtentY        =   4048
         _StockProps     =   35
         BackColor       =   0
      End
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6510
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13714
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image imgImage 
      Height          =   240
      Left            =   1920
      Picture         =   "frmSelectMuli.frx":0CCA
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSeries 
      Height          =   240
      Left            =   1320
      Picture         =   "frmSelectMuli.frx":10B4
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStudy 
      Height          =   240
      Left            =   720
      Picture         =   "frmSelectMuli.frx":144C
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   360
      Top             =   4080
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSelectMuli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_STUDY_UID As String = "1.2.840.023.500903000005970.67031919.0"
Private Const M_STR_SERIES_UID As String = "1.2.840.023.500903000005970.67031919.1"

Public mblnOK As Boolean

Private mintSelectIndex As Integer

Private mstrTitle As String
Private mlngReleationType As Long       '1--ȡ��������2--����ͼ��
Private mstrModality As String          'Ӱ�����

Private mMultiRows As Integer
Private mMultiCols As Integer

Private mlngModule  As Long         '��ǰվ��ģ��
Private mlngCurDeptId As Long     '��ǰ����ID
Private mlngAdviceID As Long      '��ǰҽ��ID
Private mblnMoved As Boolean        '����Ƿ�ת��
Private mblnSaveReportImage As Boolean  '�Ƿ񱣴汨��ͼ

Private mlngCurPageIndex As Long    '���浱ǰҳ����
Private mlngPageCount As Long       'ÿҳ��ʾ��ͼ������

Private mdcmUID As New DicomGlobal


Private mrsStudyData As ADODB.Recordset
Private mrsSeriesData As ADODB.Recordset
Private mrsImageData As ADODB.Recordset

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long




Public Function ShowImageReleation(ByVal lngModule As Long, ByVal lngAdviceID As Long, ByVal strPrivs As String, _
    ByVal blnMoveId As Boolean, ByVal blnSaveReportImage As Boolean, ByVal lngCurDeptId As Long, _
    Optional lngReleationType As Long = 1, Optional strModality = "")
    
    Dim curDate As Date
    
    mblnOK = False
    
    curDate = zlDatabase.Currentdate
    
    mlngModule = lngModule
    mlngAdviceID = lngAdviceID
    mblnMoved = blnMoveId
    mblnSaveReportImage = blnSaveReportImage
    mlngCurDeptId = lngCurDeptId
    
    mlngReleationType = lngReleationType
    mstrModality = strModality
    
    dtpStart.value = curDate - 2
    dtpEnd.value = curDate
    
    Me.Caption = IIf(mlngReleationType = 1, "ȡ������", "����ͼ��")
    cmdDel.Visible = IIf(mlngReleationType = 1, False, True)
    cmdDel.Enabled = CheckPopedom(strPrivs, "ɾ����ʱӰ��")
    
    '����ǲ���ģ�飬����Ҫ��ʾ��������ͼ��
    If glngModul = G_LNG_PATHOLSYS_NUM Then
        cboModality.Clear
        cboModality.AddItem "DG-����"
        
        cboModality.ListIndex = 0
        
        Label1.Visible = False
        cboModality.Visible = False
        
        frmTime.Left = 120
        frmTime.Width = frmFilter.Width - 240
    Else
        '���Ӱ�����
        Call FillModality
    End If


    Call InitReleationList

    
    'ˢ���б�
    If mlngReleationType = 2 Then
        Call QueryReleationData(dtpStart.value, dtpEnd.value)
        Call FilterReleationData
    Else
        Call QueryCancelReleationData
    End If
    
    Call LoadReleationDataToFace
    
    On Error GoTo 0
    
    Call InitFaceScheme

    Me.Show 1
End Function



Private Sub InitReleationList()
'��ʼ�������б�
    With vsfTree
        
        ' structure
        .Cols = 4
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .Left = 50
        
        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .AllowSelection = False
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
        
        .ColDataType(0) = flexDTBoolean
        .ColWidth(0) = 800
        
        .ColHidden(1) = True
        
        .ColWidth(2) = 1600
        
        
    End With
End Sub


Private Sub QueryCancelReleationData()
'��ѯ��Ҫȡ������������
    Dim strSql As String
    
    '��ѯ��ʱ��¼
    strSql = "select Ӱ�����,to_char(����) as ����,����,Ӣ����,�Ա�,����,���uid,λ��һ,λ��һ,λ��һ,����豸,�������� from Ӱ�����¼ where ҽ��ID=[1]"
    If mblnMoved Then strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    
    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    
    '��ѯӰ�������У�ʹ�ù���ʱ����ѯ�Ż�ʹ��������
    strSql = "select /*+ Rule*/ a.����UID,a.���UID,a.���к�,a.��������,a.�ɼ�ʱ�� from Ӱ�������� a, Ӱ�����¼ b where a.���UID=b.���UID and b.ҽ��ID=[1] order by a.���к�"
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
    End If
    
    Set mrsSeriesData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    
    
    '��ѯӰ����ͼ��
    strSql = "select /*+ Rule*/ a.ͼ��UID, a.����UID, b.���UID, a.ͼ���, a.ͼ������, a.�ɼ�ʱ�� from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID  and c.ҽ��ID=[1] order by a.����UID, a.ͼ���"
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
    End If
    
    Set mrsImageData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
End Sub


'Private Sub QueryReleationData(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
''��ѯ��������
'    Dim strSql As String
'
'    '��ѯ��ʱ��¼
'    strSql = "select /*+ Rule*/ Ӱ�����,to_char(����) as ����,����,Ӣ����,�Ա�,����,���uid,λ��һ,λ��һ,λ��һ,����豸,�������� " & _
'            " from Ӱ����ʱ��¼ where �������� between [1] and [2]"
'    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'
'    '��ѯ��ʱ����
'    strSql = "select /*+ Rule*/ a.����UID,a.���UID,a.���к�,a.��������,a.�ɼ�ʱ�� from Ӱ����ʱ���� a, Ӱ����ʱ��¼ b where a.���UID=b.���uid and b.�������� between [1] and [2] order by ���к�"
'    Set mrsSeriesData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate)
'
'
'    '��ѯ��ʱͼ��
'    strSql = "select /*+ Rule*/ a.ͼ��UID, a.����UID, b.���UID, a.ͼ���, a.ͼ������, a.�ɼ�ʱ�� from Ӱ����ʱͼ�� a, Ӱ����ʱ���� b, Ӱ����ʱ��¼ c  where a.����UID=b.����UID and b.���UID=c.���UID and b.�������� between [1] and [2]   order by a.����UID, a.ͼ���"
'    Set mrsImageData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
'
'End Sub


Private Sub QueryReleationData(ByVal dtStartDate As Date, ByVal dtEndDate As Date)
'��ѯ��������
    Dim strSql As String

    '��ѯ��ʱ��¼
    strSql = "select /*+ Rule*/ Ӱ�����,to_char(����) as ����,����,Ӣ����,�Ա�,����,���uid,λ��һ,λ��һ,λ��һ,����豸,�������� " & _
            " from Ӱ����ʱ��¼ where �������� between [1] and [2]"
    Set mrsStudyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(Format(dtStartDate, "yyyy-mm-dd 00:00")), CDate(Format(dtEndDate, "yyyy-mm-dd 23:59")))

End Sub



Private Sub FilterReleationData()
'���˹�������
    Dim strFilter As String
    

    strFilter = ""
    
    If cboModality.ListIndex >= 0 Then
        strFilter = "Ӱ����� = '" & Split(cboModality.Text, "-")(0) & "'"
    End If
    
    If Trim(Replace(txtStudyNo.Text, "'", "")) <> "" Then
        strFilter = strFilter & " and ����='" & Replace(txtStudyNo.Text, "'", "") & "'"
    End If
    
    If Trim(Replace(txtName.Text, "'", "")) <> "" Then
        strFilter = strFilter & " and ���� like '" & Replace(txtName.Text, "'", "") & "%'"
    End If
    
    
'    '����ʱ���������
'    If Not optDays(5).value And Replace(txtName.Text, "'", "") = "" And Replace(txtStudyNo.Text, "'", "") = "" Then
'        strFilter = strFilter & " And �������� > '" & Format(dtpStart.value, "yyyy-MM-dd 00:00") & "'" & " And �������� < '" & Format(dtpEnd.value, "yyyy-MM-dd 23:59") & "'"
'    End If

    
    mrsStudyData.Filter = strFilter
    
    stb.Panels(1).Text = "�������� " & mrsStudyData.RecordCount & " ���������"
End Sub



Private Sub LoadReleationDataToFace()
'����������ݵ�����
    Dim i As Long
    
    vsfTree.Rows = 0
    Call vsfTree.Clear
    Call DViewer.Images.Clear
    
    If mrsStudyData.RecordCount <= 0 Then Exit Sub
    
    With vsfTree
    
        .Redraw = flexRDNone
        .Rows = 0
        
        '��ȡ���ڵ�
        While Not mrsStudyData.EOF
            .AddItem ""
            
            .RowData(.Rows - 1) = 0
            
            .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '�����ȡ�����������Զ�ѡ��������
            .Cell(flexcpPicture, .Rows - 1, 2) = imgStudy
            
            .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsStudyData!���uid)
            .Cell(flexcpText, .Rows - 1, 2) = Nvl(mrsStudyData!����) & "(" & Nvl(mrsStudyData!����) & ")"
'            .Cell(flexcpFontBold, .Rows - 1, 2) = True
            
            .Cell(flexcpText, .Rows - 1, 3) = "����:" & Nvl(mrsStudyData!����) & "  ����:" & Nvl(mrsStudyData!����) & "  �Ա�:" & Nvl(mrsStudyData!�Ա�) & "  ����:" & Nvl(mrsStudyData!����) & "  �������:" & Nvl(mrsStudyData!��������)
            .Cell(flexcpFontSize, .Rows - 1, 3) = 9
            .Cell(flexcpForeColor, .Rows - 1, 3) = vbGrayText
            
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = 1
            
            If mlngReleationType <> 2 Then
                .RowData(.Rows - 1) = 1
                
                '��ȡ���нڵ�
                mrsSeriesData.Filter = "���UID='" & Nvl(mrsStudyData!���uid) & "'"
                If mrsSeriesData.RecordCount > 0 Then
                    While Not mrsSeriesData.EOF
                        .AddItem ""
    
                        .RowData(.Rows - 1) = 1
    
                        .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '�����ȡ�����������Զ�ѡ����������
                        .Cell(flexcpPicture, .Rows - 1, 2) = imgSeries
    
                        .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsSeriesData!����UID)
                        .Cell(flexcpText, .Rows - 1, 2) = "����" & Nvl(mrsSeriesData!���к�)
    '                    .Cell(flexcpFontBold, .Rows - 1, 2) = True
    
                        .Cell(flexcpText, .Rows - 1, 3) = "���к�:" & Nvl(mrsSeriesData!���к�) & "  ��������:" & Nvl(mrsSeriesData!��������) & "  ��������:" & Nvl(mrsSeriesData!�ɼ�ʱ��)
                        .Cell(flexcpFontSize, .Rows - 1, 3) = 9
                        .Cell(flexcpForeColor, .Rows - 1, 3) = vbGrayText
    
                        .IsSubtotal(.Rows - 1) = True
                        .RowOutlineLevel(.Rows - 1) = 2

    
                        '��ȡͼ��ڵ�
                        mrsImageData.Filter = "����UID='" & Nvl(mrsSeriesData!����UID) & "'"
                        If mrsImageData.RecordCount > 0 Then
                            While Not mrsImageData.EOF
                                .AddItem ""
    
                                .RowData(.Rows - 1) = 1
    
                                .Cell(flexcpChecked, .Rows - 1, 0) = IIf(mlngReleationType = 1, True, False) '�����ȡ�����������Զ�ѡ��ͼ������
                                .Cell(flexcpPicture, .Rows - 1, 2) = imgImage
    
                                .Cell(flexcpText, .Rows - 1, 1) = Nvl(mrsImageData!ͼ��UID)
                                .Cell(flexcpText, .Rows - 1, 2) = "ͼ��" & Nvl(mrsImageData!ͼ���)
                                .Cell(flexcpText, .Rows - 1, 3) = "ͼ���:" & Nvl(mrsImageData!ͼ���) & "  �ɼ�ʱ��:" & Nvl(mrsImageData!�ɼ�ʱ��)
    
                                .Cell(flexcpFontSize, .Rows - 1, 3) = 9
                                .Cell(flexcpForeColor, .Rows - 1, 3) = &HC0C0FF
    
                                .IsSubtotal(.Rows - 1) = True
                                .RowOutlineLevel(.Rows - 1) = 3
    
                                Call mrsImageData.MoveNext
                            Wend
                        End If
    
                        mrsSeriesData.MoveNext
                    Wend
                    
                End If
            End If
                    
            Call mrsStudyData.MoveNext
            
        Wend

        .Outline 1
        
        If .Rows > 0 Then
            .Row = 0
            .RowSel = 0
        End If

        .Redraw = flexRDBuffered
        
        For i = 0 To vsfTree.Rows - 1
            '�۵��ڵ�
            .IsCollapsed(i) = flexOutlineCollapsed
        Next i
    End With
End Sub



Private Sub InitPageControl(ByVal lngSearchType As Long, ByVal strSearchId As String)
'��ʼ����ҳ�ؼ�
    Dim strFilter As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    If mlngReleationType = 2 Then
        Select Case lngSearchType
            Case 1
                strSql = "select count(1)  as ����ֵ from Ӱ����ʱͼ�� a, Ӱ����ʱ���� b where a.����UID=b.����UID and b.���UID=[1]"
            Case 2
                strSql = "select count(1)  as ����ֵ from Ӱ����ʱͼ��  where  ����UID=[1]"
            Case 3
                strSql = "select count(1)  as ����ֵ from Ӱ����ʱͼ��  where  ͼ��UID=[1]"
        End Select
        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSearchId)
        If rsData.RecordCount > 0 Then
            lngRecordCount = Nvl(rsData!����ֵ)
        Else
            lngRecordCount = 0
        End If
    Else
        Select Case lngSearchType
            Case 1
                strFilter = "���UID='" & strSearchId & "'"
            Case 2
                strFilter = "����UID='" & strSearchId & "'"
            Case 3
                strFilter = "ͼ��UID='" & strSearchId & "'"
        End Select
    
        mrsImageData.Filter = strFilter
        lngRecordCount = mrsImageData.RecordCount
    End If
    
    ucPage.RecordCount = lngRecordCount
End Sub



Private Function GetImageViewData(ByVal lngSearchType As Long, _
    ByVal strSearchId As String, ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
'��ȡԤ��ͼ������
'intSearchType:0-�����uid����,1-������UID����,2-��ͼ��UID����

    Dim strSql As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    If mlngReleationType = 2 Then
        '����ͼ��
        strSql = "Select rownum as ˳���,  A.ͼ���,d.FTP�û��� As User1, d.FTP���� As Pwd1, d.Ip��ַ As Host1," & _
                " '/' || d.FtpĿ¼ || '/' As Root1, " & _
                " Decode(C.��������, Null, '', To_Char(C.��������, 'YYYYMMDD') || '/') || C.���uid || '/' || A.ͼ��uid As URL, " & _
                " d.�豸�� As �豸��1,A.ͼ��UID,C.���UID,B.����UID,d.FTP�û��� As User2, d.FTP���� As Pwd2," & _
                " d.Ip��ַ As Host2, '/' || d.FtpĿ¼ || '/' As Root2, " & _
                " d.�豸�� As �豸��2,A.��̬ͼ,A.��������, A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
                " From Ӱ����ʱͼ�� A, Ӱ����ʱ���� B, Ӱ����ʱ��¼ C ,Ӱ���豸Ŀ¼ D " & _
                " Where A.����UID = B.����UID And B.���UID = C.���UID And  C.λ��һ = D.�豸�� "
    Else
        'ȡ������
        
        strSql = "Select rownum as ˳���, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
            "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
            
        If mblnMoved Then
            strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
            strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
        End If
    End If

    Select Case lngSearchType
        Case 1
            strSql = strSql & " and C.���UID=[1]"
        Case 2
            strSql = strSql & " and B.����UID=[1]"
        Case 3
            strSql = strSql & " and A.ͼ��UID=[1]"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSql = "select /*+RULE*/ * from (" & strSql & " order by b.����UID, a.ͼ���) where ˳���>=" & lngStartRecord & " and ˳���<=" & lngEndRecord
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSearchId)
End Function


Private Sub LoadViewImageToFace(rsCurImageData As ADODB.Recordset)
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim iCols As Integer, iRows As Integer
    
    
    
    DViewer.Images.Clear
    
    If rsCurImageData.RecordCount > 0 Then
        '����ͼ����ʾ����
        ResizeRegion rsCurImageData.RecordCount, DViewer.Width, DViewer.Height, iRows, iCols
        
        mMultiCols = iCols
        mMultiRows = iRows

        DViewer.MultiColumns = iCols
        DViewer.MultiRows = iRows
        
        '��������Ŀ¼
        strCachePath = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
        
        Do While Not rsCurImageData.EOF
            'ѭ������ͼ��DicomViewer��
            strTmpFile = strCachePath & Nvl(rsCurImageData("URL"))
            
            If Nvl(rsCurImageData("��̬ͼ"), IMGTAG) = VIDEOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\Avi.bmp", App.Path & "..\�����ļ�\Avi.bmp")
            ElseIf Nvl(rsCurImageData("��̬ͼ"), IMGTAG) = AUDIOTAG Then
                strTmpFile = IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\wav.bmp", App.Path & "..\�����ļ�\wav.bmp")
            End If
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                
                '����FTP����
                If Nvl(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                        If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                MsgBoxD Me, "FTP�����������ӣ������������á�"
                                Exit Sub
                            End If
                        Else
                            MsgBoxD Me, "FTP�����������ӣ������������á�"
                            Exit Sub
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                    '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                    If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                    End If
                End If
            End If
  
            If Dir(strTmpFile) <> vbNullString Then
               If Nvl(rsCurImageData("��̬ͼ"), IMGTAG) <> VIDEOTAG And Nvl(rsCurImageData("��̬ͼ"), IMGTAG) <> AUDIOTAG Then
                    Set curImage = DViewer.Images.ReadFile(strTmpFile)
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    On Error GoTo continue
                        Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                    
                    Call AddVideoLabelToDicomImage(curImage, _
                        "�ɼ�ʱ�䣺" & Nvl(rsCurImageData("�ɼ�ʱ��")), _
                        "¼�Ƴ��ȣ�" & Nvl(rsCurImageData("¼�Ƴ���"), "0") & " ��", _
                        "�������ƣ�" & Nvl(rsCurImageData("��������")))
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    Call DViewer.Images.Add(curImage)
                End If
                
                
                'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                '���½�ú��DSAͼ����������ʾ
                '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
            
            rsCurImageData.MoveNext
        Loop
        
        If DViewer.Images.Count > 0 Then
            DViewer.CurrentIndex = 1
            DViewer.Images(1).BorderColour = vbRed
        End If
        
        Inet1.FuncFtpDisConnect
        Inet2.FuncFtpDisConnect
    Else
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
    End If
End Sub


Private Sub cboModality_Click()
    If Not cboModality.Visible Then Exit Sub
    
    If mlngReleationType = 2 Then '����ͼ��
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call FilterReleationData
        Call LoadReleationDataToFace
    End If
End Sub

Private Sub chkViewImage_Click()
On Error GoTo ErrHandle
    If Not vsfTree.Visible Then Exit Sub
    
    If chkViewImage.value <> 0 Then
        Call vsfTree_SelChange
    Else
        Call DViewer.Images.Clear
    End If
    
    '�������
    Call SetDeptPara(mlngCurDeptId, "Ԥ������ͼ��", chkViewImage.value)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub


Private Function GetReleationImageIds() As ADODB.Recordset
'��ѯ��������Ҫȡ��������ͼ��ID
    Dim i As Long, j As Long
    Dim strSql As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    
    '�����ѯ���
    For i = 0 To vsfTree.Rows - 1
        If vsfTree.RowOutlineLevel(i) = 3 And vsfTree.TextMatrix(i, 0) = -1 Then 'Ϊ3��ʾͼ��ڵ�
            If j > 79 Then
                strFilter = strFilter & " Or ͼ��UID ='" & vsfTree.TextMatrix(i, 1) & "'"
            Else
                If zlCommFun.ActualLen(strValue) > 3600 Then
                     strValues(j) = Mid(strValue, 2)
                     strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                     
                     strValue = ""
                     j = j + 1
                End If
                
                strValue = strValue & "," & vsfTree.TextMatrix(i, 1)
            End If
        End If
    Next i
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '���û����Ҫ���ҵ�ͼ��UID���򷵻ؿ����ݼ�
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
'    If strFilter <> "" Then strFilter = " and ( " & Mid(strFilter, 4) & ")"
    If strFilter <> "" Then strFilter = strUninTable & " Union All Select ͼ��UID from [Ӱ��ͼ��] where  ( " & Mid(strFilter, 4) & ")"
    
    '�����ƶ��ķ���ͬ��Դͼ�п����ڡ�Ӱ����ʱ��¼�����ߡ�Ӱ�����¼����
    '����ʱ����ʱ��¼���Ƶ�������¼��ȡ������ʱ��������¼���Ƶ���ʱ��¼
    strSql = "Select /*+RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ͼ�� A, Ӱ�������� B, Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID = E.ͼ��UID " & _
        "Union All " & _
        "Select /*+RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ʱͼ�� A,Ӱ����ʱ���� B, Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ʱͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID= E.ͼ��UID"
        
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function



Private Sub GetStorageDevice(ByVal lngAdviceID As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������
'�����ȡ����������ʹ��strNewStudyUID�����ܴ����ݿ��в��ҵ���Ӧ������
'strDeviceNum:�豸��
'strFtpIp: ftp��ַ
'strFtpUrl: ftpĿ¼
'strVirtualPath: ftp����洢·��
'strFtpUser: ftp�û���
'strFtpPwd: ftp����



    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSql = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
        
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '���ִ�е����˵����ִ��ͼ�����,��Ҫ�жϵ�ǰ���Ĵ洢�豸�Ƿ���Ч�������Ч�������µĴ洢�豸
        If Trim(rsData!��������) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!λ��һ)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '�����µļ��UID�ʹ洢�豸,���ִ�е����˵����ȡ������
        
        If mlngModule = 1290 Then
            '��ѯҽ������վ�У��������Ӧ�Ĵ洢�豸
            strSql = "select d.����ֵ " & _
                        " from ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��DICOM����� c, Ӱ��DICOM������� d " & _
                        " Where a.����ID = b.ִ�в���id And a.ִ�м� = b.ִ�м� And a.����豸 = c.�豸�� " & _
                        " and c.������='ͼ�����' and c.����ID=d.����ID and d.��������='�洢�豸' and b.ҽ��id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "δ�ҵ�ͼ��洢�豸,��ȷ�ϵ�ǰ��������豸�Ƿ���Ӱ���豸Ŀ¼�ķ���������������ͼ��洢��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = Nvl(rsTemp!����ֵ)
        Else
            '��ѯ��ҽ������վ�е�ͼ��洢�豸
            strDeviceNO = GetDeptPara(mlngCurDeptId, "�洢�豸��")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD Me, "δ�ҵ�ͼ��洢�豸,��ȷ����Ӱ�����̹������Ƿ�Ըÿ���������ͼ��ɼ��洢�豸��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        
        strSql = "Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
                    " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Tag, strDeviceNO)
        
        '����洢�豸ͣ�ã���ֱ���˳�
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD Me, "δ�ҵ��洢�豸,��ȷ���豸��Ϊ [" & strDeviceNO & "] ���豸�Ƿ����á�", vbInformation, gstrSysName
            Exit Sub
        End If
        
'        Call funGetStorageDevice(Me, strDeviceNO, strFtpUrl, strFTPIP, strFTPUser, strFTPPwd)
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP��ַ"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strFTPPwd = Nvl(rsTemp("FTP����"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo ErrHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '����FTPĿ¼
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
ErrHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub


Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'ɾ��ftp�������е��ļ�
    Dim objSrcFtp As New clsFtp
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strImageUID As String
    Dim strVirtualPath As String
'    Dim lngResult As Long
    
    DelTempImages = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
    
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
        
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
    

        'ɾ��ͼ���ļ�����ɾ��ʧ�ܺ����˳�ִ��
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        'ɾ�����ܴ��ڵı���ͼ��
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID & ".jpg")
        
'        If lngResult <> 0 Then
'            Call err.Raise(-1, "MoveImageToStudy", "Ftp������ͼ��ɾ��ʧ�ܡ� [ͼ��UID:" & strImageUID & "]", err.HelpFile, err.HelpContext)
'            Exit Function
'        End If
    
        'ͼ��ɾ���ɹ���ͬ��ɾ�����ݿ��е�����
        Call zlDatabase.ExecuteProcedure("ZL_Ӱ����_ɾ����ʱͼ��(3,'" & strImageUID & "')", Me.Caption)
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend


    objSrcFtp.FuncFtpDisConnect
    
    DelTempImages = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
    
End Function


Public Function MoveImageToStudy(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String, _
    ByVal strFTPIP As String, ByVal strFtpUrl As String, ByVal strFtpVirtualPath As String, _
    ByVal strFTPUser As String, ByVal strFTPPwd As String, ByRef objMoveList As Collection) As Boolean
'------------------------------------------------
'���ܣ���ѡ���ļ��ͼ���ƶ���ftp��ָ���ļ����
'���أ�True--�ɹ���False��ʧ��
'------------------------------------------------
    Dim objSrcFtp As New clsFtp
    Dim objDestFtp As New clsFtp
    Dim strVirtualPath As String
    Dim strDestPath As String
    Dim strTmpFile As String
    Dim aFiles() As String
    Dim i As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim lngResult As Long       '��¼FTP�����Ľ��
    Dim strImageUID As String
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strFileList As String
    Dim blnIsMove As Boolean
    
On Error GoTo ErrHandle
    
    blnIsMove = False
    MoveImageToStudy = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function

    '����Ŀ��Ftp
    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strVirtualPath = ""
    strFileList = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
    
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        
        If strVirtualPath <> Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url) Then
            strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            strFileList = ""
        End If
    
        
        '���ƶ����ļ�������ͬ��ftp��ַʱ����ʹ�����غ����ϴ��ķ�ʽת���ļ�
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
        
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            
                strCurFtpIp = Nvl(rsImageDatas!host)
                strCurFtpUser = Nvl(rsImageDatas!FtpUser)
                strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
                
                Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
            End If
        
            strTmpFile = App.Path & "\TmpImage\" & strImageUID
            
            If strFileList = "" Then
                strFileList = objSrcFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objSrcFtp.FuncDownloadFile(strVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "���ع���ͼ��ʧ�ܡ� [ͼ��UID:" & strImageUID & " �ļ�����Ŀ¼:" & strVirtualPath & " ����·��:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
        
                lngResult = objDestFtp.FuncUploadFile(strFtpVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "�ϴ�����ͼ��ʧ�ܡ� [ͼ��UID:" & strImageUID & " �ϴ�����Ŀ¼:" & strFtpVirtualPath & " ����·��:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
                
                blnIsMove = True
            End If
        Else
            If strFileList = "" Then
                strFileList = objDestFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                If lngResult <> 0 Then
                    '����ļ��ƶ�ʧ�ܣ���˿���������һ��
                    Call objDestFtp.FuncFtpDisConnect
                    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                    
                    lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                    
                    If lngResult <> 0 Then
                        objSrcFtp.FuncFtpDisConnect
                        objDestFtp.FuncFtpDisConnect
                        
                        Call err.Raise(-1, "MoveImageToStudy", "��Ftp���ƶ��ļ�ʱʧ�ܡ� [ͼ��UID:" & strImageUID & " ԭ����Ŀ¼:" & strVirtualPath & " ������Ŀ¼:" & strFtpVirtualPath & "]", err.HelpFile, err.HelpContext)
                        Exit Function
                    End If
                End If
                
                blnIsMove = True
                
                '��¼�Ѿ����ƶ������ļ����Ա��ڴ�������ʧ�ܵ�ʱ�򣬻��ɶ��ƶ���ͼ����лָ�
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strVirtualPath & "/" & strImageUID & ">" & strFtpVirtualPath & "/" & strImageUID)
                End If
            End If
        End If
        

        If mblnSaveReportImage Then
            '�ϴ�ftp�еı���ͼ
            
            If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
                Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 0)
            Else
                Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 1)
            End If
        End If
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend


    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    
    '���һ���ļ���û�б��ƶ�����ֱ���˳�
    If Not blnIsMove Then Exit Function
    
    MoveImageToStudy = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo ErrHandle
'�ƶ�����ͼ
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim lngResult As Long
    
    If lngWay = 0 Then
        Call objSrcFtp.FuncDelFile(strSourceVirtualPath, strImgUid & ".jpg")
        
        '��������д��ڴ�Դftp�����ص�dicomͼ����ͼ��ת����jpg�������浽Ŀ��ftp�豸��
        If FileExists(strDicomFile) Then
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strDicomFile)
    
            Call dcmImg.FileExport(strDicomFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestVirtualPath, strDicomFile & ".jpg", strImgUid & ".jpg")
            
            If FileExists(strDicomFile & ".jpg") Then Call Kill(strDicomFile & ".jpg")
        End If
    Else
        '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
        If objDestFtp.FuncFtpFileExists(strSourceVirtualPath, strImgUid & ".jpg") Then
            lngResult = objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
            
            If lngResult <> 0 Then
                '����ļ��ƶ�ʧ�ܣ���˿���������һ��
                Call objDestFtp.FuncFtpDisConnect
'                Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                Call objDestFtp.ResotreFtpConnect
                
                Call objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
                
                '��¼�Ѿ����ƶ������ļ����Ա��ڴ�������ʧ�ܵ�ʱ�򣬻��ɶ��ƶ���ͼ����лָ�
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strSourceVirtualPath & "/" & strImgUid & ".jpg" & ">" & strDestVirtualPath & "/" & strImgUid & ".jpg")
                End If
            End If
        End If
    End If
Exit Sub
ErrHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub


Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo ErrHandle
'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strNewDirectory
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strNewDirectory = App.Path & "\TmpImage\" & Format(zlDatabase.Currentdate, "YYYYMMDD")
    
    If Not DirExists(strNewDirectory) Then MkDir strNewDirectory
    If Not DirExists(strNewDirectory & "\" & strNewStudyUID) Then MkDir strNewDirectory & "\" & strNewStudyUID
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = App.Path & "\TmpImage\" & Nvl(rsImageDatas!ͼ��UID)
        
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       Ϊ������������ͼ��������ش���ͼ���ļ������ý���ɾ��
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")

        '�ƶ��ļ����µ�λ��
        Call MoveFile(App.Path & "\TmpImage\" & Nvl(rsImageDatas!Url) & "\" & Nvl(rsImageDatas!ͼ��UID), _
            strNewDirectory & "\" & strNewStudyUID & "\" & Nvl(rsImageDatas!ͼ��UID))
                

        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        'ɾ���յ�ftpĿ¼
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
ErrHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub


'����ͼ����ƶ�
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo ErrHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
ErrHandle:
    objFtp.FuncFtpDisConnect
End Sub


'ȡ�ù�����ʾ��Ϣ
Private Function GetReleationHintInfo(lngAdviceID As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strResult As String
    Dim strStudyInf As String
    
    
    GetReleationHintInfo = ""
    
    If rsReleationImage.RecordCount <= 0 Then Exit Function
    
    Call rsReleationImage.MoveFirst
    While Not rsReleationImage.EOF
        strStudyInf = "[" & Nvl(rsReleationImage!����) & "(" & Nvl(rsReleationImage!����) & ") " & Nvl(rsReleationImage!�Ա�) & " " & Nvl(rsReleationImage!����) & "]"
        
        If InStr(strResult, strStudyInf) <= 0 Then
            If strResult <> "" Then strResult = strResult & "+"
        
            strResult = strResult & strStudyInf
        End If
        Call rsReleationImage.MoveNext
    Wend
    
    strSql = "select ����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    
    GetReleationHintInfo = "�Ƿ�ȷ�Ͻ�  " & strResult & "  ��ͼ����  [" & Nvl(rsTemp!����) & "(" & Nvl(rsTemp!����) & ") " & Nvl(rsTemp!�Ա�) & " " & Nvl(rsTemp!����) & "]  �ļ����й���������"
End Function


Private Function StartReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��ʼ����
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As String
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ���ҹ����������... ]"
    
    strSql = "select ���UID,�������� from Ӱ�����¼ where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "�Ҳ����������ļ����Ϣ��", vbInformation, Me.Caption)
        Exit Function
    End If
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ��ȡ�µ�ͼ��洢��Ϣ... ]"
    
    
    
    If Trim(Nvl(rsTmp!���uid)) = "" Or Trim(Nvl(rsTmp!��������)) = "" Then
        
        '��δ�ɼ�ͼ����Ҫ�����µļ��UID
        strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '���´洢�豸��Ϣ
        strSql = "Zl_Ӱ����_�����豸(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!���uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
        
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ�ƶ�ͼ���µ�ͼ��洢λ��... ]"
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
        Exit Function
    End If
    
    
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ����ͼ���������... ]"
    
    
    '��ȡ����ͼ����Ϣ
    strSql = "Select ���UID,����ͼ�� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", lngAdviceID)
    
    strOldReportImages = ""
    lngReportImageLen = 0
    
    If rsReportImage.RecordCount > 0 Then
        strOldReportImages = Nvl(rsReportImage!����ͼ��)
        lngReportImageLen = Len(strOldReportImages)
    End If
        
    '�����µ�����UID
    strNewSeriesUid = CreateSeriesUid(mdcmUID.NewUID)
    
    strReportImageIds = ""
    rsImageDatas.MoveFirst
                
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        '����ͼ���������
        strSql = "Zl_Ӱ����_ͼ�����(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & Nvl(rsImageDatas!ͼ��UID) & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '���汨������
        If mblnSaveReportImage Then
            If InStr(1, strOldReportImages & ";" & strReportImageIds, Nvl(rsImageDatas!ͼ��UID)) <= 0 And Len(strReportImageIds) < 4000 - lngReportImageLen - 60 Then
                If strReportImageIds <> "" Then strReportImageIds = strReportImageIds & ";"
                strReportImageIds = strReportImageIds & Nvl(rsImageDatas!ͼ��UID) & ".jpg"
            End If
        End If
    
        rsImageDatas.MoveNext
    Wend
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ���±���ͼ��Ϣ... ]"
    
    '�����Ҫ���ֱ���ͼ������Ҫ�Ȳ�ѯĿǰ�Ѿ����ֵı���ͼ��UID
    If mblnSaveReportImage Then
        
        If rsReportImage.RecordCount > 0 Then
            strReportImageIds = IIf(strOldReportImages <> "", strOldReportImages & ";", "") & strReportImageIds
            strReportImageIds = Replace(strReportImageIds, ";;", ";")
        End If
        
        strSql = "Zl_Ӱ����_���±���ͼ(" & mlngAdviceID & ",'" & strReportImageIds & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    
    '�ύ����
    Call gcnOracle.CommitTrans
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼɾ����Ч��FTPͼ���ļ�... ]"
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    StartReleation = True
    
    Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '�����׳�����
End Function

Private Function CancelReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��������
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As Long
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ��ȡ�µ�ͼ��洢��Ϣ... ]"
    
    '����ͼ�����
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
    
    Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
        Exit Function
    End If
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ�ƶ�ͼ���µ�ͼ��洢λ��... ]"
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
        Exit Function
    End If
    
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ����ͼ���������... ]"
    
    strSql = "Select ���UID,����ͼ�� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", mlngAdviceID)
    
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = Nvl(rsReportImage!����ͼ��)
        strReportImageIds = Replace(strReportImageIds, " ", "") '�ɼ�ͼ��ʱ�����ܻ��ڱ���ͼ���ݺ���ӿո�
    End If
        
        
    '��������
    rsImageDatas.MoveFirst
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        strSql = "Zl_Ӱ����_��������(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!ͼ��UID) & "','" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
                                        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '�޸ı���ͼ����
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!ͼ��UID) & ".jpg;", "")
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!ͼ��UID) & ".jpg", "")
        
        rsImageDatas.MoveNext
    Wend
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼ���±���ͼ��Ϣ... ]"
    
    '���±���ͼ��
    strSql = "Zl_Ӱ����_���±���ͼ(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call gcnOracle.CommitTrans
    
    stb.Panels(2).Text = "����ִ�У���ȴ�!  [ ��ʼɾ����Ч��FTPͼ���ļ�... ]"
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("CancelReleation", err)
    
    Call RaiseErr(err)
End Function


Public Function ReleationImage() As Boolean
'-----------------------------------------------------------------------------
'����:����ͼ���ƶ�FTPͼ���µ�λ�ã��޸����ݿ��¼������ʱ��ת����ʽ����
'���أ���
'-----------------------------------------------------------------------------
    Dim rsImageDatas As ADODB.Recordset
    Dim strHint As String
    Dim blnResult As Boolean
    
    On Error GoTo ErrHandle
        ReleationImage = False

        
        '�����ݿ��в�ѯͼ������
        Set rsImageDatas = GetReleationImageIds()
    
        If rsImageDatas Is Nothing Then
            Call MsgBoxD(Me, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '��ǰ���UID�����ݿ��в����ڣ����˳�������
        If rsImageDatas.RecordCount <= 0 Then
            Call MsgBoxD(Me, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, Me.Caption)
            Exit Function
        End If
        
        
        If mlngReleationType = 2 Then
            '����ͼ����ʾ
            strHint = GetReleationHintInfo(mlngAdviceID, rsImageDatas)
            
            If strHint = "" Then
                Call MsgBoxD(Me, "���ܲ�ѯ����Ҫ������������Ϣ������������", vbOKOnly, Me.Caption)
                Exit Function
            End If
            
            If MsgBoxD(Me, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            
        Else
            'ȡ��������ʾ
            If MsgBoxD(Me, "�Ƿ�ȷ�϶���ѡͼ�����ȡ������������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        End If

        If mlngReleationType = 2 Then '����2��ʾ����ͼ��
            blnResult = StartReleation(mlngAdviceID, rsImageDatas)
        Else
            blnResult = CancelReleation(mlngAdviceID, rsImageDatas)
        End If
        

        stb.Panels(2).Text = "��ǰ������ִ����ϡ�"
        
        ReleationImage = blnResult
        
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub cmdDel_Click()
'ɾ����ʱͼ��ֻ���ڹ���ͼ��Ĵ����в��ܽ���ɾ������������ʱ������ִ�иò�����
On Error GoTo ErrHandle
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '�����ݿ��в�ѯͼ������
    Set rsImageDatas = GetReleationImageIds()

    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    
    If MsgBoxD(Me, "�Ƿ�ȷ��ɾ����ѡͼ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    
    If DelTempImages(rsImageDatas) Then
        For i = vsfTree.Rows - 1 To 0 Step -1
            
            If vsfTree.TextMatrix(i, 0) = -1 Then
                Call vsfTree.RemoveItem(i)
            Else
                If vsfTree.GetNode(i).Children <= 0 And vsfTree.RowOutlineLevel(i) < 3 And vsfTree.RowData(i) = 1 Then
                    Call vsfTree.RemoveItem(i)
                End If
            End If
        Next i
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHandle
    mblnOK = ReleationImage
    
    If mblnOK Then
        Call MsgBoxD(Me, "��ǰ������ִ����ϡ�", vbInformation, Me.Caption)
        
        stb.Panels(2).Text = ""
        Unload Me
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo ErrHandle
    Cancel = IIf((dkpMain.Panes(1).Hidden Or dkpMain.Panes(2).Hidden) And Action = 8 Or ((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.Hidden), True, False)
ErrHandle:
End Sub


Private Sub dtpEnd_Change()
On Error GoTo ErrHandle
    Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpStart_Change()
On Error GoTo ErrHandle
    Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub DViewer_DblClick()
    If DViewer.Images.Count = 0 Then Exit Sub
    If mintSelectIndex <= 0 Then Exit Sub

    If DViewer.MultiColumns = 1 And DViewer.MultiRows = 1 Then
        DViewer.MultiColumns = mMultiCols
        DViewer.MultiRows = mMultiRows
        DViewer.CurrentIndex = 1
    Else
        mMultiCols = DViewer.MultiColumns
        mMultiRows = DViewer.MultiRows
        DViewer.MultiColumns = 1
        DViewer.MultiRows = 1
        DViewer.CurrentIndex = mintSelectIndex
    End If
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    If DViewer.Images.Count = 0 Then Exit Sub
    
    If Button = 1 And Shift = 0 Then
        mintSelectIndex = DViewer.ImageIndex(X, Y)
        
        If mintSelectIndex <= 0 Then Exit Sub
        
        For i = 1 To DViewer.Images.Count
            DViewer.Images(i).BorderColour = vbWhite
        Next i
        DViewer.Images(mintSelectIndex).BorderColour = vbBlue
    End If
End Sub





Private Sub Form_Activate()
    vsfTree.SetFocus
End Sub


Private Sub FillModality()
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ�������")
    
    cboModality.Clear
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!���� = mstrModality Then cboModality.ListIndex = cboModality.ListCount - 1
        rsTemp.MoveNext
    Loop
    
    If cboModality.ListIndex = -1 Then
        If cboModality.ListCount >= 1 Then
            cboModality.ListIndex = 1
        End If
    End If
End Sub




Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .CloseAll
'        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 150, 200, DockLeftOf, Nothing)
    Pane1.Title = "�����б�"
    Pane1.Handle = PicList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable Or PaneNoCaption
    Pane1.MinTrackSize.Width = 520
    
    
    Set Pane2 = dkpMain.CreatePane(2, 150, 200, DockRightOf, Pane1)
    Pane2.Title = "ͼ��Ԥ��"
    Pane2.Handle = picImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable Or PaneNoCaption
    Pane2.MinTrackSize.Width = 450

End Sub

Private Sub Form_Load()
    '�ָ�����״̬
    Call RestoreWinState(Me, App.ProductName)
    
    chkViewImage.value = GetDeptPara(mlngCurDeptId, "Ԥ������ͼ��", 0)
    ucPage.PageRecord = GetDeptPara(mlngCurDeptId, "Ԥ������ͼ������", 9)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���洰��״̬
    Call SaveWinState(Me, App.ProductName)
    
    Call SetDeptPara(mlngCurDeptId, "Ԥ������ͼ������", Val(ucPage.PageRecord))
End Sub

Private Sub optDays_Click(Index As Integer)
On Error GoTo ErrHandle
    Dim i As Integer
    Dim dtNow As Date
    
    If mlngReleationType = 2 Then '����ͼ��
    
        dtpStart.Enabled = True
        dtpEnd.Enabled = True
        
        dtNow = zlDatabase.Currentdate
                        
        '����ʱ���������
        For i = 0 To 5
            If optDays(i).value = True Then
                Select Case i
                    Case 0
                        dtpStart.value = dtNow
                        dtpEnd.value = dtNow
                    Case 1
                        dtpStart.value = DateAdd("d", -1, dtNow)
                        dtpEnd.value = dtNow
                    Case 2
                        dtpStart.value = DateAdd("d", -2, dtNow)
                        dtpEnd.value = dtNow
                    Case 3
                        dtpStart.value = DateAdd("d", -4, dtNow)
                        dtpEnd.value = dtNow
                    Case 4
                        dtpStart.value = DateAdd("d", -6, dtNow)
                        dtpEnd.value = dtNow
                    Case 5
                        dtpStart.value = DateAdd("d", -14, dtNow)
                        dtpEnd.value = dtNow
'                        dtpStart.Enabled = False
'                        dtpEnd.Enabled = False
                End Select
            End If
        Next i
    
        Call QueryReleationData(dtpStart.value, dtpEnd.value)
        
        Call FilterReleationData
        Call LoadReleationDataToFace
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picCommand_Resize()
On Error GoTo ErrHandle
    cmdCancel.Left = picCommand.ScaleWidth - cmdCancel.Width - 120
    cmdCancel.Top = 60
    
    CmdOK.Left = cmdCancel.Left - CmdOK.Width - 120
    CmdOK.Top = 60
    
    cmdDel.Left = 120
    cmdDel.Top = 60
ErrHandle:
End Sub

Private Sub picImage_Resize()
On Error GoTo ErrHandle
    Dim iCols As Integer, iRows As Integer
    
    DViewer.Left = 0
    DViewer.Top = 0
    DViewer.Width = picImage.ScaleWidth
    DViewer.Height = picImage.ScaleHeight - ucPage.Height - 60 - stb.Height
    
    ucPage.Left = 0
    ucPage.Top = picImage.ScaleHeight - ucPage.Height - stb.Height
    
    chkViewImage.Left = ucPage.Left + ucPage.Width + 120
    chkViewImage.Top = ucPage.Top + 30

    ResizeRegion DViewer.Images.Count, DViewer.Width, DViewer.Height, iRows, iCols
    DViewer.MultiRows = iRows
    DViewer.MultiColumns = iCols
ErrHandle:
End Sub

Private Sub picList_Resize()
On Error GoTo ErrHandle
    frmFilter.Top = PicList.Height - frmFilter.Height - picCommand.Height - 240 - stb.Height
    frmFilter.Width = PicList.ScaleWidth - 180
    
    If mlngReleationType = 1 Then    'ȡ������
        frmFilter.Visible = False
        vsfTree.Height = PicList.ScaleHeight - picCommand.Height - 240 - stb.Height
    ElseIf mlngReleationType = 2 Then    '����ͼ��
        frmFilter.Visible = True
        vsfTree.Height = PicList.ScaleHeight - frmFilter.Height - picCommand.Height - 240 - stb.Height
    End If
    
    vsfTree.Left = 0
    vsfTree.Top = 0
    vsfTree.Width = PicList.ScaleWidth
    
    picCommand.Left = 0
    picCommand.Top = PicList.Height - picCommand.Height - 120 - stb.Height
    picCommand.Width = PicList.ScaleWidth
ErrHandle:
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii <> 13 Then Exit Sub
    
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii <> 13 Then Exit Sub
    
    Call FilterReleationData
    Call LoadReleationDataToFace
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    
    Dim strSearchId As String
    Dim lngSearchType As Long
    
    If Not vsfTree.Visible Then Exit Sub
    
    If chkViewImage.value = 0 Then Exit Sub
    
    strSearchId = vsfTree.TextMatrix(vsfTree.Row, 1)
    lngSearchType = vsfTree.RowOutlineLevel(vsfTree.Row)
    
    Set rsData = GetImageViewData(lngSearchType, strSearchId, lngPageIndex, lngPageCount)
    Call LoadViewImageToFace(rsData)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsfTree_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim lngCurLevel As Long
    Dim i As Long
    
    If Col <> 0 Then Exit Sub
    
    lngCurLevel = vsfTree.RowOutlineLevel(Row)
    

    For i = Row + 1 To vsfTree.Rows - 1
        If vsfTree.RowOutlineLevel(i) <= lngCurLevel Then Exit For
        
        vsfTree.Cell(flexcpChecked, i, 0) = vsfTree.Cell(flexcpChecked, Row, Col)
    Next i
    
    
    i = Row - 1
    While i >= 0
        If vsfTree.RowOutlineLevel(i) < lngCurLevel Then
            If vsfTree.Cell(flexcpChecked, Row, 0) = 2 Then
                vsfTree.Cell(flexcpChecked, i, 0) = False
                lngCurLevel = vsfTree.RowOutlineLevel(i)
            End If
        End If
        
        i = i - 1
    Wend

    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vsfTree_DblClick()
On Error GoTo ErrHandle
    If vsfTree.Rows <= 0 Then Exit Sub
        
    
    If vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineCollapsed Then
        vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineExpanded
    Else
        vsfTree.IsCollapsed(vsfTree.Row) = flexOutlineCollapsed
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadSeriesNode(ByVal lngStudyRow As Long)
'�������нڵ�
    Dim strStudyUID As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnChecked As Boolean
    
    If vsfTree.GetNode(lngStudyRow).Children > 0 Then Exit Sub
    
    vsfTree.RowData(lngStudyRow) = 1
    
    strStudyUID = vsfTree.TextMatrix(lngStudyRow, 1)
    blnChecked = vsfTree.Cell(flexcpChecked, lngStudyRow, 0) = 1
    
    strSql = "select  ����UID, ���UID, ���к�, ��������, �ɼ�ʱ�� from Ӱ����ʱ���� where ���UID=[1] order by ���к� desc"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strStudyUID)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    With vsfTree
        '��ʼ��������
        While Not rsData.EOF
            .AddItem "", lngStudyRow + 1
    
            .RowData(lngStudyRow + 1) = 1
    
            .Cell(flexcpChecked, lngStudyRow + 1, 0) = blnChecked
            .Cell(flexcpPicture, lngStudyRow + 1, 2) = imgSeries
    
            .Cell(flexcpText, lngStudyRow + 1, 1) = Nvl(rsData!����UID)
            .Cell(flexcpText, lngStudyRow + 1, 2) = "����" & Nvl(rsData!���к�)
    
            .Cell(flexcpText, lngStudyRow + 1, 3) = "���к�:" & Nvl(rsData!���к�) & "  ��������:" & Nvl(rsData!��������) & "  ��������:" & Nvl(rsData!�ɼ�ʱ��)
            .Cell(flexcpFontSize, lngStudyRow + 1, 3) = 9
            .Cell(flexcpForeColor, lngStudyRow + 1, 3) = vbGrayText
    
            .IsSubtotal(lngStudyRow + 1) = True
            .RowOutlineLevel(lngStudyRow + 1) = 2
            
            Call LoadImageNode(lngStudyRow + 1)
            
            .IsCollapsed(lngStudyRow + 1) = flexOutlineCollapsed
            
            rsData.MoveNext
        Wend
    End With
End Sub

Private Sub LoadImageNode(ByVal lngSeriesRow As Long)
'����ͼ��ڵ�
    Dim strSeriesUID As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnChecked As Boolean
    
    If vsfTree.GetNode(lngSeriesRow).Children > 0 Then Exit Sub
    
    vsfTree.RowData(lngSeriesRow) = 1
    
    strSeriesUID = vsfTree.TextMatrix(lngSeriesRow, 1)
    blnChecked = vsfTree.Cell(flexcpChecked, lngSeriesRow, 0) = 1
    
    strSql = "select  ͼ��UID, ����UID, ͼ���, ͼ������, �ɼ�ʱ��  from Ӱ����ʱͼ�� where ����UID=[1] order by ͼ��� desc"
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSeriesUID)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    With vsfTree
        '��ʼ����ͼ��ڵ�
        While Not rsData.EOF
            .AddItem "", lngSeriesRow + 1

            .RowData(lngSeriesRow + 1) = 1

            .Cell(flexcpChecked, lngSeriesRow + 1, 0) = blnChecked
            .Cell(flexcpPicture, lngSeriesRow + 1, 2) = imgImage

            .Cell(flexcpText, lngSeriesRow + 1, 1) = Nvl(rsData!ͼ��UID)
            .Cell(flexcpText, lngSeriesRow + 1, 2) = "ͼ��" & Nvl(rsData!ͼ���)
            .Cell(flexcpText, lngSeriesRow + 1, 3) = "ͼ���:" & Nvl(rsData!ͼ���) & "  �ɼ�ʱ��:" & Nvl(rsData!�ɼ�ʱ��)

            .Cell(flexcpFontSize, lngSeriesRow + 1, 3) = 9
            .Cell(flexcpForeColor, lngSeriesRow + 1, 3) = &HC0C0FF

            .IsSubtotal(lngSeriesRow + 1) = True
            .RowOutlineLevel(lngSeriesRow + 1) = 3

            Call rsData.MoveNext
        Wend
    End With
    
End Sub


Private Sub vsfTree_SelChange()
On Error GoTo ErrHandle
    Dim rsData As ADODB.Recordset
    Dim strSearchId As String
    Dim lngSearchType As Long
    
    ucPage.RecordCount = 0
    
    If vsfTree.Row < 0 Then Exit Sub
    If vsfTree.RowSel < 0 Then Exit Sub
    
    
    If mlngReleationType = 2 Then
        '����ڵ�Ϊ��飬����Ҫ�ж��Ƿ����Ӳ�ڵ㣬���û�������
        If vsfTree.RowOutlineLevel(vsfTree.Row) = 1 Then
            Call LoadSeriesNode(vsfTree.Row)
        End If
        
'        '����ڵ�Ϊ���У�����Ҫ�ж��Ƿ����Ӳ�ڵ㣬���û�������
'        If vsfTree.RowOutlineLevel(vsfTree.Row) = 2 Then
'            Call LoadImageNode(vsfTree.Row)
'        End If
    End If
    
    'û������ͼ��Ԥ��ʱ���ⲻ����ͼ��
    If chkViewImage.value = 0 Then Exit Sub
    
    strSearchId = vsfTree.TextMatrix(vsfTree.Row, 1)
    lngSearchType = vsfTree.RowOutlineLevel(vsfTree.Row)
    
    
    Call InitPageControl(lngSearchType, strSearchId)
    
    Set rsData = GetImageViewData(lngSearchType, strSearchId, 1, ucPage.PageRecord)
    Call LoadViewImageToFace(rsData)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsfTree_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
