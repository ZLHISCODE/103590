VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.dockingpane.9600.ocx"
Begin VB.Form frmServiceHistory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picApp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   3735
      ScaleHeight     =   6180
      ScaleWidth      =   8340
      TabIndex        =   2
      Top             =   0
      Width           =   8340
      Begin VB.PictureBox picDetail 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         Enabled         =   0   'False
         FillColor       =   &H8000000C&
         ForeColor       =   &H8000000C&
         Height          =   5205
         Left            =   420
         ScaleHeight     =   5175
         ScaleWidth      =   6090
         TabIndex        =   3
         Top             =   555
         Width           =   6120
         Begin VB.TextBox txtSum 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4095
            TabIndex        =   33
            Top             =   4305
            Width           =   1920
         End
         Begin VB.TextBox txtTotal 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            TabIndex        =   30
            Top             =   4305
            Width           =   1920
         End
         Begin VB.TextBox txtRemark 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            TabIndex        =   28
            Top             =   4680
            Width           =   5175
         End
         Begin VB.TextBox txtAppTime 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4095
            TabIndex        =   26
            Top             =   3930
            Width           =   1920
         End
         Begin VB.ComboBox cboAppStyle 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   3930
            Width           =   1920
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfItem 
            Height          =   1140
            Left            =   75
            TabIndex        =   23
            Top             =   2715
            Width           =   5940
            _cx             =   10477
            _cy             =   2011
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmServiceHistory.frx":0000
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
         Begin VB.ComboBox cboPayStyle 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   2310
            Width           =   1920
         End
         Begin VB.ComboBox cboFeeType 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4095
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1920
            Width           =   1920
         End
         Begin VB.ComboBox cboMedType 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1920
            Width           =   1920
         End
         Begin VB.ComboBox cboDoc 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4095
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1395
            Width           =   1920
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   0
            TabIndex        =   15
            Top             =   1800
            Width           =   6800
         End
         Begin VB.TextBox txtDept 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            TabIndex        =   13
            Top             =   1395
            Width           =   1920
         End
         Begin VB.TextBox txtSN 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4095
            TabIndex        =   11
            Top             =   1005
            Width           =   1920
         End
         Begin VB.TextBox txtNO 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   840
            TabIndex        =   9
            Top             =   1005
            Width           =   1920
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   0
            TabIndex        =   7
            Top             =   885
            Width           =   6800
         End
         Begin VB.ComboBox cboNO 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   4515
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   525
            Width           =   1530
         End
         Begin VB.Label lblCancel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   5640
            TabIndex        =   35
            Top             =   45
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "ʵ�պϼ�"
            Height          =   180
            Left            =   3300
            TabIndex        =   34
            Top             =   4365
            Width           =   720
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "ԤԼ�Һŵ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2235
            TabIndex        =   32
            Top             =   105
            Width           =   1500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ�պϼ�"
            Height          =   180
            Left            =   60
            TabIndex        =   31
            Top             =   4365
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "��ע"
            Height          =   180
            Left            =   420
            TabIndex        =   29
            Top             =   4740
            Width           =   360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ԤԼʱ��"
            Height          =   180
            Left            =   3300
            TabIndex        =   27
            Top             =   3990
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "ԤԼ��ʽ"
            Height          =   180
            Left            =   60
            TabIndex        =   25
            Top             =   3990
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "���ʽ"
            Height          =   180
            Left            =   75
            TabIndex        =   22
            Top             =   2370
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "�ѱ�"
            Height          =   180
            Left            =   3660
            TabIndex        =   20
            Top             =   1980
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "ҽ�����"
            Height          =   180
            Left            =   75
            TabIndex        =   18
            Top             =   1980
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��"
            Height          =   180
            Left            =   3660
            TabIndex        =   14
            Top             =   1455
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   435
            TabIndex        =   12
            Top             =   1455
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   3660
            TabIndex        =   10
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�ű�"
            Height          =   180
            Left            =   435
            TabIndex        =   8
            Top             =   1065
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "���ݺ�"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   3930
            TabIndex        =   5
            Top             =   585
            Width           =   540
         End
      End
      Begin VB.Label lblNote 
         Caption         =   "�ò��˵�ǰû��ԤԼ�Һŵ�..."
         Height          =   270
         Left            =   100
         TabIndex        =   4
         Top             =   100
         Visible         =   0   'False
         Width           =   3420
      End
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   720
      ScaleHeight     =   4185
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   1500
      Width           =   3645
      Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
         Height          =   1755
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   2970
         _cx             =   5239
         _cy             =   3096
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
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmServiceHistory.frx":006A
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   195
      Top             =   1335
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmServiceHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitPanel()
    Dim objPane As Pane
    
    Err = 0: On Error GoTo errHandle
    Set objPane = dkpMain.CreatePane(1, 145, 80, DockLeftOf, Nothing)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Title = "������Ϣ"
    objPane.Handle = picInfo.hWnd
    objPane.MaxTrackSize.Width = 300
    objPane.MinTrackSize.Width = 100
    
    Set objPane = dkpMain.CreatePane(2, 145, 90, DockRightOf, objPane)
    objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption
    objPane.Title = "ԤԼ��Ϣ"
    objPane.Handle = picApp.hWnd
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .PaintManager.HighlighActiveCaption = False
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub LoadData(ByVal lng��ϢID As Long)
    Dim i As Integer, lngRow As Long, rsItem As ADODB.Recordset
    Dim dblTotal As Double, dblSum As Double
    Dim strSQL As String, rsInfo As ADODB.Recordset, rsTemp As ADODB.Recordset
    strSQL = "Select b.����, b.�Ա�, b.����, b.��������, b.���֤��, b.�����, b.��ͥ�绰 As ��ϵ�绰, b.��ͥ��ַ As ��ס��ַ, b.���ڵ�ַ, c.���� As ����, d.���� As ��Ŀ," & vbNewLine & _
            "       a.ҽ������ As ҽ��, a.��ʼʱ��, a.��ֹʱ��, a.֪ͨԭ�� As ԤԼԭ��, a.�Ǽ���, a.�Ǽ�ʱ��, a.�Һ�id" & vbNewLine & _
            "From ���˷�����Ϣ��¼ A, ������Ϣ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.Id = [1] And a.����id = b.����id And a.����id = c.Id And a.��Ŀid = d.Id"
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ϢID)
    With vsfInfo
        .Clear
        .Rows = 18: .Cols = 2
        lngRow = 0
        .TextMatrix(lngRow, 0) = "���˻�����Ϣ"
        .TextMatrix(lngRow, 1) = "���˻�����Ϣ"
        .MergeRow(lngRow) = True
        .Cell(flexcpBackColor, lngRow, 0, lngRow, 1) = &HFFC0C0
        .Cell(flexcpFontBold, lngRow, 0, lngRow, 1) = True
        .Cell(flexcpAlignment, lngRow, 0, lngRow, 1) = 1
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "����"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!����)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "�Ա�"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!�Ա�)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "����"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!����)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "��������"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!��������)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "���֤��"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!���֤��)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "�����"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!�����)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "��ϵ�绰"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!��ϵ�绰)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "��ס��ַ"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!��ס��ַ)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "���ڵ�ַ"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!���ڵ�ַ)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "ԤԼ�Ǽ���Ϣ"
        .TextMatrix(lngRow, 1) = "ԤԼ�Ǽ���Ϣ"
        .MergeRow(lngRow) = True
        .Cell(flexcpBackColor, lngRow, 0, lngRow, 1) = &HFFC0C0
        .Cell(flexcpFontBold, lngRow, 0, lngRow, 1) = True
        .Cell(flexcpAlignment, lngRow, 0, lngRow, 1) = 1
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "����"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!����)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "��Ŀ"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!��Ŀ)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "ҽ��"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!ҽ��)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "ԤԼʱ��"
        .TextMatrix(lngRow, 1) = Format(Nvl(rsInfo!��ʼʱ��), "yyyy-mm-dd") & "��" & Format(Nvl(rsInfo!��ֹʱ��), "yyyy-mm-dd")
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "ԤԼԭ��"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!ԤԼԭ��)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "�Ǽ���"
        .TextMatrix(lngRow, 1) = Nvl(rsInfo!�Ǽ���)
        lngRow = lngRow + 1
        .TextMatrix(lngRow, 0) = "�Ǽ�ʱ��"
        .TextMatrix(lngRow, 1) = Format(Nvl(rsInfo!�Ǽ�ʱ��), "yyyy-mm-dd hh:mm:ss")
        lngRow = lngRow + 1
        .AutoSize 0
        For i = 0 To .Rows - 1
            .RowHeight(i) = 322
        Next i
    End With
    If Val(Nvl(rsInfo!�Һ�ID)) = 0 Then
        lblNote.Visible = True
        picDetail.Visible = False
    Else
        lblNote.Visible = False
        picDetail.Visible = True
        strSQL = "Select a.��¼״̬, a.No, a.�ű�, a.����, b.���� As ����, a.ִ���� As ҽ��, c.ҽ�����, d.�ѱ�, a.ҽ�Ƹ��ʽ, a.ԤԼ��ʽ, a.ԤԼʱ��, a.ժҪ" & vbNewLine & _
                "From ���˹Һż�¼ A, ���ű� B, ����ǼǼ�¼ C, ������ü�¼ D" & vbNewLine & _
                "Where a.Id = [1] And a.ִ�в���id = b.Id And a.����id = c.����id(+) And a.���� = c.����(+) And Nvl(c.��ҳid(+), 0) = 0 And" & vbNewLine & _
                "      a.�Ǽ�ʱ�� = c.����ʱ��(+) And a.No = d.No And d.��¼���� = 4 And d.��¼״̬ In (0, 1, 3) And d.��� = 1"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsInfo!�Һ�ID)))
        If rsTemp.EOF Then
            MsgBox "��ȡԤԼ��¼���ִ���,�޷���ʾԤԼ�Һŵ�!", vbInformation, gstrSysName
            lblNote.Visible = True
            picDetail.Visible = False
            Exit Sub
        End If
        strSQL = "Select b.����, a.Ӧ�ս��, a.ʵ�ս��" & vbNewLine & _
                "From ������ü�¼ A, �շ���ĿĿ¼ B" & vbNewLine & _
                "Where a.�շ�ϸĿid = b.Id And a.No = [1] And a.��¼���� = 4 And a.��¼״̬ In (0, 1, 3)"
        Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(rsTemp!NO))
        With vsfItem
            .Clear 1
            .Rows = 2
            Do While Not rsItem.EOF
                .TextMatrix(.Rows - 1, 0) = Nvl(rsItem!����)
                .TextMatrix(.Rows - 1, 1) = Format(Val(Nvl(rsItem!Ӧ�ս��)), "0.00")
                dblTotal = dblTotal + Val(Nvl(rsItem!Ӧ�ս��))
                .TextMatrix(.Rows - 1, 2) = Format(Val(Nvl(rsItem!ʵ�ս��)), "0.00")
                dblSum = dblSum + Val(Nvl(rsItem!ʵ�ս��))
                .Rows = .Rows + 1
                rsItem.MoveNext
            Loop
            If .Rows > 2 Then .Rows = .Rows - 1
        End With
        
        lblCancel.Visible = Val(Nvl(rsTemp!��¼״̬)) = 3
        cboNO.AddItem Nvl(rsTemp!NO)
        cboNO.ListIndex = cboNO.NewIndex
        txtNO.Text = Nvl(rsTemp!�ű�)
        txtSN.Text = Nvl(rsTemp!����)
        txtDept.Text = Nvl(rsTemp!����)
        cboDoc.AddItem Nvl(rsTemp!ҽ��)
        cboDoc.ListIndex = cboDoc.NewIndex
        cboMedType.AddItem Nvl(rsTemp!ҽ�����)
        cboMedType.ListIndex = cboMedType.NewIndex
        cboFeeType.AddItem Nvl(rsTemp!�ѱ�)
        cboFeeType.ListIndex = cboFeeType.NewIndex
        cboPayStyle.AddItem Nvl(rsTemp!ҽ�Ƹ��ʽ)
        cboPayStyle.ListIndex = cboPayStyle.NewIndex
        cboAppStyle.AddItem Nvl(rsTemp!ԤԼ��ʽ)
        cboAppStyle.ListIndex = cboAppStyle.NewIndex
        txtAppTime.Text = Format(Nvl(rsTemp!ԤԼʱ��), "yyyy-mm-dd hh:mm:ss")
        txtTotal.Text = Format(dblTotal, "0.00")
        txtSum.Text = Format(dblSum, "0.00")
        txtRemark.Text = Nvl(rsTemp!ժҪ)
        
    End If
End Sub


Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call InitPanel
    lblTitle = gstrUnitName & "ԤԼ�Һŵ�"
    lblTitle.Left = (picDetail.ScaleWidth - lblTitle.Width) / 2
End Sub


Private Sub picApp_Resize()
    On Error Resume Next
    With picDetail
        .Left = (picApp.ScaleWidth - .Width) / 2
        .Top = 800
        .Height = picApp.ScaleHeight - 1600
    End With
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    Label11.Top = picDetail.ScaleHeight - 435
    Label12.Top = picDetail.ScaleHeight - 810
    Label13.Top = picDetail.ScaleHeight - 810
    Label9.Top = picDetail.ScaleHeight - 1185
    Label10.Top = picDetail.ScaleHeight - 1185
    txtRemark.Top = picDetail.ScaleHeight - 495
    txtTotal.Top = picDetail.ScaleHeight - 870
    txtSum.Top = picDetail.ScaleHeight - 870
    cboAppStyle.Top = picDetail.ScaleHeight - 1245
    txtAppTime.Top = picDetail.ScaleHeight - 1245
    vsfItem.Height = txtAppTime.Top - vsfItem.Top - 120
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    With vsfInfo
        .Width = picInfo.ScaleWidth - 30
        .Height = picInfo.ScaleHeight - 30
    End With
End Sub

Private Sub vsfInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

