VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanPlanDatSet 
   AutoRedraw      =   -1  'True
   Caption         =   "�ҺŰ��żƻ�ʱ������"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   Icon            =   "frmRegistPlanPlanDatSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10440
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7440
      TabIndex        =   31
      Top             =   6795
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8760
      TabIndex        =   30
      Top             =   6795
      Width           =   1100
   End
   Begin VB.Frame fraӦ���� 
      Caption         =   "Ӧ����(&B)"
      Height          =   615
      Left            =   240
      TabIndex        =   26
      Top             =   6640
      Width           =   7095
      Begin VB.OptionButton opt���� 
         Caption         =   "Ӧ��������"
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "Ӧ���뱾����"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt��ҽ�� 
         Caption         =   "Ӧ���ڱ�ҽ��"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "������Ϣ"
      Height          =   1380
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   10095
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   20
         Text            =   "cbo����"
         Top             =   307
         Width           =   1155
      End
      Begin VB.TextBox txt��Լ 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   14
         Top             =   307
         Width           =   1215
      End
      Begin VB.TextBox txt�޺� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4965
         MaxLength       =   5
         TabIndex        =   13
         Top             =   307
         Width           =   975
      End
      Begin VB.CheckBox chk��ſ��� 
         Caption         =   "��ſ���"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�Һ�ʱ���뽨����"
         Enabled         =   0   'False
         Height          =   195
         Left            =   8040
         TabIndex        =   5
         Top             =   360
         Width           =   1845
      End
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         TabIndex        =   2
         Text            =   "cbo����"
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   4
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   3
         Text            =   "cboItem"
         Top             =   705
         Width           =   2235
      End
      Begin VB.TextBox txt�ű� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   5
         TabIndex        =   0
         Top             =   307
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "��Լ"
         Height          =   180
         Left            =   6240
         TabIndex        =   16
         Top             =   367
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�޺�"
         Height          =   180
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3000
         TabIndex        =   12
         Top             =   367
         Width           =   360
      End
      Begin VB.Label lblҽ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ժ��ҽ��"
         Height          =   180
         Left            =   5940
         TabIndex        =   10
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ"
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   367
         Width           =   390
      End
   End
   Begin VB.Frame fraDate 
      Caption         =   "ʱ������"
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   10215
      Begin VB.PictureBox picTime 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   9945
         TabIndex        =   17
         Top             =   240
         Width           =   9945
         Begin VB.CommandButton cmdOtherCalc 
            Caption         =   "���������ƻ�(&R)"
            Height          =   360
            Left            =   3765
            TabIndex        =   32
            Top             =   30
            Width           =   1860
         End
         Begin VB.CommandButton cmd����ʱ�� 
            Caption         =   "��������(&F)"
            Height          =   350
            Left            =   2520
            TabIndex        =   25
            ToolTipText     =   "������¼���ʱ��"
            Top             =   35
            Width           =   1150
         End
         Begin VB.TextBox txtTimeOut 
            Height          =   300
            Left            =   1560
            TabIndex        =   23
            Text            =   "10"
            Top             =   60
            Width           =   500
         End
         Begin MSComCtl2.UpDown udTime 
            Height          =   345
            Left            =   2160
            TabIndex        =   22
            Top             =   38
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComctlLib.TabStrip tbWeekTime 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTime 
            Height          =   3825
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   9765
            _cx             =   17224
            _cy             =   6747
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12632256
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanPlanDatSet.frx":000C
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
            Begin VB.CommandButton cmdɾ�� 
               Caption         =   "ɾ"
               Height          =   255
               Left            =   7275
               TabIndex        =   33
               Top             =   2145
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton cmdԤԼ 
               Caption         =   "Ԥ"
               Height          =   255
               Left            =   7320
               TabIndex        =   21
               Top             =   1560
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ʱ����(��)"
            Height          =   180
            Left            =   360
            TabIndex        =   24
            Top             =   120
            Width           =   1080
         End
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "Ժ��ҽ��"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "����Ԯҽ��"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanPlanDatSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit 'Ҫ���������
'
 
Private mViewMode         As ViewMode    'ҳ����ʾģʽ
Private mlng�ƻ�Id        As Long        '�ƻ�ID
Private mlngPre�ƻ�ID     As Long
Private mrsTime          As ADODB.Recordset
Private mrs�޺�          As ADODB.Recordset
Private mrs�ϰ�ʱ���     As ADODB.Recordset
Private mrs���żƻ�          As ADODB.Recordset
Private mblnCellChange   As Boolean
Private mstrKey         As String
Private mblnChange      As Boolean
Private mblnReload      As Boolean '�ڹҺŰ��żƻ�����ҳ����� ShowMe�Ժ� �Ƿ���Ҫˢ��
Private mrs�ϴμƻ�ʱ�� As Recordset '�����52275
Private mbln׷�Ӻ� As Boolean '�����52275


'�����ϰ�ʱ��
Private Type t_�ϰ�ʱ��
  dat_�����ϰ� As Date
  dat_�����°� As Date
  dat_�����ϰ� As Date
  dat_�����°� As Date
End Type
Private t_ʱ�� As t_�ϰ�ʱ��
Private Const strMaskKey As String = "09:00-09:00"
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther '�����:51429
Attribute mfrmOtherCalc.VB_VarHelpID = -1

Private Sub chk��ſ���_Click()
    cmdOtherCalc.Visible = chk��ſ���.Value = 1
End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
    cmdOK.Enabled = False
    zlCommFun.ShowFlash "���ڱ���Һżƻ�ʱ������,���Ժ򡭡�"
    If SaveDate() = True Then
        '************************
        '�������ɹ���Ҫ���¶�
        '�Һżƻ�ʱ�ν�����ȡ
        '************************
        Call InitData
        mblnChange = False
        mblnReload = True
       ' If tbWeekTime.Tabs.Count > 0 Then tbWeekTime.Tabs(1).Selected = True
    End If
    zlCommFun.StopFlash
    cmdOK.Enabled = True
End Sub

Private Sub cmdɾ��_Click()
    Call DeleteSelectPain
End Sub

Private Sub mfrmOtherCalc_zlRefreshCon(ByVal VarTimes As Variant)
    Dim strTemp  As String, varData As Variant, varTemp As Variant
    Dim i As Long, int���� As Integer, dtStart As Date, dtEnd As Date
    Dim lngRow As Long, lng��� As Long, dtTemp As Date, j As Long
    Dim lng�޺��� As Long, lng��Լ�� As Long, str���� As String
    Dim lng�ѹ������� As Long '�����:51427
    Dim lngCol As Long '�����:54127
    Dim lng�ƻ�ID As Long '�����:54127
    Dim K As Long '�����:54127
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng�ƻ�ID = Val(txt�ű�.Tag) '�����:51427
    
    If Get�޺���(str����, lng�޺���, lng��Լ��) = False Then Exit Sub
    
    'VarTiems
    '       "ʱ����"
    '       "�ֶμ��":ʱ��(��:8:00��9:00),2;ʱ��2,���;....
    If VarTimes("ʱ����") <> "" Then
        txtTimeOut.Text = Val(VarTimes("ʱ����"))
        Call cmd����ʱ��_Click
        Exit Sub
    End If
    strTemp = VarTimes("�ֶμ��")
    If strTemp = "" Then Exit Sub
    
    '�����:52275
    If mbln׷�Ӻ� = False Then
        '�����:51427
        lng�ѹ������� = ExistsBooking(lng�ƻ�ID, str����)
        If lng�ѹ������� <> -1 Then
             If MsgBox("�ð��������б��ҳ�ȥ�ĺ�,ֻ���޸ĺ�ɫ������ʾ��ʱ��" & vbCrLf & "��ȷ��Ҫ�����޸���?", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
             End If
        End If
    Else
        lng�ѹ������� = mrs�ϴμƻ�ʱ��.RecordCount
    End If
    
    varData = Split(strTemp, ";")
    lngRow = -2: lng��� = 1: lngCol = 1
    '�����:51427
    For i = 0 To vsTime.Rows - 1
        For j = 0 To vsTime.Cols - 1
            If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                If CLng(vsTime.TextMatrix(i, j)) = lng�ѹ������� Then
                    lngRow = i: lngCol = j
                End If
            End If
        Next
    Next
    
    '��ʼ��vsTime�ؼ�
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = IIf(lngRow = -2, 2, lngRow + 2): lngRow = IIf(lngRow = 0 And lngCol = 1, -2, lngRow): i = 0: .FixedCols = 1
        .FixedRows = 0
    If lngRow = -2 Then
            .Rows = 0
            .Rows = 2
        End If
    lng��� = IIf(lng�ѹ������� = -1, 1, lng�ѹ������� + 1)
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int���� = Val(varTemp(1))
        varTemp = Split(varTemp(0), "��")
        dtStart = CDate(varTemp(0))
        dtEnd = CDate(varTemp(1))
        'ͬһʱ�����û�йҳ��ĺ���
        If dtStart = IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            j = IIf(lngCol = 1 And lngRow = -2, 0, lngCol) + 1
            '���û�ҳ���ѡ��
            For K = j To .Cols - 1
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), K) = ""
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, K) = ""
            Next
            .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = Format(dtStart, "HH:00")
            .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0) = Format(dtStart, "HH:00")
            If lngCol = 1 Then
                dtStart = .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0)
            Else
                dtStart = Split(.TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, lngCol), "-")(1)
            End If
            '�����:52275
            If mbln׷�Ӻ� = True Then
                dtStart = Split(.TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, lngCol), "-")(1)
            End If
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int���� * 1 / 24 / 60, "HH:MM")
                '�����:52275
                 If IIf(mbln׷�Ӻ� = False, dtTemp > dtEnd, 1 = 0) Or lng��� > lng�޺��� Then Exit Do
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), j) = lng���
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng��� = lng��� + 1
                j = j + 1
            Loop
        dtStart = "00:00:00"
        End If
        '��ͬʱ���û�б��ҳ��ĺ���
        If dtStart > IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            If IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) <> Format(dtStart, "HH:00") Then
                 If lng��� > 1 Then
                     lngRow = IIf(lngRow = -2, 0, lngRow)
                 End If
                 lngRow = lngRow + 2
                .Rows = .Rows + 2
                .TextMatrix(lngRow, 0) = Format(dtStart, "HH:00")
                .TextMatrix(lngRow + 1, 0) = Format(dtStart, "HH:00")
            End If
            j = 1
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int���� * 1 / 24 / 60, "HH:MM")
                '�����:52275
                 If IIf(mbln׷�Ӻ� = False, dtTemp > dtEnd, 1 = 0) Or lng��� > lng�޺��� Then Exit Do
                .TextMatrix(lngRow, j) = lng���
                .TextMatrix(lngRow + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng��� = lng��� + 1
                j = j + 1
            Loop
        End If
    Next
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignCenterCenter
        .ColWidth(i) = 1200
    Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
    If .Rows > 0 Then
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
    End If
    .Redraw = flexRDBuffered
    End With
    Call setVsFlexBgColor(True)
End Sub

Private Sub cmd����ʱ��_Click()
'�ԹҺżƻ�ʱ�ν�������
    Dim str����         As String
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrsTime.Filter = "����='" & str���� & "'"
    If mrsTime.RecordCount > 0 Then
      '****************************************************************
      '�����йҺżƻ�ʱ�ε������
      '��ʾ����Ա �Ƿ���Ҫ���¼���ʱ��
      '****************************************************************
        If MsgBox("�˰��żƻ���" & str���� & "�Ѿ�����ʱ�� " & vbCrLf & "�Ƿ����¼���ʱ��?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            mrsTime.Filter = 0
            Exit Sub
        End If
    End If
    Select Case chk��ſ���.Value = 1
    Case True:
        Setר�Һ�ʱ��
    Case False:
        Set��ͨ��ʱ��
    End Select
    setVsFlexBgColor (chk��ſ���.Value = 1)
    mblnChange = True
End Sub
Private Sub Initʱ���()
  '--------------------------------
  '����:��ȡ���°�ʱ���
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("08:00:00")
    End If
   
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 18:00:00")
    End If
    With t_ʱ��
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
    End With
    strSQL = _
    "       Select ʱ���,��ǩ,�ϰ�, �°� " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select ʱ���,To_Date('1900-01-01 ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ʼʱ��," & vbNewLine & _
    "                               To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ֹʱ��," & _
    "                               Sign(��ʼʱ�� - ��ֹʱ��) As ����, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��, " & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��," & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��"
    strSQL = strSQL & vbNewLine & _
    "                       From ʱ��� )" & vbNewLine & _
    "           Select ʱ���, '��' As ��ǩ, 0 As ��־, ��ʼʱ�� As �ϰ�, ��ֹʱ�� As �°�, ��ʼʱ��, ��ֹʱ��," & _
    "                  �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ��" & vbNewLine & _
    "            From Tb  Where (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) And " & _
    "                      (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & vbNewLine & _
    "                        Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, " & _
    "                        �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "           From Tb a Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & _
    "                   Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "         From Tb a   Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By ʱ���,�ϰ�"
     Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
    

Private Sub Set��ͨ��ʱ��()
    Dim strSQL      As String
    Dim str����     As String
    Dim strʱ��     As String
    Dim lng�޺�     As Long
    Dim lng��Լ     As Long
    Dim lng���     As Long
    Dim dblDatCount As Long '��ʱ����
    Dim datʱ��     As Date 'ÿ��ʱ��ε�
    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ��żƻ�����" & vbCrLf & "�뵽�ҺŰ��żƻ�������!", vbOKOnly, Me.Caption
        Exit Sub '����ҺŰ��żƻ���û�����ô������Ϣ �Ͳ���������
    End If
    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt�޺�.Text = lng�޺�
    Me.txt��Լ.Text = lng��Լ
    If lng��Լ = 0 Then lng��Լ = lng�޺� '�����ԤԼû����������Ϊ�����Լ�����޺�����ͬ
    strʱ�� = Nvl(mrs���żƻ�(str����).Value)
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    lng��� = Val(txtTimeOut.Text)
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
    End With
    '*************************************
    '��ͨ��
    '*************************************
    With vsTime
        .Cols = 8: .FixedCols = 0
        .Rows = 1: .FixedRows = 1
        For i = 0 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "ʱ���"
        Next
        For i = 1 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "ԤԼ����"
        Next
        lngRow = 1: lngCol = -1
        j = 1: lngStart = 1
      Do While Not mrs�ϰ�ʱ���.EOF
            If blnExit Then Exit Do
            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
            For i = j To lng�޺�
                If lngStart > lng�޺� Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                
                lngCol = lngCol + 1
                If lngCol * 2 > .Cols - 2 Then lngRow = lngRow + 1: lngCol = 0
                strData = IIf(lng��Լ >= i, 1, 0)
                strTime = Format(datʱ��, "HH:mm") & "-" & _
                      IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                      Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
               
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, lngCol * 2) = strTime
                .TextMatrix(lngRow, lngCol * 2 + 1) = strData
                lngStart = lngStart + 1
                datʱ�� = DateAdd("n", lng���, datʱ��)
            Next
            mrs�ϰ�ʱ���.MoveNext
        Loop
 
         For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub
Private Sub Setר�Һ�ʱ��()
    Dim strSQL      As String
    Dim str����     As String
    Dim strʱ��     As String
    Dim lng�޺�     As Long
    Dim lng��Լ     As Long
    Dim lng���     As Long
    Dim dblDatCount As Long '��ʱ����
    Dim datʱ��     As Date 'ÿ��ʱ��ε�
    Dim strʱ��     As String
    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs�ϰ�ʱ��� Is Nothing Then
        strSQL = _
        "     Select ʱ���, To_Char(��ʼʱ��, 'HH24:MI:SS') As ��ʼʱ��, To_Char(��ֹʱ��, 'HH24:MI:SS') As ��ֹʱ�� From ʱ���    "
        Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If mrs�ϰ�ʱ���.EOF Then Set mrs�ϰ�ʱ��� = Nothing: Exit Sub
    End If
    If mrs�޺� Is Nothing Then
        strSQL = _
        "Select �ƻ�id, ������Ŀ as ���� , �޺���, ��Լ�� From �ҺŰ��żƻ����� Where �ƻ�id = [1]"
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt�ű�.Tag))
        If mrsTime.RecordCount = 0 Then
        MsgBox "��ǰ�ű�û�ж�Ӧ�ĹҺŰ��żƻ�����" & vbCrLf & "�뵽�ҺŰ��żƻ�������!", vbOKOnly, Me.Caption
        Set mrs�޺� = Nothing
        Exit Sub '����ҺŰ��żƻ���û�����ô������Ϣ �Ͳ���������
    End If
    End If
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ��żƻ�����" & vbCrLf & "�뵽�ҺŰ��żƻ�������!", vbOKOnly, Me.Caption
        Exit Sub '����ҺŰ��żƻ���û�����ô������Ϣ �Ͳ���������
    End If
    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt�޺�.Text = lng�޺�
    Me.txt��Լ.Text = lng��Լ
    lng��Լ = lng�޺�
    strʱ�� = Nvl(mrs���żƻ�(str����).Value)
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
'*************************************************************
'ʱ�������� ���õļ��
'*************************************************************
      lng��� = Val(Me.txtTimeOut.Text)
   
      With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
      End With
    '*************************************
    'ר�Һ�
    '���������
    '���� ʱ��α��е� ���°�ʱ�����ж�
    '���� ȫ���������  ��Ϊ���������
    '*************************************
    
    With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         j = 1
         lngStart = 1
         Do While Not mrs�ϰ�ʱ���.EOF
            If blnExit Then Exit Do
             
            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
             For i = j To lng��Լ
                If lngStart > lng��Լ Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                 End If
                lngCol = lngCol + 1
                If strʱ�� <> Format(datʱ��, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
                If lngCol = 1 Then
                     If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                     strʱ�� = Format(datʱ��, "HH") & ":00"
                     vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
                     vsTime.TextMatrix(lngRow, 0) = strʱ��
                
                End If
                strData = lngStart
                lngStart = lngStart + 1
                strTime = Format(datʱ��, "HH:mm") & "-" & _
                           IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                           Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
    
                If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
                vsTime.TextMatrix(lngRow - 1, lngCol) = strData
                vsTime.TextMatrix(lngRow, lngCol) = strTime
                '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
                
                datʱ�� = DateAdd("n", lng���, datʱ��)
             Next
             mrs�ϰ�ʱ���.MoveNext
         Loop
 
         For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .ColWidth(0) = 1200
         .FixedAlignment(0) = flexAlignRightTop
         .ColAlignment(0) = flexAlignRightTop
         If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
         End If
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdԤԼ_Click()
    '��ʱ����ܷ�ԤԼ��������
    On Error GoTo ErrHandl:
    If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Then Exit Sub
    If mViewMode = ViewMode.ViewItem Or vsTime.TextMatrix(vsTime.MouseRow, vsTime.MouseCol) = "" Then Exit Sub
    With vsTime
        If IsNumeric(.Cell(flexcpText, .Row, .Col)) = False And chk��ſ���.Value = 1 Then
            .Row = .Row - 1
        ElseIf IsNumeric(.Cell(flexcpText, .Row, .Col)) = True And chk��ſ���.Value <> 1 Then
            .Col = .Col - 1
        End If
        If .CellForeColor = vbBlue Then
            If chk��ſ���.Value = 1 Then
                .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = &H80000008
                .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = False
            Else
                .Cell(flexcpForeColor, .Row, .Col, .Row, .Col + 1) = &H80000008
                '.Cell(flexcpFontBold, .Row, .Col, .Row, .Col + 1) = False
            End If
        Else
            If chk��ſ���.Value = 1 Then
                .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = vbBlue
                .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = True
            Else
                .Cell(flexcpForeColor, .Row, .Col, .Row, .Col + 1) = vbBlue
                '.Cell(flexcpFontBold, .Row, .Col, .Row, .Col + 1) = True
            End If
        End If
    End With
    mblnChange = True
ErrHandl:
    mblnChange = True
End Sub

Private Sub Form_Activate()
    Me.Icon = frmRegistPlan.Icon
End Sub

Private Sub Form_Load()
    Initʱ���
     '�����:52275
    Set mrs�ϴμƻ�ʱ�� = Get��һ�μƻ�ʱ��
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '********************************************
  '�������� �������С��Ⱥ���С�߶�
  '********************************************
  If Me.Width < 701 * Screen.TwipsPerPixelX Then Me.Width = 701 * Screen.TwipsPerPixelX
  If Me.Height < 511 * Screen.TwipsPerPixelY Then Me.Height = 511 * Screen.TwipsPerPixelY
  '********************************************
  '�ҺŰ��Ż�����Ϣ λ�ò��ƶ��ƶ�
  '���ƶ� ʱ������
  '********************************************
  With fraDate
     .Width = Me.ScaleWidth - 2 * .Left
     .Height = Me.ScaleHeight - Me.fraInfo.Top - Me.fraInfo.Height - 65 * Screen.TwipsPerPixelY
  End With
  
  With picTime
     .Width = fraDate.Width - 2 * .Left
     .Height = fraDate.Height - .Top * 2
  End With
  With Me.tbWeekTime
    .Width = picTime.ScaleWidth - 2 * .Left
  End With
  With Me.vsTime
    .Width = picTime.ScaleWidth - 2 * .Left
    .Height = picTime.ScaleHeight - .Top - cmd����ʱ��.Top
  End With
  '-------------------------------------------
  'Ӧ���� λ�õĵ���
  '-------------------------------------------
  With Me.fraӦ����
       .Left = .Left
       .Top = Me.fraDate.Top + Me.fraDate.Height + 5 * Screen.TwipsPerPixelY
   
  End With
  
  '********************************************
  'ȷ����ť��ȡ����ť���ƶ�
  '********************************************
  
  With Me.cmdCancel
       .Left = Me.ScaleWidth - 40 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
  With Me.cmdOK
       .Left = cmdCancel.Left - 20 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
     mlngPre�ƻ�ID = -1
     mblnChange = False
     Set mrsTime = Nothing
     Set mrs�޺� = Nothing
     Set mrs�ϰ�ʱ��� = Nothing
     Set mrs���żƻ� = Nothing
End Sub

 

Private Sub tbWeekTime_Click()
    Dim i       As Integer
    Dim j As Long '�����:51427
    Dim lng�ѹ������� As Long '�����:51427
    Dim rs��ǰ�ƻ�ʱ�� As Recordset '�����:52221
    Dim strMsg As String
    Dim vMsgResult As VbMsgBoxResult
    Dim rs�Ӻ� As Recordset
    Dim str���ʱ�䷶Χ As String '�����:5555
    Dim bln����ʱ�� As Boolean '�����:55555
    Dim lngĬ�ϼ��ʱ�� As Long '�����:55555
     '�����:52275
    mbln׷�Ӻ� = False
    If mblnChange Then
        mblnChange = False
        If MsgBox("��ǰ�ҺŰ��żƻ���" & mstrKey & "��ʱ���Ѹı�!�Ƿ񱣴�?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption) = vbYes Then
            cmdOk_Click
         For i = 1 To tbWeekTime.Tabs.Count
            If tbWeekTime.Tabs(i).Key = "K" & mstrKey Then
                tbWeekTime.Tabs(i).Selected = True
                Exit For
            End If
         Next
        End If
    End If

    mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2)
     '�����:52275
    Set rs��ǰ�ƻ�ʱ�� = Get��ǰ�ƻ�ʱ��
    If Not mrs�ϴμƻ�ʱ�� Is Nothing Then
        rs��ǰ�ƻ�ʱ��.Filter = "���� = '" & mstrKey & "'"
        mrs�ϴμƻ�ʱ��.Filter = " ���� ='" & mstrKey & "'"
        mrs�޺�.Filter = ""
        
        If rs��ǰ�ƻ�ʱ��.RecordCount <= 0 And mrs�ϴμƻ�ʱ��.RecordCount > 0 Then
         If mrs�ϴμƻ�ʱ��!ʱ��� = Mid(tbWeekTime.SelectedItem.Caption, InStr(tbWeekTime.SelectedItem.Caption, "(") + 1, 2) Then
            If Val(Nvl(mrs�ϴμƻ�ʱ��!��ſ���, "0")) = chk��ſ��� Then
                strMsg = "������������ʱ��,�Ƿ���ȡ���ŵ�ʱ����Ϊ�ƻ���ʱ����Ϣ? " & vbCrLf
                strMsg = strMsg & "[��(Y)]��ȡ���ŵ�ʱ����Ϣ��Ϊ�ƻ���ʱ��" & vbCrLf
                strMsg = strMsg & "[��(N)]����ȡ���ŵ�ʱ��,��������ʱ��" & vbCrLf
                vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
                If vMsgResult = vbYes Then
                    Set rs�Ӻ� = New Recordset
                    mbln׷�Ӻ� = True
                    mrs�޺�.Filter = " ���� ='" & mstrKey & "'"
                    rs�Ӻ�.Fields.Append "����", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "����", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "����", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "ʱ��", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "���", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "ʱ�䷶Χ", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "��������", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "�Ƿ�ԤԼ", adLongVarChar, 100
                    rs�Ӻ�.Fields.Append "��ſ���", adLongVarChar, 100
                    rs�Ӻ�.CursorLocation = adUseClient
                    rs�Ӻ�.LockType = adLockOptimistic
                    rs�Ӻ�.CursorType = adOpenDynamic
                    rs�Ӻ�.Open
                    
                    If Val(Nvl(mrs�޺�!�޺���, 0)) < mrs�ϴμƻ�ʱ��.RecordCount Then
                        '����,����,ʱ��,���,ʱ�䷶Χ,��������,�Ƿ�ԤԼ,��ſ���
                        For i = 0 To Val(Nvl(mrs�޺�!�޺���, 0)) - 1
                                rs�Ӻ�.AddNew
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!ʱ�� = mrs�ϴμƻ�ʱ��!ʱ��
                                rs�Ӻ�!��� = mrs�ϴμƻ�ʱ��!���
                                rs�Ӻ�!ʱ�䷶Χ = mrs�ϴμƻ�ʱ��!ʱ�䷶Χ
                                rs�Ӻ�!�������� = mrs�ϴμƻ�ʱ��!��������
                                rs�Ӻ�!�Ƿ�ԤԼ = mrs�ϴμƻ�ʱ��!�Ƿ�ԤԼ
                                rs�Ӻ�!��ſ��� = mrs�ϴμƻ�ʱ��!��ſ���
                           mrs�ϴμƻ�ʱ��.MoveNext
                        Next
                    Else
                        mrs�ϴμƻ�ʱ��.MoveFirst
                        For i = 1 To Val(Nvl(mrs�޺�!�޺���, 0))
                           If i <= mrs�ϴμƻ�ʱ��.RecordCount Then
                                rs�Ӻ�.AddNew
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!ʱ�� = mrs�ϴμƻ�ʱ��!ʱ��
                                rs�Ӻ�!��� = mrs�ϴμƻ�ʱ��!���
                                rs�Ӻ�!ʱ�䷶Χ = mrs�ϴμƻ�ʱ��!ʱ�䷶Χ
                                rs�Ӻ�!�������� = mrs�ϴμƻ�ʱ��!��������
                                rs�Ӻ�!�Ƿ�ԤԼ = mrs�ϴμƻ�ʱ��!�Ƿ�ԤԼ
                                rs�Ӻ�!��ſ��� = mrs�ϴμƻ�ʱ��!��ſ���
                                If i = mrs�ϴμƻ�ʱ��.RecordCount Then
                                    str���ʱ�䷶Χ = mrs�ϴμƻ�ʱ��!ʱ�䷶Χ
                                    lngĬ�ϼ��ʱ�� = DateDiff("n", CDate(Format(Split(str���ʱ�䷶Χ, "-")(0), "HH:mm")), CDate(Format(Split(str���ʱ�䷶Χ, "-")(1), "HH:mm")))
                                    mrs�ϰ�ʱ���.Filter = "ʱ��� ='" & Mid(tbWeekTime.SelectedItem.Caption, InStr(tbWeekTime.SelectedItem.Caption, "(") + 1, 2) & "'"
                                    If mrs�ϰ�ʱ���.RecordCount = 2 Then
                                        mrs�ϰ�ʱ���.Filter = "��ǩ='��-����'"
                                        bln����ʱ�� = Format("1900/1/1 " & DateAdd("n", lngĬ�ϼ��ʱ��, Split(str���ʱ�䷶Χ, "-")(1)), "yyyy-MM-dd hh:mm:ss") <= Format(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00"), "yyyy-MM-dd hh:mm:ss")
                                    End If
                                End If
                                mrs�ϴμƻ�ʱ��.MoveNext
                           Else
                                mrs�ϴμƻ�ʱ��.MoveLast
                                rs�Ӻ�.AddNew
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!���� = mrs�ϴμƻ�ʱ��!����
                                rs�Ӻ�!ʱ�� = mrs�ϴμƻ�ʱ��!ʱ��
                                rs�Ӻ�!��� = i
                                If bln����ʱ�� = True Then
                                    mrs�ϰ�ʱ���.Filter = "ʱ��� ='" & Mid(tbWeekTime.SelectedItem.Caption, InStr(tbWeekTime.SelectedItem.Caption, "(") + 1, 2) & "'"
                                    If mrs�ϰ�ʱ���.RecordCount = 2 Then
                                        mrs�ϰ�ʱ���.Filter = "��ǩ='��-����'"
                                        If Format("1900/1/1 " & DateAdd("n", lngĬ�ϼ��ʱ��, Split(str���ʱ�䷶Χ, "-")(1)), "yyyy-MM-dd hh:mm:ss") > Format(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00"), "yyyy-MM-dd hh:mm:ss") Then
                                            mrs�ϰ�ʱ���.Filter = ""
                                            mrs�ϰ�ʱ���.Filter = "��ǩ='��-����'"
                                            str���ʱ�䷶Χ = Format(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"), "HH:mm") & "-" & Format(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"), "HH:mm")
                                            bln����ʱ�� = False
                                        End If
                                    End If
                                End If
                                rs�Ӻ�!ʱ�䷶Χ = Format(Split(str���ʱ�䷶Χ, "-")(1), "hh:mm") & "-" & Format(DateAdd("n", lngĬ�ϼ��ʱ��, Split(str���ʱ�䷶Χ, "-")(1)), "HH:mm")
                                str���ʱ�䷶Χ = rs�Ӻ�!ʱ�䷶Χ
                                '�����:55628
                                rs�Ӻ�!�������� = 0 'mrs�ϴμƻ�ʱ��!��������
                                rs�Ӻ�!�Ƿ�ԤԼ = mrs�ϴμƻ�ʱ��!�Ƿ�ԤԼ
                                rs�Ӻ�!��ſ��� = mrs�ϴμƻ�ʱ��!��ſ���
                           End If
                        Next
                        
                    End If
                    str���ʱ�䷶Χ = ""
                    Set mrsTime = rs�Ӻ�
                    LoadEditTimePlan mlng�ƻ�Id, chk��ſ��� = 1
                    setVsFlexBgColor (chk��ſ���.Value = 1)
                    Exit Sub
                End If
            End If
         End If
        End If
        rs��ǰ�ƻ�ʱ��.Filter = ""
        mrs�ϴμƻ�ʱ��.Filter = ""
        mrs�޺�.Filter = ""
        Set mrsTime = rs��ǰ�ƻ�ʱ��
    End If
    
    Select Case mViewMode
        Case ViewMode.ViewItem:
             Call LoadTimePlan(mlng�ƻ�Id, Me.chk��ſ���.Value = 1)
        Case ViewMode.Edit:
            cmdԤԼ.Visible = False
            cmdɾ��.Visible = False
            Call LoadEditTimePlan(mlng�ƻ�Id, Me.chk��ſ���.Value = 1)
    End Select
    setVsFlexBgColor (chk��ſ���.Value = 1)
End Sub


 

Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
   
    '���Ʒ���������
    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    If txtTimeOut.Text = "" And KeyAscii = Asc(0) Then KeyAscii = 0
End Sub

Private Sub txtTimeOut_Validate(Cancel As Boolean)
    If Val(txtTimeOut.Text) < 1 Then Cancel = True
End Sub

Private Sub udTime_DownClick()
    If Val(txtTimeOut.Text) < 2 Then Exit Sub
    txtTimeOut.Text = Val(txtTimeOut.Text) - 1
End Sub

Private Sub udTime_UpClick()
  txtTimeOut.Text = Val(txtTimeOut.Text) + 1
End Sub


 
 
'Private Sub vsTime_Click()
'  Select Case mViewMode
'    Case ViewMode.Edit, ViewMode.NewItem:
'       If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Or (chk��ſ���.Value = 0 And vsTime.MouseRow < 1) Then Exit Sub
'       Select Case chk��ſ���.Value = 1
'            Case True:
'            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            Case False:
'            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'       End Select
'        If vsTime.MouseRow < 0 Or vsTime.MouseCol < 1 Then Exit Sub
'
'        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
'            cmdԤԼ.Left = vsTime.MouseCol * 1200 + 20
'            cmdԤԼ.Top = vsTime.MouseRow * 400 + 20
'            cmdԤԼ.Visible = True
'        End If
'
'    Case ViewMode.ViewItem:
'         vsTime.Editable = flexEDNone
'  End Select
'End Sub

Public Function ShowMe(lng�ƻ�ID As Long, mode As ViewMode) As Boolean
    mViewMode = mode: mlng�ƻ�Id = lng�ƻ�ID
    If InitData() = False Then
        '���عҺŰ��żƻ�������Ϣ
         Exit Function
    End If
    Select Case mViewMode
         Case ViewMode.ViewItem:
                vsTime.Editable = flexEDNone
                Me.txtTimeOut.Enabled = False
                Me.cmd����ʱ��.Enabled = False
               '�鿴
              Call LoadTimePlan(mlng�ƻ�Id, chk��ſ���.Value = 1, False)
         Case ViewMode.Edit
              If LoadEditTimePlan(mlng�ƻ�Id, chk��ſ���.Value = 1, False) = False Then
                Exit Function
              End If
    End Select
    setVsFlexBgColor (chk��ſ���.Value = 1)
    Me.Show 1
    ShowMe = mblnReload
End Function
'------------------------------------------------------------------------
'ҳ����ù����뷽��
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL          As String
    Dim lng�ƻ�ID       As Long
    If mlng�ƻ�Id = -1 Then Exit Function
     lng�ƻ�ID = mlng�ƻ�Id
     On Error GoTo Hd
     strSQL = " " & _
        "   Select a.Id as ����ID,a.�ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,NVL(A.Ĭ��ʱ�μ��,5) as Ĭ��ʱ�μ��, " & _
        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
        "   From ( " & vbNewLine & _
        "       Select B.ID,a.id As �ƻ�id, B.����, A.����, B.����id, A.��Ŀid, B.ҽ������, B.ҽ��id, A.����, A.��һ, A.�ܶ�, A.����," & _
        "              A.����, A.����, A.����, B.��������, A.���﷽ʽ, A.��ſ���, A.��Чʱ�� As ��ʼʱ��, A.ʧЧʱ�� As ��ֹʱ��,A.Ĭ��ʱ�μ�� As Ĭ��ʱ�μ�� " & _
        "        From �ҺŰ��� B, �ҺŰ��żƻ� A " & _
        "       Where A.����id = B.ID And A.Id=[1] " & _
        ") A,�շ���ĿĿ¼ B,���ű� D " & _
        "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
        "        "
         Set mrs���żƻ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
         
         If mrs���żƻ�.EOF Then
              ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
             Exit Function
        End If
        strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �Һżƻ����� where �ƻ�ID=[1]  Order BY ������Ŀ      "
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
        cbo����.Text = Nvl(mrs���żƻ�!����)
        txt�ű�.Tag = Nvl(mrs���żƻ�!�ƻ�Id)
        txt�ű�.Text = Nvl(mrs���żƻ�!����)
        txtTimeOut.Tag = Val(Nvl(mrs���żƻ�!Ĭ��ʱ�μ��, 5))
        txtTimeOut.Text = txtTimeOut.Tag
        cbo����.Text = Nvl(mrs���żƻ�!����)
        cboItem.Text = Nvl(mrs���żƻ�!��Ŀ)
        cboDoctor.Text = Nvl(mrs���żƻ�!ҽ������)
        chk����.Value = IIf(Val(Nvl(mrs���żƻ�!��������)) = 1, 1, 0)
       chk��ſ���.Value = IIf(Val(Nvl(mrs���żƻ�!��ſ���)) = 1, 1, 0):  chk��ſ���.Tag = chk��ſ���.Value
        strSQL = "" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ" & _
        "   From  �Һżƻ�ʱ�� " & _
        "   Where �ƻ�ID=[1]" & _
        "   Order by ����,ʱ��,���"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID)
       InitData = True
Exit Function
Hd:
     If ErrCenter() = 1 Then Resume
     SaveErrLog
End Function

 
Private Function LoadEditTimePlan(ByVal lng�ƻ�ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim j                As Long
    Dim r                As Integer
    Dim lngRow           As Long
    Dim lngCol           As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
    Dim lng�ѹ�������  As Long
     
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        mlngPre�ƻ�ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre�ƻ�ID = -1
    End If
    If mlngPre�ƻ�ID <> lng�ƻ�ID Then
        mlngPre�ƻ�ID = lng�ƻ�ID
        tbWeekTime.Tabs.Clear
         With tbWeekTime
            If Not mrs�޺�.EOF Then
                mrs�޺�.Filter = "����='��һ'"
                If mrs�޺�.RecordCount > 0 Then
                '�޺���,  ��Լ��,������Ŀ
                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K��һ", "��һ" & IIf(Nvl(mrs���żƻ�!��һ) = "", "", "(" & Nvl(mrs���żƻ�!��һ) & ")")
                    End If
                End If
                mrs�޺�.Filter = "����='�ܶ�'"
                If mrs�޺�.RecordCount > 0 Then
                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K�ܶ�", "�ܶ�" & IIf(Nvl(mrs���żƻ�!�ܶ�) = "", "", "(" & Nvl(mrs���żƻ�!�ܶ�) & ")")
                    End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                    End If
                 End If
                 
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                  If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                      "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                  End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                     End If
                End If
                
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                          "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                   End If
                End If
                mrs�޺�.Filter = "����='����'"
                If mrs�޺�.RecordCount > 0 Then
                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                    End If
                End If
                mrs�޺�.Filter = 0
            End If
            .Visible = tbWeekTime.Tabs.Count <> 0
            If .Tabs.Count > 0 Then
                .Tabs(1).Selected = True
            Else
                MsgBox "�üƻ�û�����ö�Ӧ���޺�������Լ��,����!", vbOKOnly, Me.Caption
                Exit Function
            End If
            
'            If Not mrs�޺�.EOF Then
'                mrs�޺�.Filter = "����='��һ'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                        "K��һ", "��һ" & IIf(Nvl(mrs���żƻ�!��һ) = "", "", "(" & Nvl(mrs���żƻ�!��һ) & ")")
'
'                mrs�޺�.Filter = "����='�ܶ�'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                        "K�ܶ�", "�ܶ�" & IIf(Nvl(mrs���żƻ�!�ܶ�) = "", "", "(" & Nvl(mrs���żƻ�!�ܶ�) & ")")
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                        "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                      "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                      "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                      "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then tbWeekTime.Tabs.Add , _
'                      "K����", "����" & IIf(Nvl(mrs���żƻ�!����) = "", "", "(" & Nvl(mrs���żƻ�!����) & ")")
                
'            End If
'            .Visible = tbWeekTime.Tabs.Count <> 0
'            If .Tabs.Count > 0 Then
'                .Tabs(1).Selected = True
'           End If
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ʱ���"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ԤԼ����"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mrsTime!��������))
                strTime = mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                  .Cell(flexcpForeColor, r, i * 2, r, i * 2 + 1) = vbBlue
                End If
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
            LoadEditTimePlan = True
            Exit Function
        End If
        .Cols = 7: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        i = 1: r = -1
        lngRow = -1: lngCol = 1
        '******************************************
        With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         '***********************
         '������
         '**********************
         r = mrsTime.RecordCount
         For i = 1 To r
            If mrsTime.EOF Then Exit For
            lngCol = lngCol + 1
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                strʱ�� = Nvl(mrsTime!ʱ��)
                If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
                vsTime.TextMatrix(lngRow, 0) = strʱ��
             End If
            strData = mrsTime!���
            strTime = mrsTime!ʱ�䷶Χ
            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
            'If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
            vsTime.TextMatrix(lngRow, lngCol) = strTime
            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
            If lngCol = 1 Then
            End If
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            mrsTime.MoveNext
         Next
         
         End With
        '******************************************
'        Do While Not mrsTime.EOF
'            If i = 1 Then
'                r = r + 2
'                strʱ�� = Nvl(mrsTime!ʱ��)
'                If r > .Rows - 1 Then .Rows = .Rows + 2
'                .TextMatrix(r, 0) = strʱ��
'                .TextMatrix(r - 1, 0) = strʱ��
'            End If
'            i = i + 1
'            strData = mrsTime!���
'            strTime = mrsTime!ʱ�䷶Χ
'            If i >= .Cols - 1 Then i = 1
'            If r > .Rows - 1 Then .Rows = .Rows + 2
'            .TextMatrix(r, i) = strTime
'            .TextMatrix(r - 1, i) = strData
'
'        Loop
        
        
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        End If
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    
    '�����:51427
    lng�ѹ������� = ExistsBooking(mlng�ƻ�Id, Mid(tbWeekTime.SelectedItem.Key, 2))
    '�����:51427
    If chk��ſ���.Value = 1 Then
        For i = 0 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1
                If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                    If CLng(vsTime.TextMatrix(i, j)) <= lng�ѹ������� Then
                        vsTime.Cell(flexcpForeColor, i, j) = &HC0C0C0
                        vsTime.Cell(flexcpForeColor, i + 1, j) = &HC0C0C0
                    End If
                End If
            Next
        Next
    End If
    
    LoadEditTimePlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
 
Private Sub LoadEditTimePlantext(ByVal lng�ƻ�ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim r                As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
     
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
        mlngPre�ƻ�ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre�ƻ�ID = -1
    End If
    If mlngPre�ƻ�ID <> lng�ƻ�ID Then
        mlngPre�ƻ�ID = lng�ƻ�ID
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!����) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
                    strTime = Nvl(mrsTime!����)
                End If
                .MoveNext
            Loop
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ʱ���"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "ԤԼ����"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                If i * 2 > .Cols - 2 Then r = r + 1: i = -1
                i = i + 1
                strData = Val(Nvl(mrsTime!��������))
                strTime = mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
                r = r + 2
                strʱ�� = Nvl(mrsTime!ʱ��)
                If r > .Rows - 1 Then .Rows = .Rows + 2
                .TextMatrix(r, 0) = strʱ��
                .TextMatrix(r - 1, 0) = strʱ��
                i = 0
            End If
            i = i + 1
            strData = mrsTime!���
            strTime = mrsTime!ʱ�䷶Χ
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            .TextMatrix(r - 1, i) = strData
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                 
                .Cell(flexcpForeColor, r - 1, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r - 1, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
 
Private Sub LoadTimePlan(ByVal lng�ƻ�ID As Long, ByVal bln��ſ��� As Boolean, _
    Optional bln�ƻ� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:
    '���:
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str����          As String
    Dim i                As Long
    Dim r                As Integer
    Dim strʱ��          As String
    Dim strTime          As String
    Dim strKey           As String
    On Error GoTo errHandle
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    If mrsTime Is Nothing Then
         mlngPre�ƻ�ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre�ƻ�ID = -1
    End If
    If mlngPre�ƻ�ID <> lng�ƻ�ID Then
        mlngPre�ƻ�ID = lng�ƻ�ID
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!����) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
                    strTime = Nvl(mrsTime!����)
                End If
                .MoveNext
            Loop
           
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
           
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
    End If
    str���� = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "����='" & str���� & "'"
    mrs�޺�.Filter = "����='" & str���� & "'"
    txt�޺�.Text = ""
    txt��Լ.Text = ""
    If mrs�޺�.RecordCount <> 0 Then
        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
    End If
     strʱ�� = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln��ſ��� Then
             .Cols = 8: .FixedCols = 0
             r = 0: i = 0
            Do While Not mrsTime.EOF
               i = i + 1
                If i > .Cols - 1 Then r = r + 1: i = 0
                strTime = "ԤԼ" & Val(Nvl(mrsTime!��������)) & "��" & vbCrLf & vbCrLf
                strTime = strTime & mrsTime!ʱ�䷶Χ
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i) = strTime
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
                r = r + 1
                strʱ�� = Nvl(mrsTime!ʱ��)
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = strʱ��
                i = 0
            End If
            i = i + 1
            strTime = mrsTime!��� & vbCrLf & vbCrLf
            strTime = strTime & mrsTime!ʱ�䷶Χ
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
    
Private Sub vsTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
 If vsTime.Row < 0 Or vsTime.Col < 0 Or (chk��ſ���.Value = 0 And vsTime.Row < 1) Then cmdԤԼ.Visible = False: mblnCellChange = False: Exit Sub
    '�����:51429
    SetCtrlMove
    Select Case mViewMode
    Case ViewMode.Edit, ViewMode.NewItem:
       Select Case chk��ſ���.Value = 1
            Case True:
            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '�������������ʽ
            '******************************************
            If vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
            Case False:
            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '�������������ʽ
            '******************************************
            If NewCol Mod 2 = 0 And vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
       End Select
        If vsTime.Row < 0 Or vsTime.Col < 1 Then Exit Sub
        
        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
            mblnCellChange = True
        Else
           mblnCellChange = False
        End If
        
    Case ViewMode.ViewItem:
         mblnCellChange = False
         vsTime.Editable = flexEDNone
  End Select
End Sub
Private Sub vsTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmdɾ��.Visible = False Then Exit Sub
    If KeyCode = 46 Then '��ݼ�Delete
            Call DeleteSelectPain
    End If
End Sub
Private Sub vsTime_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    '**************************************************************
    '������Ա �϶�������ʱ �� ԤԼ��ť ����
    '**************************************************************
    Me.cmdԤԼ.Visible = False
    Me.cmdɾ��.Visible = False
End Sub

Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If mViewMode = ViewItem Then Exit Sub
    Select Case chk��ſ���.Value = 1
        Case True:
            '******************************************
            'ר�Һ�ʱ ��������
            '******************************************
            If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
        Case False:
            '******************************************
            '��ͨ��ʱ ��������
            '******************************************
            If Col Mod 2 = 0 Then
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
            Else
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
            End If
            
    End Select
   
 
End Sub
 
Private Function validateVsFlex() As Boolean
    '***************************************
    '��֤�û��ԹҺżƻ�ʱ�ε��޸�
    '***************************************
     Dim i          As Long
     Dim j          As Long
     Dim lngԤԼ    As Long
     Dim lng��Լ    As Long
     Dim lng�޺�    As Long
     Dim str����    As String
     If tbWeekTime.SelectedItem Is Nothing Then Exit Function
      str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
     lng�޺� = Val(txt�޺�.Text)
     lng��Լ = Val(txt��Լ.Text)
     If lng��Լ = 0 Then lng��Լ = lng�޺�
     Select Case chk��ſ���.Value = 1
     Case True:
     '*************************************
     'ר�Һż����Լ���Ƿ�����޺���
     '*************************************
        With vsTime
            For i = 0 To .Rows - 1 Step 2
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j, i, j) = vbBlue And .TextMatrix(i, j) <> "" Then
                        lngԤԼ = lngԤԼ + 1
                    End If
                Next
            Next
        End With
     Case False:
     '*************************************
     '��ͨ�ż����Լ���Ƿ�����޺���
     '*************************************
        With vsTime
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) <> "" Then
                        lngԤԼ = lngԤԼ + Val(.TextMatrix(i, j))
                    End If
                Next
            Next
        End With
     End Select
     If lngԤԼ > lng��Լ Then
        MsgBox "��" & str���� & "���õ�ԤԼ��" & lngԤԼ & "������" & IIf(lng�޺� = lng��Լ, "�޺���" & lng��Լ, "��Լ��" & lng��Լ) & ",����!", vbOKOnly, Me.Caption
        Exit Function
     End If
    validateVsFlex = True
    Exit Function
End Function

Private Function SaveDate() As Boolean
    '*********************************
    '�ԹҺżƻ�ʱ�ν��б���
    '*********************************
    Dim strSQL      As String
    Dim cllSQL      As Collection
    Dim i           As Long
    Dim j           As Long
    Dim blnTrans    As Boolean
    Dim lng�ƻ�ID   As Long
    Dim str����     As String
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim blnԤԼ     As Boolean
    Dim lng����     As Long '�Һżƻ�ʱ�ε���������
    Dim blnר�Һ�   As Boolean
    Dim lng���     As Long
    Dim lngType     As Long
    Dim str��ֹʱ�� As String '�����:55555
    Dim str��ʼʱ�� As String '�����:55555
    Dim strMsg As String '�����:55555
    Dim vMsgResult As VbMsgBoxResult '�����:55555
    If validateVsFlex() = False Then Exit Function '�������ݵ���֤
    
    
    lng�ƻ�ID = Val(txt�ű�.Tag)
    str���� = mstrKey
    blnר�Һ� = chk��ſ���.Value = 1
    
    Set cllSQL = New Collection
    '****************************************************
    'CREATE OR REPLACE Procedure Zl_�Һżƻ�ʱ��_Delete(
    '�ƻ�ID_In �Һżƻ�ʱ��.�ƻ�ID%Type,
    '����_In   �Һżƻ�ʱ��.���� %Type)
    '**********ɾ����ǰ�Դ����ڰ��żƻ���ʱ��*****************
    strSQL = "Zl_�Һżƻ�ʱ��_Delete(" & lng�ƻ�ID & ",'" & str���� & "')"
    zlAddArray cllSQL, strSQL
    
   
    Select Case blnר�Һ�
    Case True:
       lng��� = 0
       For i = 1 To vsTime.Rows - 1 Step 2
            For j = 1 To vsTime.Cols - 1
               If vsTime.TextMatrix(i, j) = "" Then Exit For
               str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
               str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
               str��ʼʱ�� = Split(vsTime.TextMatrix(i, j), "-")(0)
               str��ֹʱ�� = Split(vsTime.TextMatrix(i, j), "-")(1)
               '�����:55555
               If Check���ʱ��(Mid(tbWeekTime.SelectedItem.Caption, InStr(tbWeekTime.SelectedItem.Caption, "(") + 1, 2), str��ʼʱ��, str��ֹʱ��) = True Then
                    strMsg = "��ǰ�����ʱ�����������õ�ʱ�䳬�����°�ʱ��,���Ƿ�ȷ��Ҫ���������? " & vbCrLf
                    strMsg = strMsg & "[��(Y)]������Ч���ϰ�ʱ������" & vbCrLf
                    strMsg = strMsg & "[��(N)]������,��������" & vbCrLf
                    vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
                    If vMsgResult = vbYes Then
                        GoTo ����ʱ��
                    Else
                        Exit Function
                    End If
               End If
               lng���� = 1
               lng��� = lng��� + 1
               blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
               strSQL = GetInsertSql(lng�ƻ�ID, lng���, str��ʼʱ��, str����ʱ��, 1, blnԤԼ, str����)
               zlAddArray cllSQL, strSQL
            Next
       Next
    Case False:
        lng��� = 0
        For i = 1 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1 Step 2
               If vsTime.TextMatrix(i, j) <> "" Then
                str��ʼʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
                str����ʱ�� = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
                '�����:55555
                str��ʼʱ�� = Split(vsTime.TextMatrix(i, j), "-")(0)
                str��ֹʱ�� = Split(vsTime.TextMatrix(i, j), "-")(1)
                If Check���ʱ��(Mid(tbWeekTime.SelectedItem.Caption, InStr(tbWeekTime.SelectedItem.Caption, "(") + 1, 2), str��ʼʱ��, str��ֹʱ��) = True Then
                    strMsg = "��ǰ�����ʱ�����������õ�ʱ�䳬�����°�ʱ��,���Ƿ�ȷ��Ҫ���������? " & vbCrLf
                    strMsg = strMsg & "[��(Y)]������Ч���ϰ�ʱ������" & vbCrLf
                    strMsg = strMsg & "[��(N)]������,��������" & vbCrLf
                    vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
                    If vMsgResult = vbYes Then
                        GoTo ����ʱ��
                    Else
                        Exit Function
                    End If
                End If

                lng���� = Val(vsTime.TextMatrix(i, j + 1))
                lng��� = lng��� + 1
                blnԤԼ = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
                strSQL = GetInsertSql(lng�ƻ�ID, lng���, str��ʼʱ��, str����ʱ��, lng����, blnԤԼ, str����)
                zlAddArray cllSQL, strSQL
               End If
            Next
        Next
    End Select
����ʱ��:
    
    If opt��ҽ��.Value Then
        lngType = 1
    ElseIf opt����.Value Then
        lngType = 2
    ElseIf opt����.Value Then
        lngType = 3
    End If
    If lngType <> 0 Then
        '--type_in
        '--1-Ӧ���뱾��
        '--2-Ӧ���뱾����
        '--3 or others -Ӧ��������
       'CREATE OR REPLACE Procedure zl_�ҺŰ���ʱ��_����Ӧ��
       strSQL = "zl_�Һżƻ�ʱ��_����Ӧ��("
       '����Id_in �ҺŰ���ʱ��.����Id%Type,
       strSQL = strSQL & lng�ƻ�ID & ","
       'Type_In Number:=1
       strSQL = strSQL & lngType & ")"
       zlAddArray cllSQL, strSQL
    End If
    
  On Error GoTo ErrHand
    gcnOracle.BeginTrans
    
    For i = 1 To cllSQL.Count
        strSQL = cllSQL(i)
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans
    SaveDate = True
 Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    SaveErrLog
    
End Function

Private Function GetInsertSql(ByVal lngID As Long, ByVal lng��� As Long, ByVal str��ʼʱ�� As String, _
        ByVal str����ʱ�� As String, ByVal lng�������� As Long, ByVal bln�Ƿ�ԤԼ As Boolean, ByVal str���� As String)
    '�����ṩ����Ϣ����sql���
    Dim strSQL      As String
   '********************************************************
    '    'CREATE OR REPLACE Procedure Zl_�Һżƻ�ʱ��_Insert
    '    (
    '    �ƻ�ID_In   �Һżƻ�ʱ��.�ƻ�ID%Type,
    '    ���_In     �Һżƻ�ʱ��.���%Type,
    '    ��ʼʱ��_In �Һżƻ�ʱ��.��ʼʱ��%Type,
    '    ����ʱ��_In �Һżƻ�ʱ��.����ʱ��%Type,
    '    ��������_In �Һżƻ�ʱ��.��������%Type,
    '    �Ƿ�ԤԼ_In �Һżƻ�ʱ��.�Ƿ�ԤԼ%Type,
    '    ����_In     �Һżƻ�ʱ��.����%Type
    '    )
    '********************************************************
    strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
     '�ƻ�ID_In   �Һżƻ�ʱ��.�ƻ�ID%Type,
    strSQL = strSQL & lngID & ","
     '���_In     �Һżƻ�ʱ��.���%Type,
    strSQL = strSQL & lng��� & ","
     '��ʼʱ��_In �Һżƻ�ʱ��.��ʼʱ��%Type,
     strSQL = strSQL & str��ʼʱ�� & ","
      '����ʱ��_In �Һżƻ�ʱ��.����ʱ��%Type,
    strSQL = strSQL & str����ʱ�� & ","
      '��������_In �Һżƻ�ʱ��.��������%Type,
    strSQL = strSQL & lng�������� & ","
     '�Ƿ�ԤԼ_In �Һżƻ�ʱ��.�Ƿ�ԤԼ%Type,
    strSQL = strSQL & IIf(bln�Ƿ�ԤԼ, 1, 0) & ","
     '����_In     �Һżƻ�ʱ��.����%Type
    strSQL = strSQL & "'" & str���� & "')"
    GetInsertSql = strSQL
End Function

                             

Private Function ConvertToDate(ByVal strDate As String, Optional ByVal haveYear = False) As String
    '**********************************************************
    '���ַ���ת����oracle���ݿ��ܹ�ʶ�������
    '**********************************************************
    Select Case haveYear
    Case True:
        ConvertToDate = "To_Date('" & strDate & "', 'YYYY-MM-DD HH24:MI:SS')"
    Case False:
        ConvertToDate = "To_Date('" & strDate & "', 'HH24:MI:SS')"
    End Select
End Function



Private Sub vsTime_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng�޺�   As Long
  Dim lng��Լ   As Long
  Dim lngԤԼ�� As Long
  If mViewMode = ViewItem Then Exit Sub

   '*************************************
  'ʱ�������֤ ������ʱ�䷶Χ
  '**************************************
  If vsTime.Editable = flexEDKbdMouse And vsTime.ColEditMask(vsTime.Col) = strMaskKey Then
    Validateʱ�� Row, Col, Cancel
    If Not Cancel Then mblnChange = True
    Exit Sub
  End If
  '****************************************
  '����ͨ�� ��ʱ�� �����������ԤԼ����������
  '****************************************
   If chk��ſ���.Value = 0 And vsTime.ColEditMask(vsTime.Col) <> strMaskKey And vsTime.Editable = flexEDKbdMouse Then
        If vsTime.EditText = "" Then vsTime.EditText = "0"
        mblnChange = True
   End If
End Sub

Private Sub Validateʱ��(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng�޺�   As Long
  Dim lng��Լ   As Long
  Dim lngԤԼ�� As Long
   
  Dim strʱ��()  As String
  If mViewMode = ViewItem Then Exit Sub
  
  '*************************************
  '��֤ʱ��
  '**************************************
  strʱ�� = Split(vsTime.EditText, "-")
  If UBound(strʱ��) <> 1 Then Cancel = True: Exit Sub
   If Not IsDate(strʱ��(0)) Then Cancel = True: Exit Sub
   If Not IsDate(strʱ��(1)) Then Cancel = True: Exit Sub
   If CDate(strʱ��(0)) >= CDate(strʱ��(1)) Then
        MsgBox "��ʼʱ�����С�ڽ���ʱ��!����!", vbOKOnly, Me.Caption
        Cancel = True
   End If
   
End Sub

Private Sub setVsFlexBgColor(Optional ByVal bln��ſ��� As Boolean = False)
    '**************************************************************
    '��ʱ������ü������
    '**************************************************************
     Dim i           As Long
     If (bln��ſ��� And vsTime.Rows = 0) Or (bln��ſ��� = False And vsTime.Rows = 1) Then Exit Sub
     For i = IIf(bln��ſ���, 0, 1) To vsTime.Rows - 1 Step 2
            vsTime.Cell(flexcpBackColor, i, IIf(bln��ſ���, 1, 0), i, vsTime.Cols - 1) = &HE0E0D3
     Next
End Sub

Private Function ExistsBooking(ByVal lng�ƻ�ID As String, str���� As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�;str����-���ڼ��İ���
    '����:����,�������Һ����,�����ڷ���-1
    '����:
    '����:2012-04-26 10:32:02
    '�����:51657
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "" & _
    "   Select max(����) as ����  From ���˹Һż�¼ A, �ҺŰ��� B,�ҺŰ��żƻ� C" & _
    "   Where A.�ű� = B.���� And B.ID=C.����ID " & _
    "       And ��¼״̬ = 1 and C.id=[1]  " & _
    "       And Decode(To_Char(A.����ʱ��, 'D'), '1', '����', '2','��һ', '3', '�ܶ�', '4', '����', '5', '����', '6','����', '7', '����', Null) =[2]" & _
    "       And A.����ʱ�� >= Trunc(Sysdate)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID, str����)
    ExistsBooking = CLng(Nvl(rsTmp!����, "-1"))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmdOtherCalc_Click()
    Dim str���� As String
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    str���� = Replace(Split(tbWeekTime.SelectedItem.Caption & "(", "(")(1), ")", "")
    Call mfrmOtherCalc.zlShowMe(Me, str����, Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing
End Sub

Private Sub DeleteSelectPain()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ѡ�е�ʱ�����
    '����:����
    '����:2012-07-12 10:32:02
    '�����:51429
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
    Dim lng�ƻ�ID As Long
    Dim lng������ As Long
    Dim lng��ǰ���� As Long
    Dim lng��ǰ����к� As Long
    Dim blnDel As Boolean
    Dim i As Long
    Dim j As Long
    
    If chk��ſ���.Value <> 1 Then Exit Sub
    If vsTime.TextMatrix(vsTime.Row, vsTime.Col) = "" Then Exit Sub
    cmdɾ��.Visible = False
    cmdԤԼ.Visible = False
    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng�ƻ�ID = Val(txt�ű�.Tag)
    lng������ = ExistsBooking(lng�ƻ�ID, str����)
    
    '����Ƿ��Ǵ����ʼɾ��
    With vsTime
'         For i = 0 To vsTime.Rows - 1
'            For j = 0 To vsTime.Cols - 1
'                If IsNumeric(.TextMatrix(i, j)) = True Then
'                    If lng������ < IIf(.TextMatrix(i, j) = "", "0", .TextMatrix(i, j)) Then
'                        lng������ = .TextMatrix(i, j)
'                    End If
'                End If
'            Next
'         Next

'         If lng������ <> CLng(IIf(.TextMatrix(lng��ǰ����к�, .Col) = "", "0", .TextMatrix(lng��ǰ����к�, .Col))) Then
'                MsgBox "ֻ�ܴ����ĺ���ʼɾ����", vbInformation, Me.Caption
'                Exit Sub
'         End If
   
     If .Row Mod 2 = 0 Then
            lng��ǰ����к� = .Row
         Else
            lng��ǰ����к� = .Row - 1
     End If
     lng��ǰ���� = Val(.TextMatrix(lng��ǰ����к�, .Col))
    '����Ƿ�úű��Ѿ����ҳ�
     If lng������ >= lng��ǰ���� Then
                MsgBox lng������ & "���Ѿ��кű��ҳ�,ֻ��ɾ���ú��Ժ����ţ�", vbInformation, Me.Caption
                Exit Sub
     End If

     SetVsTime lng��ǰ����к�, .Col
     '��ո������Ϣ
     
'     .TextMatrix(lng��ǰ����к�, .Col) = ""
'     .TextMatrix(lng��ǰ����к� + 1, .Col) = ""
    End With
End Sub


Public Sub SetVsTime(lngRow As Long, lngCol As Long)
    Dim i As Long
    Dim j As Long
    Dim lng��ǰ��� As Long
    
    With vsTime
         lng��ǰ��� = Val(.TextMatrix(lngRow, .Col))
         .TextMatrix(lngRow, .Col) = ""
         .TextMatrix(lngRow + 1, .Col) = ""
         For i = lngRow + 2 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                        .TextMatrix(i, j) = lng��ǰ���
                         lng��ǰ��� = lng��ǰ��� + 1
                    End If
            Next
         Next
    End With
End Sub
Private Function Get�޺���(ByVal str���� As String, ByRef lng�޺��� As Long, ByRef lng��Լ�� As Long) As Boolean
    Dim strSQL As String
    If mrs�޺� Is Nothing Then
        strSQL = _
        "Select �ƻ�id, ������Ŀ as ���� , �޺���, ��Լ�� From �Һżƻ����� Where �ƻ�id = [1]"
        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt�ű�.Tag))
        If mrs�޺�.RecordCount = 0 Then
            MsgBox "��ǰ�ű�û�ж�Ӧ�ĹҺżƻ�����" & vbCrLf & "�뵽�Һżƻ�������!", vbOKOnly, Me.Caption
            Set mrs�޺� = Nothing
            Exit Function
        End If
    End If
    mrs�޺�.Filter = "����='" & str���� & "'"
    If mrs�޺�.RecordCount <> 0 Then
        lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
        lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
        Get�޺��� = True
    End If
End Function
Private Sub SetCtrlMove()
    Dim blnDel As Boolean
    With vsTime
         If chk��ſ���.Value = 1 Then
            cmdɾ��.Left = .CellLeft + .CellWidth - cmdɾ��.Width
            If .Row Mod 2 <> 0 Then
                cmdɾ��.Top = .CellTop - .CellHeight - 15
            Else
                cmdɾ��.Top = .CellTop + 15
            End If
            cmdԤԼ.Left = .CellLeft + 30
            cmdԤԼ.Top = cmdɾ��.Top
            If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
            Else
                blnDel = True
            End If
            blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> ""
            cmdɾ��.Visible = blnDel And chk��ſ���.Value = 1
            cmdԤԼ.Visible = Val(txt��Լ.Text) <> 0
         Else
            cmdԤԼ.Left = .CellLeft + 15
            cmdԤԼ.Top = .CellTop + 15
            cmdԤԼ.Visible = False 'Val(txt��Լ.Text) <> 0
         End If
    End With
    cmdԤԼ.Refresh
    cmdɾ��.Refresh
End Sub
Private Function Get��һ�μƻ�ʱ��() As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ϴμƻ�ʱ����Ϣ
    '����:����
    '����:2012-08-1 10:32:02
    '�����:52221
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strʱ�� As String
    Dim rs�ϴμƻ�ʱ�� As Recordset
    
    On Error GoTo errH:
    strʱ�� = "" & _
    "     Select Distinct A.Id,Decode(Nvl(A.Id,0),0,'��','����') As ����,A.��ſ���,A.����,A.��һ,A.�ܶ�,A.����,A.����,A.����,A.����" & _
    "     From �ҺŰ��� A,�ҺŰ���ʱ�� B,�ҺŰ��żƻ� C " & _
    "     Where a.ͣ������ Is Null " & _
    "     And A.ID=B.����ID " & _
    "     And A.ID = C.����ID " & _
    "     And C.Id=[1] " & _
    "     And Not Exists " & _
    "          (Select 1" & _
    "               From �ҺŰ��żƻ� d,�Һżƻ�ʱ�� E" & _
    "               Where d.����id = a.Id And d.���ʱ�� Is Not Null And d.ID=E.�ƻ�ID And" & _
    "                     Sysdate Between Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & _
    "                     d.ʧЧʱ�� " & _
    "                And Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = " & _
    "               (Select Max(a.��Чʱ��) As ��Ч " & _
    "                From �ҺŰ��żƻ� a,(Select Count(1) as ����� From �ҺŰ��żƻ� A,�ҺŰ��żƻ� B Where A.ID=[1] And A.����ID=B.����ID And B.���ʱ�� Is Not Null) K" & _
    "                Where a.���ʱ�� Is Not Null And K.�����>1 And" & _
    "                      Sysdate Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And " & _
    "                      a.ʧЧʱ�� And a.����id =d.����id)) "
    strʱ�� = strʱ�� & " " & _
    "      Union All" & _
    "      Select Distinct A.Id,Decode(Nvl(A.Id,0),0,'��','�ƻ�') As ����,A.��ſ���,A.����,A.��һ,A.�ܶ�,A.����,A.����,A.����,A.����" & _
    "      From �ҺŰ��żƻ� a, �Һżƻ�ʱ�� b,(Select C.����Id,C.ID,C.��ſ��� From �ҺŰ��żƻ� C Where C.Id=[1]) D" & _
    "      Where a.����Id=D.����ID And a.���ʱ�� Is Not Null And " & _
    "      a.Id=b.�ƻ�Id  And" & _
    "                Sysdate Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And" & _
    "                a.ʧЧʱ�� And A.Id Not In D.Id" & _
    "                And Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = " & _
    "               (Select Max(a.��Чʱ��) As ��Ч " & _
    "                From �ҺŰ��żƻ� a,(Select Count(1) as ����� From �ҺŰ��żƻ� A,�ҺŰ��żƻ� B Where A.ID=[1] And A.����ID=B.����ID And B.���ʱ�� Is Not Null) K" & _
    "                Where a.���ʱ�� Is Not Null And K.�����>1 And" & _
    "                      Sysdate Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And " & _
    "                      a.ʧЧʱ�� And a.����id =d.����id) "
    
    '��ȡ��һ�μƻ�ʱ��
        strSQL = "" & _
        "   Select ����,����,ʱ��,���,ʱ�䷶Χ,��������,�Ƿ�ԤԼ,��ſ���,ʱ��� From (" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ,B.��ſ���," & _
        "   decode(����,'����',B.����,'��һ',B.��һ,'�ܶ�',B.�ܶ�,'����',B.����,'����',B.����,'����',B.����,B.����) as ʱ���" & _
        "   From  �ҺŰ���ʱ�� A,(" & strʱ�� & ") B" & _
        "   Where ����ID=Decode(B.����,'����',B.ID,0)" & _
        "   Order by ����,ʱ��,���)"
        strSQL = strSQL & " Union All " & _
        "   Select ����,����,ʱ��,���,ʱ�䷶Χ,��������,�Ƿ�ԤԼ,��ſ���,ʱ��� From (" & _
        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               ��������,�Ƿ�ԤԼ,B.��ſ���," & _
        "   decode(����,'����',B.����,'��һ',B.��һ,'�ܶ�',B.�ܶ�,'����',B.����,'����',B.����,'����',B.����,B.����) as ʱ���" & _
        "   From  �Һżƻ�ʱ�� ,(" & strʱ�� & ") B" & _
        "   Where �ƻ�ID=Decode(B.����,'�ƻ�',B.ID,0)" & _
        "   Order by ����,ʱ��,���)"
    Set rs�ϴμƻ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
    
    Set Get��һ�μƻ�ʱ�� = rs�ϴμƻ�ʱ��
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get��ǰ�ƻ�ʱ��() As Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ�ƻ�ʱ����Ϣ
    '����:����
    '����:2012-08-1 10:32:02
    '�����:52221
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH:
    strSQL = "" & _
        "   Select decode(sd.����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,sd.����,to_char(sd.��ʼʱ��,'HH24')||':00' as ʱ��,sd.���,to_char(sd.��ʼʱ��,'hh24:mi')||'-' ||to_char(sd.����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
        "               sd.��������,sd.�Ƿ�ԤԼ," & _
        "   decode(sd.����,'����',jh.����,'��һ',jh.��һ,'�ܶ�',jh.�ܶ�,'����',jh.����,'����',jh.����,'����',jh.����,jh.����) as ʱ���" & _
        "   From  �Һżƻ�ʱ�� sd,�ҺŰ��żƻ� jh" & _
        "   Where sd.�ƻ�ID=[1] And sd.�ƻ�ID=jh.ID" & _
        "   Order by ����,ʱ��,���"
        
    Set Get��ǰ�ƻ�ʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Check���ʱ��(strʱ��� As String, str��ʼʱ�� As String, str��ֹʱ�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ���ʱ���Ƿ��Ѿ��������ϰ�ʱ��
    '����:����
    '����:True ����;False δ����
    '����:2012-08-1 10:32:02
    '�����:55555
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str�ϰ� As String
    Dim str�°� As String
    Dim i As Long
    
    mrs�ϰ�ʱ���.Filter = "ʱ��� ='" & strʱ��� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 1 Then
        str�ϰ� = Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00")
        str�°� = Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")
    ElseIf mrs�ϰ�ʱ���.RecordCount = 2 Then
        While mrs�ϰ�ʱ���.EOF = False
            If i = 0 Then
                str�ϰ� = Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00")
            Else
                str�°� = Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")
            End If
            i = i + 1
            mrs�ϰ�ʱ���.MoveNext
        Wend
    End If
    If str�°� <> "" Then
        If Format("1900/1/1 " & str��ֹʱ��, "yyyy-MM-dd hh:mm:ss") > Format(str�°�, "yyyy-MM-dd hh:mm:ss") _
        Or Format("1900/1/1 " & str��ʼʱ��, "yyyy-MM-dd hh:mm:ss") < Format(str�ϰ�, "yyyy-MM-dd hh:mm:ss") Then
            Check���ʱ�� = True
            Exit Function
        Else
            Check���ʱ�� = False
            Exit Function
        End If
    End If
    Check���ʱ�� = False
End Function

