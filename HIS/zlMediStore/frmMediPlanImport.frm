VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmMediPlanImport 
   Caption         =   "ҩƷ�ƻ�����������"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   Icon            =   "frmMediPlanImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   12240
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6840
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   16
      Top             =   6120
      Width           =   3855
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   20
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "��ͣ��"
         Height          =   180
         Left            =   2640
         TabIndex        =   19
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.PictureBox picCaption 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   2640
      Width           =   4815
      Begin VB.ComboBox cboָ�� 
         Height          =   300
         Left            =   2730
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ָ������ⷿ"
         Height          =   180
         Left            =   1560
         TabIndex        =   22
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "�ƻ�����"
         Height          =   180
         Left            =   100
         TabIndex        =   11
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   840
      ScaleHeight     =   1695
      ScaleWidth      =   5655
      TabIndex        =   10
      Top             =   600
      Width           =   5655
      Begin VB.CommandButton cmdGetData 
         Caption         =   "��ȡ����(&G)"
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox chkZeroInput 
         Caption         =   "����0�ƻ�������ʾ"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdUnchoose 
         Caption         =   "ȫ��ѡ(&U)"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdChoose 
         Caption         =   "ȫѡ(&A)"
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   975
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _cx             =   2355
         _cy             =   1720
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
         SelectionMode   =   0
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
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   840
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   8
      Top             =   3000
      Width           =   7095
      Begin VB.PictureBox picOperation 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   6135
         TabIndex        =   9
         Top             =   1920
         Width           =   6135
         Begin VB.CheckBox chk������ͣ��ҩƷ 
            Caption         =   "������ͣ��ҩƷ"
            Height          =   180
            Left            =   360
            TabIndex        =   14
            Top             =   147
            Width           =   1815
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "ȡ��(&C)"
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   50
            Width           =   1095
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "����(&I)"
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   50
            Width           =   1095
         End
         Begin VB.Label lblInfo 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1215
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "ִ�������ͼƻ���������Ӵֱ�ʾִ�����������˼ƻ��������üƻ����Ѿ���������ⵥ�������"
         Top             =   120
         Width           =   2655
         _cx             =   4683
         _cy             =   2143
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
         SelectionMode   =   0
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
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   6435
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPlanImport.frx":014A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16510
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin XtremeDockingPane.DockingPane dkpView 
      Left            =   8280
      Top             =   2520
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPlanImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INT_WIDTH = 50
Private mrsCBO_Provider As ADODB.Recordset
Private mlngID As Long
Private mlngSum As Long '��¼ͣ��ҩƷ����
Private mstrMsg As String '��������ͣ��ҩƷ����ͣ��ҩƷʱ����ʾ��Ϣ

Private mstrPrivs As String
Private mlng�ⷿID As Long

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Public mintCostDigit As Integer        '�ɱ���С��λ��
Public mintPriceDigit As Integer       '�ۼ�С��λ��
Public mintNumberDigit As Integer      '����С��λ��
Public mintMoneyDigit As Integer       '���С��λ��
Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mblnLoad As Boolean

Private Sub cboָ��_Click()
    Dim i As Integer
    Dim blncheck�ⷿ As Boolean, bln��¼ As Boolean

    'ȡ��Ӧ�ⷿ�ľ���
    Call GetDrugDigit(Val(cboָ��.ItemData(cboָ��.ListIndex)), "", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
     'δ���ع������˳�
    If Not mblnLoad Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    With vsfDetail
        
        For i = 1 To .rows - 1
            '�ı䵼��ⷿ��id
            If .TextMatrix(i, .ColIndex("whid")) <> "" Then .TextMatrix(i, .ColIndex("whid")) = Val(cboָ��.ItemData(cboָ��.ListIndex))
            
            '��Ӧ�ⷿ�ľ���
            .ColFormat(.ColIndex("planqty")) = "#0." & String(mintNumberDigit, "0")
            .ColFormat(.ColIndex("execqty")) = "#0." & String(mintNumberDigit, "0")
            .ColFormat(.ColIndex("costprice")) = "#0." & String(mintCostDigit, "0")
            .ColFormat(.ColIndex("cost")) = "#0." & String(mintMoneyDigit, "0")
            .ColFormat(.ColIndex("sale")) = "#0." & String(mintMoneyDigit, "0")
            .ColFormat(.ColIndex("saleprice")) = "#0." & String(mintPriceDigit, "0")
            
            GetҩƷ�������� i
            
            '�ж�ҩƷ�Ƿ����ô洢�ⷿ�ڸõ���ⷿ��
            blncheck�ⷿ = Check�ⷿ(Val(.TextMatrix(i, .ColIndex("id"))), Val(cboָ��.ItemData(cboָ��.ListIndex)))
            If blncheck�ⷿ = False Then
                If Trim(.TextMatrix(i, .ColIndex("id"))) <> "" Then 'ֻ�ı���ҩƷ����
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HFFC0C0
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                    
                    bln��¼ = True
                End If
            Else
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbBack
                .TextMatrix(i, .ColIndex("choose")) = "-1"
                
                vsfDetail_AfterEdit i, .Row
            End If
            
            If bln��¼ Then
                lblInfo.Caption = "��ɫ�����Ǹ�ҩƷ�ڵ���ⷿ��δ���ô洢���ʣ�����������Ӧ�⹺���ݣ�"
                lblInfo.ForeColor = vbRed
                lblInfo.Visible = True
            Else
                lblInfo.Visible = False
            End If
            
            '�ж��Ƿ�ͣ�ã�ͣ����ʾδ
            If �Ƿ�ͣ��(Val(.TextMatrix(i, .ColIndex("id")))) Then
                .Cell(flexcpForeColor, i, .ColIndex("choose"), i, .Cols - 1) = &HFF00FF
            End If
            
        Next
    End With
    
    staThis.Panels(2).Text = CountBuilds
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkZeroInput_Click()
    Dim i As Integer
    
    With vsfMain
        vsfDetail.rows = 1

        For i = 1 To .rows - 1
            Call vsfMain_AfterEdit(i, .ColIndex("choose"))
        Next
    End With
End Sub

Private Sub chk������ͣ��ҩƷ_Click()
    If Not mblnLoad Then Exit Sub
    staThis.Panels(2).Text = CountBuilds
End Sub

Private Sub cmdCancel_Click()
    Unload frmMediPlanGetData
    Unload Me
End Sub

Private Sub cmdChoose_Click()
    If vsfMain.rows > 1 Then
        Dim i As Integer
        Dim blnCancel As Boolean
        Screen.MousePointer = vbHourglass
        vsfDetail.rows = 1
        For i = 1 To vsfMain.rows - 1
            vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "-1"
            vsfMain_BeforeEdit i, vsfMain.ColIndex("choose"), blnCancel
            If blnCancel = False Then
                vsfMain_AfterEdit i, vsfMain.ColIndex("choose")
            Else
                vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "0"
            End If
            DoEvents
        Next
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdGetData_Click()
'��ȡ�ƻ�������
    Dim strsql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset, rsTmp1 As ADODB.Recordset
    
    On Error GoTo errHandle
    frmMediPlanGetData.Show vbModal, Me
    If frmMediPlanGetData.SQLWhere = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    DoEvents
    strWhere = frmMediPlanGetData.SQLWhere
    
    
    strsql = "select a.ID,a.NO,a.�ڼ�,a.�ⷿID,b.���� �ⷿ,a.����˵��,a.������,a.��������,a.�����,a.������� " _
           & "from ҩƷ�ɹ��ƻ� a, ���ű� b " _
           & "where a.�ⷿid=b.id(+) and a.������� is not null "
               
    strsql = strsql & strWhere & " order by a.NO"
    
    vsfDetail.rows = 1
    vsfMain.rows = 1
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    If Not rsTmp.EOF Then
        'װ�ؼƻ�������
        DataLoading 1, rsTmp
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbDefault
        'MsgBox "δ��ȡ�����ݣ�", , gstrSysName
    End If
    
    vsfMain.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdImport_Click()
    If vsfDetail.rows <= 1 Then Exit Sub
    If MsgBox("��ȷ��Ҫ��������'ҩƷ�ƻ���'���ݣ�", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Dim strInsert As String, strNo As String
    Dim i As Integer, intSN As Integer
    Dim dateNO As Date
    Dim lngWHID As Long, lngPID As Long
    Dim rsTmp As New ADODB.Recordset
    Dim intժҪ���� As Integer
    Dim strժҪ As String
    Dim colժҪ As New Collection
    
    Screen.MousePointer = vbHourglass
    On Error GoTo errSoft
    
    intժҪ���� = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    
    '���ݼ��,������������¼�����
    With vsfDetail
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) = "-1" And .Cell(flexcpBackColor, i, 1, i, 2) <> &HFFC0C0 And Val(.TextMatrix(i, .ColIndex("id"))) <> 0 And Val(.TextMatrix(i, .ColIndex("batch"))) = 1 And Trim(.TextMatrix(i, .ColIndex("producer"))) = "" Then
                MsgBox "��" & i & "��ҩƷ�ڴ˿ⷿ�Ƿ������ʣ���Դ�ҩƷ�����ϴ������̣�"
                
                .SetFocus
                .Row = i
                .MsfObj.TopRow = i
                .Col = .ColIndex("producer")
                Exit Sub
            End If
        Next
    End With
    With rsTmp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        '.Fields.Append "���", adInteger, , adFldIsNullable
        .Fields.Append "�ⷿID", adBigInt, , adFldIsNullable
        .Fields.Append "��ҩ��λID", adBigInt, , adFldIsNullable
        .Fields.Append "ҩƷID", adBigInt, , adFldIsNullable
        .Fields.Append "������", adVarChar, 60, adFldIsNullable
        .Fields.Append "��������", adDBDate, , adFldIsNullable
        .Fields.Append "Ч��", adDBDate, , adFldIsNullable
        .Fields.Append "ʵ������", adDouble, , adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, , adFldIsNullable
        .Fields.Append "�ɱ����", adDouble, , adFldIsNullable
        .Fields.Append "���ۼ�", adDouble, , adFldIsNullable
        .Fields.Append "���۽��", adDouble, , adFldIsNullable
        .Fields.Append "���", adDouble, , adFldIsNullable
        .Fields.Append "�ӳ���", adDouble, , adFldIsNullable
        .Fields.Append "ҩ���װ", adDouble, , adFldIsNullable
        .Fields.Append "�ƻ�ID", adBigInt, , adFldIsNullable
        .Fields.Append "��׼�ĺ�", adVarChar, 40, adFldIsNullable
        .Fields.Append "ժҪ", adLongVarChar, intժҪ����, adFldIsNullable
        .Open
    End With
    
    With vsfDetail
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) = "-1" And .Cell(flexcpBackColor, i, 1, i, 2) <> &HFFC0C0 And �Ƿ���(i) Then
                rsTmp.AddNew
                rsTmp!�ⷿid = .TextMatrix(i, .ColIndex("whid"))
                rsTmp!��ҩ��λID = GetProviderID(.TextMatrix(i, .ColIndex("provider")))
                rsTmp!ҩƷid = .TextMatrix(i, .ColIndex("id"))
                rsTmp!������ = .TextMatrix(i, .ColIndex("producer"))
                rsTmp!�������� = IIf(.TextMatrix(i, .ColIndex("pdate")) = "", Null, .TextMatrix(i, .ColIndex("pdate")))
                If Not IsNull(rsTmp!��������) Then
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                        rsTmp!Ч�� = DateAdd("d", 1, DateAdd("m", .TextMatrix(i, .ColIndex("avail_day")), rsTmp!��������))
                    Else
                        rsTmp!Ч�� = DateAdd("m", .TextMatrix(i, .ColIndex("avail_day")), rsTmp!��������)
                    End If
                End If
                rsTmp!ʵ������ = .TextMatrix(i, .ColIndex("planqty"))
                rsTmp!�ɱ��� = .TextMatrix(i, .ColIndex("costprice"))
                rsTmp!�ɱ���� = .TextMatrix(i, .ColIndex("cost"))
                rsTmp!���ۼ� = .TextMatrix(i, .ColIndex("saleprice"))
                rsTmp!���۽�� = .TextMatrix(i, .ColIndex("sale"))
                rsTmp!��� = IIf(IsNull(rsTmp!���۽��), 0, rsTmp!���۽��) - IIf(IsNull(rsTmp!�ɱ����), 0, rsTmp!�ɱ����)
                rsTmp!�ӳ��� = Val(.TextMatrix(i, .ColIndex("add_rate")))
                rsTmp!ҩ���װ = .TextMatrix(i, .ColIndex("store_pak"))
                rsTmp!�ƻ�id = .TextMatrix(i, .ColIndex("planid"))
                rsTmp!��׼�ĺ� = .TextMatrix(i, .ColIndex("approval"))
                
                '�ϲ�ժҪ��ͬһ����Ӧ�̵�ժҪ�����ͬ����л��ܣ���;�ָ���
                If Trim(.TextMatrix(i, .ColIndex("Summary"))) <> "" Then
                    If ExistsColObject(colժҪ, "_" & Val(rsTmp!��ҩ��λID)) = False Then
                        '����û�ҵ�Ԫ����������Ԫ��
                        colժҪ.Add Trim(.TextMatrix(i, .ColIndex("Summary"))), "_" & Val(rsTmp!��ҩ��λID)
                    Else
                        '�����ҵ�Ԫ�أ�����ԭ��ֵ�Ļ����Ͻ��л���
                        strժҪ = colժҪ("_" & Val(rsTmp!��ҩ��λID))
                        If strժҪ = "" Then
                            strժҪ = Trim(.TextMatrix(i, .ColIndex("Summary")))
                        ElseIf InStr(1, ";" & strժҪ & ";", ";" & Trim(.TextMatrix(i, .ColIndex("Summary"))) & ";") = 0 Then
                            If LenB(StrConv(strժҪ & ";" & Trim(.TextMatrix(i, .ColIndex("Summary"))), vbFromUnicode)) <= intժҪ���� Then
                                strժҪ = strժҪ & ";" & Trim(.TextMatrix(i, .ColIndex("Summary")))
                            End If
                        End If
            
                        colժҪ.Remove "_" & Val(rsTmp!��ҩ��λID)
                        colժҪ.Add strժҪ, "_" & Val(rsTmp!��ҩ��λID)
                    End If
                End If
                
                rsTmp.Update
            End If
        Next
        
        '�ϲ�ժҪ
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If ExistsColObject(colժҪ, "_" & Val(rsTmp!��ҩ��λID)) = True Then
                strժҪ = colժҪ("_" & Val(rsTmp!��ҩ��λID))
                rsTmp!ժҪ = strժҪ
            Else
                rsTmp!ժҪ = ""
            End If
            
            rsTmp.Update
            rsTmp.MoveNext
        Loop
        
        rsTmp.MoveFirst
        
        rsTmp.Sort = "�ⷿID,��ҩ��λID"
    End With
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
            
    With rsTmp
        dateNO = Sys.Currentdate
        .MoveFirst
        Do While Not .EOF
            If lngWHID <> rsTmp!�ⷿid Or lngPID <> rsTmp!��ҩ��λID Then
                lngWHID = rsTmp!�ⷿid
                lngPID = rsTmp!��ҩ��λID
                intSN = 1
                strNo = Sys.GetNextNo(21, rsTmp!�ⷿid)
            End If
            'ִ�д洢���̣��ύ����
            strInsert = "zl_ҩƷ�⹺_INSERT("
            'NO
            strInsert = strInsert & "'" & strNo & "'"
            '���
            strInsert = strInsert & "," & intSN
            '�ⷿID
            strInsert = strInsert & "," & rsTmp!�ⷿid
            '�Է�����ID
            strInsert = strInsert & ",null"
            strInsert = strInsert & "," & rsTmp!��ҩ��λID
            strInsert = strInsert & "," & rsTmp!ҩƷid
            strInsert = strInsert & ",'" & rsTmp!������ & "'"
            '����
            strInsert = strInsert & ",'1'"
            'Ч��
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!Ч��), "null", "to_date('" & Format(rsTmp!Ч��, "yyyy-mm-dd") & "', 'yyyy-mm-dd')")
            'ʵ������
            strInsert = strInsert & "," & Round(rsTmp!ʵ������ * rsTmp!ҩ���װ, mintNumberDigit)
            '�ɱ���
            strInsert = strInsert & "," & Round(rsTmp!�ɱ��� / rsTmp!ҩ���װ, mintCostDigit)
            '�ɱ����
            strInsert = strInsert & "," & Round(rsTmp!�ɱ����, mintMoneyDigit)
            '����
            strInsert = strInsert & ",100"
            '���ۼ�
            strInsert = strInsert & "," & Round(rsTmp!���ۼ� / rsTmp!ҩ���װ, mintPriceDigit)
            strInsert = strInsert & "," & Round(rsTmp!���۽��, mintMoneyDigit)
            strInsert = strInsert & "," & Round(rsTmp!���, mintMoneyDigit)
            'ժҪ
            strInsert = strInsert & ",'" & rsTmp!ժҪ & "'"
            '������
            strInsert = strInsert & ",'" & UserInfo.�û����� & "'"
            '��Ʊ��
            strInsert = strInsert & ",null"
            '��Ʊ����
            strInsert = strInsert & ",null"
            '��Ʊ���
            strInsert = strInsert & ",Null"
            '��������
            strInsert = strInsert & ",to_date('" & dateNO & "','yyyy-mm-dd HH24:MI:SS')"
            '���
            strInsert = strInsert & ",Null"
            '��Ʒ�ϸ�֤
            strInsert = strInsert & ",Null"
            '�˲���
            strInsert = strInsert & ",Null"
            '�˲�����
            strInsert = strInsert & ",Null"
            '����
            strInsert = strInsert & ",Null"
            '�Ƿ��˻�
            strInsert = strInsert & ",1"
            '��������
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!��������), "null", "to_date('" & Format(rsTmp!��������, "yyyy-mm-dd") & "', 'yyyy-mm-dd')")
            '��׼�ĺ�
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!��׼�ĺ�), "Null", "'" & rsTmp!��׼�ĺ� & "'")
            '�������
            strInsert = strInsert & ",Null"
            '����
            strInsert = strInsert & ",Null"
            '�ӳ���
            strInsert = strInsert & "," & IIf(IsNull(rsTmp!�ӳ���), "null", rsTmp!�ӳ���)
            '��Ʊ����
            strInsert = strInsert & ",Null"
            '�ƻ�id
            strInsert = strInsert & "," & rsTmp!�ƻ�id
            strInsert = strInsert & ")"
            
            zlDatabase.ExecuteProcedure strInsert, Me.Caption
            
            intSN = intSN + 1
            lngWHID = rsTmp!�ⷿid
            lngPID = rsTmp!��ҩ��λID
            .MoveNext
        Loop
        gcnOracle.CommitTrans
    End With
    
    '��ʾ��Ϣ
    If mlngSum > 0 Then
        MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��ҩƷ��ͣ�ã��ⲿ��ҩƷ���������⹺��ⵥ�У�", "��" & mlngSum & "��ҩƷ��ͣ�ã��ⲿ��ҩƷ���������⹺��ⵥ�У�"), vbInformation, gstrSysName
        
        mlngSum = 0
        mstrMsg = ""
    End If
    
    '����ɹ����������
    vsfMain.rows = 1
    vsfDetail.rows = 1
    rsTmp.Close
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    Exit Sub
    
errSoft:
    Screen.MousePointer = vbDefault
    Call SaveErrLog
    Exit Sub

errHandle:
    gcnOracle.RollbackTrans
    'If ErrCenter() = 1 Then Resume
    Screen.MousePointer = vbDefault
    rsTmp.Close
    Call SaveErrLog
End Sub

'���ܣ��ж�ҩƷ�Ƿ�ͣ�ã��ٸ��ݸ�ѡ��������ͣ��ҩƷ������ֵ
'����ѡʱ��������ͣ��ҩƷ���������ж�ҩƷ�Ƿ�ͣ��ֱ�ӷ���TRUE
'������ѡʱ����������ͣ��ҩƷ�����ж�ҩƷ�Ƿ�ͣ�ã�ͣ�÷���false
Private Function �Ƿ���(Row As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "0" Then Exit Function
    
    If chk������ͣ��ҩƷ.Value = 1 Then  '������ͣ��ҩƷ
        �Ƿ��� = True
        Exit Function
    Else '��������ͣ��ҩƷ
    
        '�ж�ҩƷ�Ƿ�ͣ��
        gstrSQL = "select ����,��� from �շ���ĿĿ¼ where ID = [1] and nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�Ƿ�ͣ��", Val(vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("id"))))
        
        If rsTemp.RecordCount = 0 Then 'rsTemp.RecordCount = 0˵����ҩƷδͣ��
            �Ƿ��� = True
        Else
            �Ƿ��� = False
            
            mlngSum = mlngSum + 1
            If mlngSum <= 3 Then 'ƴ��ʾ��Ϣ��
                mstrMsg = mstrMsg & "��" & rsTemp!���� & "(" & rsTemp!��� & ")��" & Chr(10)
            End If
            
        End If
    End If

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdUnchoose_Click()
    If vsfMain.rows > 1 Then
        Dim i As Integer
        vsfDetail.rows = 1
        For i = 1 To vsfMain.rows - 1
            vsfMain.TextMatrix(i, vsfMain.ColIndex("choose")) = "0"
        Next
        staThis.Panels(2).Text = ""
        lblInfo.Visible = False
    End If
End Sub

Private Sub dkpView_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = picMain.hWnd
    End Select
End Sub

Private Sub dkpView_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Left = INT_WIDTH
    Right = INT_WIDTH
    Bottom = staThis.Height
End Sub

Private Sub dkpView_Resize()
    On Error Resume Next
    
    Dim lngL As Long, lngT As Long, lngR As Long, lngB As Long

    dkpView.GetClientRect lngL, lngT, lngR, lngB
    Me.picCaption.Move lngL, lngT, lngR - lngL
    Me.picDetail.Move lngL, lngT + picCaption.Height, lngR - lngL, lngB - lngT - picCaption.Height

'    With Me.picDetail
'        .Left = Me.sstMain.Left + INT_WIDTH
'        .Top = Me.sstMain.TabHeight + INT_WIDTH
'        .Width = Me.sstMain.Width - INT_WIDTH * 3
'        .Height = Me.sstMain.Height - Me.sstMain.TabHeight - INT_WIDTH * 2
'    End With

    With Me.vsfMain
        .Left = 0: .Top = 0
        .Width = Me.picMain.Width
        .Height = Me.picMain.Height - Me.picOperation.Height
    End With
    chkZeroInput.Top = vsfMain.Height + INT_WIDTH * 2
    chkZeroInput.Left = 0
    Me.cmdUnchoose.Top = vsfMain.Height + INT_WIDTH * 2
    Me.cmdChoose.Top = Me.cmdUnchoose.Top
    Me.cmdGetData.Top = Me.cmdChoose.Top
    
    Me.cmdUnchoose.Left = Me.picMain.Width - Me.cmdUnchoose.Width - INT_WIDTH * 2
    Me.cmdChoose.Left = Me.cmdUnchoose.Left - Me.cmdChoose.Width - INT_WIDTH
    Me.cmdGetData.Left = Me.cmdChoose.Left - Me.cmdGetData.Width - INT_WIDTH
    With Me.vsfDetail
        .Left = 0: .Top = 0
        .Width = Me.picDetail.Width
        .Height = Me.picDetail.Height - Me.picOperation.Height
    End With
    
    With Me.picOperation
        .Left = 0: .Top = Me.vsfDetail.Height
        .Width = Me.vsfDetail.Width
    End With
    Me.cmdCancel.Left = Me.picOperation.Width - Me.cmdCancel.Width - INT_WIDTH * 2
    Me.cmdImport.Left = Me.cmdCancel.Left - Me.cmdImport.Width - INT_WIDTH
    
    Me.chk������ͣ��ҩƷ.Left = Me.cmdImport.Left - Me.chk������ͣ��ҩƷ.Width - INT_WIDTH
    lblInfo.Width = cmdGetData.Left - 50
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    
    mblnLoad = False
    
    mstrPrivs = gstrprivs
    
    SetWarehouse

    chk������ͣ��ҩƷ.Value = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\zl9MediStore", "������ͣ��ҩƷ", 0)
    
    staThis.Panels(2).Picture = picColor
    
    '��ʼ��
    InitVSF 1
    vsfMain.AllowSelection = False
    picMain.BackColor = &H8000000F
    InitVSF 2
    vsfDetail.ExplorerBar = flexExNone

    Call InitTabs
    
    mblnLoad = True
End Sub

Private Sub InitVSF(ByVal bytIndex As Byte)
'��ʼ��VsfView
    Dim objVSF As VSFlex8Ctl.VSFlexGrid
    Dim strCols As String
    Dim arrCols As Variant
    Dim i As Integer
    
    If bytIndex = 1 Then
        Set objVSF = vsfMain
        strCols = "||ѡ��,choose,440|H_�ƻ�ID,planid,1000|�ƻ�����,no,880|H_�ⷿID,whid,1000|�ⷿ,wh,1000|�ڼ�,length,660|�������,verifydate,1900" _
                & "|�����,verifyer,660|��������,builddate,1900|������,builder,660|����˵��,explain,3000"
    Else
        Set objVSF = vsfDetail
        strCols = "||ѡ��,choose,440|H_�ƻ�id,planid,1000|H_�ⷿid,whid,0|�ƻ�����,planno,880|���,sn,440|H_ҩƷID,id,1000|ҩƷ����,name,1800|�ƻ�����,planqty,1500|ִ������,execqty,1500" _
                & "|ҩ�ⵥλ,unit,850|�ɱ���,costprice,1500|�ɱ����,cost,1500|�ۼ�,saleprice,1500|�ۼ۽��,sale,1500|��Ӧ��,provider,2000" _
                & "|�ϴ�������,producer,2000|H_��������,pdate,0|H_���Ч��,avail_day,0|H_�ӳ���,add_rate,0|H_ҩ���װ,store_pak,0|H_��������,batch,0|��׼�ĺ�,approval,2000|ժҪ,summary,0"
    End If
    
    With objVSF
        .rows = 1
        .ColWidth(0) = 130 * 2                               '��һ�п�
        .ColWidth(1) = 130
        .FixedCols = 2                                       '�̶�ǰ����
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns               '����ʱ�ɵ���Columns���
        .AllowSelection = True                               '�൥Ԫѡ����ƿ���
        .SelectionMode = flexSelectionListBox                '�൥Ԫѡ�����
        .ExplorerBar = flexExSortShow
        '.BackColorSel = &HC0E0FF
        .BackColorBkg = vbWhite
    End With
    
    arrCols = Split(strCols, "|")
    With objVSF
        .Cols = UBound(arrCols) + 1
        For i = LBound(arrCols) To UBound(arrCols)
            If arrCols(i) = "" Then
                .TextMatrix(0, i) = ""
            Else
                .TextMatrix(0, i) = Split(arrCols(i), ",")(0)
                .ColKey(i) = Split(arrCols(i), ",")(1)
                .ColWidth(i) = Split(arrCols(i), ",")(2)
                'H_Ϊ������
                If Mid(Split(arrCols(i), ",")(0), 1, 2) = "H_" Then
                    .colHidden(i) = True
                Else
                    .colHidden(i) = False
                End If
            End If
        Next
        If .ColIndex("choose") > 0 Then
            .ColDataType(.ColIndex("choose")) = flexDTBoolean    '����ΪCheck�ؼ�
        End If
    End With
    If bytIndex = 2 Then
        vsfDetail.ColComboList(vsfDetail.ColIndex("provider")) = GetComboVSF("select ���� from ��Ӧ�� order by ����")
        vsfDetail.ColComboList(vsfDetail.ColIndex("producer")) = GetComboVSF("Select ���� From ҩƷ������")
    End If
End Sub

Private Sub InitTabs()
'��ʼ��Tabs
    Dim objPane1 As Pane

    Set objPane1 = dkpView.CreatePane(1, 0, Me.ScaleY(Me.Height * 0.5, vbTwips, vbPoints), DockTopOf)
    objPane1.Title = "�ƻ���"
    objPane1.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable

    With dkpView
        .Options.ThemedFloatingFrames = True
        .Options.LunaColors = False
        .Options.AlphaDockingContext = True
        .VisualTheme = ThemeOffice2003
        '.Options.FloatingFrameCaption = "Panes"
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    Dim blnData As Boolean
    If vsfMain.rows <= 1 Then Exit Sub
'    If mrsDetail.State = adStateClosed Then Exit Sub
'    If mrsDetail.RecordCount = 0 Then Exit Sub
'�й�ѡ��������ʾ
    For i = 1 To vsfDetail.rows - 1
        If vsfDetail.TextMatrix(i, vsfDetail.ColIndex("choose")) = "-1" Then
            blnData = True
            Exit For
        End If
    Next
    If blnData Then
        If MsgBox("��������δ����ȷ��Ҫȡ����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7050 Then Me.Width = 7050
    If Me.Height < 6500 Then Me.Height = 6500
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
End Sub

Public Sub DataLoading(ByVal bytIndex As Byte, ByRef rsVal As ADODB.Recordset)
    Dim i As Integer, j As Integer
    Dim strName As String, strSpec As String, strUnit As String, strProvider As String
    Dim blnGet As Boolean
    Dim vsfTmp As VSFlex8Ctl.VSFlexGrid
    Dim blncheck�ⷿ As Boolean
'    Dim intCostDigit As Integer, intPriceDigit As Integer, intNumberDigit As Integer, intMoneydigit As Integer
'    'ҩ�⾫��
'    intCostDigit = frmMediPlanGetData.mintCostDigit
'    intPriceDigit = frmMediPlanGetData.mintPriceDigit
'    intNumberDigit = frmMediPlanGetData.mintNumberDigit
'    intMoneydigit = frmMediPlanGetData.mintMoneyDigit

    If bytIndex = 1 Then
        Set vsfTmp = vsfMain
    Else
        Set vsfTmp = vsfDetail
        If chkZeroInput.Value = False Then
            rsVal.Filter = "�ƻ�����<>0"
        End If
    End If
    
    With vsfTmp
        If rsVal.RecordCount = 0 Then Exit Sub
        If bytIndex = 1 Then
            j = 1
            .rows = rsVal.RecordCount + 1
        Else
            j = .rows
            .rows = .rows + rsVal.RecordCount
        End If
        rsVal.MoveFirst
        For i = j To j + rsVal.RecordCount - 1
            'strName = "": strSpec = "": strUnit = ""
            'blnGet = GetMedicalInfo(rsVal!ҩƷid, strName, strSpec, strUnit)
            '���
            .TextMatrix(i, 1) = i
            '�ƻ���
            If bytIndex = 1 Then
                .TextMatrix(i, .ColIndex("choose")) = 0
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!id), "", rsVal!id)
                .TextMatrix(i, .ColIndex("no")) = IIf(IsNull(rsVal!NO), "", rsVal!NO)
                .TextMatrix(i, .ColIndex("whid")) = IIf(IsNull(rsVal!�ⷿid), 0, rsVal!�ⷿid)
                .TextMatrix(i, .ColIndex("wh")) = IIf(IsNull(rsVal!�ⷿ), "ȫԺ", rsVal!�ⷿ)
                .TextMatrix(i, .ColIndex("length")) = IIf(IsNull(rsVal!�ڼ�), "", rsVal!�ڼ�)
                .TextMatrix(i, .ColIndex("verifydate")) = IIf(IsNull(rsVal!�������), "", rsVal!�������)
                .ColFormat(.ColIndex("verifydate")) = "yyyy-mm-dd hh:mm:ss"
                .TextMatrix(i, .ColIndex("verifyer")) = IIf(IsNull(rsVal!�����), "", rsVal!�����)
                .TextMatrix(i, .ColIndex("builddate")) = IIf(IsNull(rsVal!��������), "", rsVal!��������)
                .ColFormat(.ColIndex("builddate")) = "yyyy-mm-dd hh:mm:ss"
                .TextMatrix(i, .ColIndex("builder")) = IIf(IsNull(rsVal!������), "", rsVal!������)
                .TextMatrix(i, .ColIndex("explain")) = IIf(IsNull(rsVal!����˵��), "", rsVal!����˵��)
            '�ƻ�����
            Else
                .TextMatrix(i, .ColIndex("approval")) = IIf(IsNull(rsVal!��׼�ĺ�), "", rsVal!��׼�ĺ�)
                .TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(rsVal!����˵��), "", rsVal!����˵��)
                .TextMatrix(i, .ColIndex("planid")) = IIf(IsNull(rsVal!�ƻ�id), "", rsVal!�ƻ�id)
                If IIf(IsNull(rsVal!�ϴι�Ӧ��), "", rsVal!�ϴι�Ӧ��) = "" Then
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                Else
                    .TextMatrix(i, .ColIndex("choose")) = IIf(IsNull(rsVal!ѡ��), "", rsVal!ѡ��)
                End If
'                If mblnȫԺ Then
'                    .TextMatrix(i, .ColIndex("whid")) = mlngID
'                Else
'                    .TextMatrix(i, .ColIndex("whid")) = IIf(IsNull(rsVal!�ⷿid), "", rsVal!�ⷿid)
'                End If
                
                .TextMatrix(i, .ColIndex("whid")) = Val(cboָ��.ItemData(cboָ��.ListIndex))
                
                .TextMatrix(i, .ColIndex("planno")) = IIf(IsNull(rsVal!NO), "", rsVal!NO)
                .TextMatrix(i, .ColIndex("sn")) = IIf(IsNull(rsVal!���), "", rsVal!���)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rsVal!ҩƷid), "", rsVal!ҩƷid)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rsVal!����), "", rsVal!����)
                .TextMatrix(i, .ColIndex("planqty")) = IIf(IsNull(rsVal!�ƻ�����), "", rsVal!�ƻ�����)
                .TextMatrix(i, .ColIndex("execqty")) = IIf(IsNull(rsVal!ִ������), "", rsVal!ִ������)
                .ColFormat(.ColIndex("planqty")) = "#0." & String(mintNumberDigit, "0")
                .ColFormat(.ColIndex("execqty")) = "#0." & String(mintNumberDigit, "0")
                .TextMatrix(i, .ColIndex("unit")) = IIf(IsNull(rsVal!ҩ�ⵥλ), "", rsVal!ҩ�ⵥλ)
                .TextMatrix(i, .ColIndex("costprice")) = IIf(IsNull(rsVal!�ɱ���), "", rsVal!�ɱ���) * IIf(IsNull(rsVal!ҩ���װ), 0, rsVal!ҩ���װ)
                .ColFormat(.ColIndex("costprice")) = "#0." & String(mintCostDigit, "0")
                
                .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(rsVal!�ƻ�����), 0, rsVal!�ƻ�����) * .TextMatrix(i, .ColIndex("costprice"))
                .ColFormat(.ColIndex("cost")) = "#0." & String(mintMoneyDigit, "0")
                '�����ۼ�
                Dim dblTmp As Double
                dblTmp = 1 + IIf(IsNull(rsVal!�ӳ���), 0, rsVal!�ӳ���)
                If IIf(IsNull(rsVal!�Ƿ���), 0, rsVal!�Ƿ���) = 1 Then
                    '���
                    dblTmp = dblTmp * IIf(IsNull(rsVal!�ɱ���), 0, rsVal!�ɱ���)
                    If dblTmp >= IIf(IsNull(rsVal!ָ�����ۼ�), 0, rsVal!ָ�����ۼ�) Then
                        dblTmp = IIf(IsNull(rsVal!ָ�����ۼ�), 0, rsVal!ָ�����ۼ�)
                    Else
                        dblTmp = dblTmp _
                               + (IIf(IsNull(rsVal!ָ�����ۼ�), 0, rsVal!ָ�����ۼ�) - dblTmp) _
                               * (1 - (IIf(IsNull(rsVal!���������), 0, rsVal!���������) / 100))
                    End If
                Else
                    '����
                    'dblTmp = IIf(IsNull(rsVal!�ּ�), 0, rsVal!�ּ�)
                    'If dblTmp >= IIf(IsNull(rsVal!ָ�����ۼ�), 0, rsVal!ָ�����ۼ�) Then
                        dblTmp = Get�ۼ�(False, Val(.TextMatrix(i, .ColIndex("id"))), Val(.TextMatrix(i, .ColIndex("whid"))), 0) 'IIf(IsNull(rsVal!�ۼ�), 0, rsVal!�ۼ�)
                    'End If
                End If
                .TextMatrix(i, .ColIndex("saleprice")) = dblTmp * IIf(IsNull(rsVal!ҩ���װ), 0, rsVal!ҩ���װ)
                .ColFormat(.ColIndex("saleprice")) = "#0." & String(mintPriceDigit, "0")
                
                .TextMatrix(i, .ColIndex("sale")) = IIf(IsNull(rsVal!�ƻ�����), 0, rsVal!�ƻ�����) * .TextMatrix(i, .ColIndex("saleprice"))
                .ColFormat(.ColIndex("sale")) = "#0." & String(mintMoneyDigit, "0")
                
                If IsNull(rsVal!��Ӧ��id) Or rsVal!��Ӧ��id = "" Then
                    .TextMatrix(i, .ColIndex("provider")) = ""
                    .TextMatrix(i, .ColIndex("choose")) = "0"
                Else
                    .TextMatrix(i, .ColIndex("provider")) = IIf(IsNull(rsVal!�ϴι�Ӧ��), "", rsVal!�ϴι�Ӧ��)
                End If
                .TextMatrix(i, .ColIndex("producer")) = IIf(IsNull(rsVal!�ϴ�������), "", rsVal!�ϴ�������)
                '���÷�������
                Call GetҩƷ��������(i)
                
                .TextMatrix(i, .ColIndex("pdate")) = IIf(IsNull(rsVal!�ϴ���������), "", rsVal!�ϴ���������)
                .TextMatrix(i, .ColIndex("avail_day")) = IIf(IsNull(rsVal!���Ч��), "", rsVal!���Ч��)
                .TextMatrix(i, .ColIndex("add_rate")) = IIf(IsNull(rsVal!�ӳ���), "", rsVal!�ӳ���)
                .TextMatrix(i, .ColIndex("store_pak")) = IIf(IsNull(rsVal!ҩ���װ), "", rsVal!ҩ���װ)
                
                blncheck�ⷿ = Check�ⷿ(Val(IIf(IsNull(rsVal!ҩƷid), 0, rsVal!ҩƷid)), Val(cboָ��.ItemData(cboָ��.ListIndex)))
                If blncheck�ⷿ = False Then
                    If Trim(.TextMatrix(i, .ColIndex("id"))) <> "" Then 'ֻ�ı���ҩƷ����
                        .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HFFC0C0
                        .TextMatrix(i, .ColIndex("choose")) = "0"
                    End If
                End If
               
                'ִ������>�ƻ����� ���Ӧ����������Ӵ�
                If Val(.TextMatrix(i, .ColIndex("execqty"))) > Val(.TextMatrix(i, .ColIndex("planqty"))) Then
                    .Cell(flexcpFontBold, i, .ColIndex("planqty"), i, .ColIndex("execqty")) = True
                End If
                
                '�ж��Ƿ�ͣ�ã�ͣ����ʾδ
                If �Ƿ�ͣ��(Val(rsVal!ҩƷid)) Then
                    .Cell(flexcpForeColor, i, .ColIndex("choose"), i, .Cols - 1) = &HFF00FF
                End If

            End If
            rsVal.MoveNext
        Next
        
        'ȷ��vsf��ŵĿ��
        .ColWidth(1) = IIf(.rows > 0, Len(Trim(Str(.rows))) * 130 + 70, 200)
        
        'ΪvsfDetail�ϲ���
        If bytIndex = 2 Then
            With vsfDetail
                .rows = .rows + 1
                .Row = .rows - 1
                .TextMatrix(.Row, 1) = .Row
                .TextMatrix(.Row, .ColIndex("planno")) = .TextMatrix(.Row - 1, .ColIndex("planno"))
                .MergeCells = flexMergeFree
                '.ColDataType(.ColIndex("choose")) = flexDTSingle
                .MergeRow(.rows - 1) = True
                .Cell(flexcpText, .Row, .ColIndex("planno") + 1, .Row, .Cols - 1) = " "
                .Cell(flexcpForeColor, .Row, .ColIndex("choose"), .Row, .Cols - 1) = &H80000010
            End With
        End If
    End With
End Sub

Private Sub GetҩƷ��������(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strsql As String
    Dim int�������� As Integer      '0-������;1-����
    Dim intҩ����� As Integer      '0-������;1-����
    Dim intҩ������ As Integer      '0-������;1-����
    Dim bln�Ƿ����ҩ������ As Boolean  'True-����ҩ������;False-������ҩ������
        
    On Error GoTo errHandle
    With vsfDetail
        If Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("id"))) = 0 Then Exit Sub
        
        strsql = "SELECT NVL(ҩ�����, 0) ҩ�����,NVL(ҩ������, 0) ҩ������ " & _
                " From ҩƷ��� WHERE ҩƷID = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "ȡҩƷ�ⷿ��������", Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("id"))))
        
        If rsTemp.RecordCount > 0 Then
            intҩ����� = rsTemp!ҩ�����
            intҩ������ = rsTemp!ҩ������
        End If
        
        If intҩ������ = 1 Then     '���ҩ�����������������Ϊ1
            int�������� = 1
        Else
            If intҩ����� = 1 Then
                strsql = "SELECT ����ID From ��������˵�� " & _
                        " WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) AND ����ID = [1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "ȡ��������", Val(vsfDetail.TextMatrix(intBillRow, .ColIndex("whid"))))
                
                bln�Ƿ����ҩ������ = (rsTemp.RecordCount > 0)
                        
                If bln�Ƿ����ҩ������ Then
                    int�������� = 0
                Else
                    int�������� = 1
                End If
            End If
        End If
        
        vsfDetail.TextMatrix(intBillRow, .ColIndex("batch")) = int��������
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check�ⷿ(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long) As Boolean
    Dim rsTemp As Recordset
    
    gstrSQL = "select �շ�ϸĿid from �շ�ִ�п��� where �շ�ϸĿid=[1] and ִ�п���id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���洢�ⷿ", lngҩƷID, lng�ⷿID)
    If rsTemp.RecordCount > 0 Then
        Check�ⷿ = True
    Else
        Check�ⷿ = False
    End If
End Function

Private Sub txtPlanNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '����ע�����Ϣ(�Ƿ���ʾͣ��ҩƷ)
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\zl9MediStore", "������ͣ��ҩƷ", chk������ͣ��ҩƷ.Value
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Trim(vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("provider"))) = "" Then
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "0"
    ElseIf Col = vsfDetail.ColIndex("provider") Then
        vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "-1"
    End If
    staThis.Panels(2).Text = CountBuilds
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDetail
        '�ⷿid�ޣ������޸�
        If .TextMatrix(Row, .ColIndex("whid")) = "" Then
            Cancel = True
            Exit Sub
        End If
        
        'ѡ������޸�
        If Col = .ColIndex("choose") Then
            If Trim(.TextMatrix(Row, .ColIndex("provider"))) = "" Or .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HFFC0C0 Then
                Cancel = True
            Else
                Cancel = False
            End If
        ElseIf Col = .ColIndex("provider") Then
            Cancel = False
        ElseIf Col = .ColIndex("producer") Then
            Cancel = False
        ElseIf Col = .ColIndex("approval") Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfDetail_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
'    If Col = vsfDetail.ColIndex("provider") Then
'        If vsfDetail.EditText <> "" And vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) <> "-1" Then
'            vsfDetail.TextMatrix(Row, vsfDetail.ColIndex("choose")) = "-1"
'        End If
'    End If
End Sub
Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = vsfDetail.ColIndex("approval") Then
        If KeyAscii <> vbKeyBack Then
            If LenB(StrConv(vsfDetail.EditText, vbFromUnicode)) >= 40 Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub vsfDetail_GotFocus()
    picCaption.BackColor = &HFFFFE9
End Sub

Private Sub vsfDetail_LostFocus()
    picCaption.BackColor = &H8000000A
End Sub

Private Sub vsfMain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strsql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim bln��¼ As Boolean
    
    On Error GoTo errHandle
    If Col = vsfMain.ColIndex("choose") Then
        If vsfMain.TextMatrix(Row, Col) = -1 Then
            'װ��ҩƷ�ƻ�����
            strWhere = " and id=" & vsfMain.TextMatrix(Row, vsfMain.ColIndex("planid"))
            strsql = "select -1 ѡ��,b.�ⷿid, B.NO, A.�ƻ�id, A.ҩƷid," _
                   & "  D.����, A.���, A.�ƻ����� / C.ҩ���װ as �ƻ�����,nvl(A.ִ������,0) / C.ҩ���װ as ִ������, C.ҩ�ⵥλ, C.ҩ���װ," _
                   & "  Nvl(case When Nvl(A.����, 0) = 0 then " _
                   & "        (Select �ϴβɹ��� From ҩƷ��� Where Nvl(����, 0) =" _
                   & "           (Select nvl(Max(����),0) From ҩƷ��� Where B.�ⷿid = �ⷿid And A.ҩƷid = ҩƷid And ���� = 1) " _
                   & " and b.�ⷿid = �ⷿid And a.ҩƷid = ҩƷid And ���� = 1 and rownum=1 ) " _
                   & "      else  A.���� End, C.�ɱ���) �ɱ���,a.�ۼ�, " _
                   & "A.�ϴι�Ӧ��, A.�ϴ�������, A.˵��, F.�Ƿ���, " _
                   & "c.�ӳ���/100 as �ӳ���," _
                   & "G.id ��Ӧ��ID, (select max(�ϴ���������) from ҩƷ��� where ҩƷid=c.ҩƷid) �ϴ���������, F.���Ч��, C.ָ�����ۼ�, C.���������, Nvl(a.��׼�ĺ�, Nvl(c.�ϴ���׼�ĺ�, c.��׼�ĺ�)) As ��׼�ĺ�, b.����˵�� " _
                   & "From ҩƷ�ƻ����� A," _
                   & "     (Select ID, NO, �ⷿid,����˵�� From ҩƷ�ɹ��ƻ� Where ������� is not null " & strWhere _
                   & ") B, ҩƷ��� C, �շ���ĿĿ¼ D, ҩƷĿ¼ F,��Ӧ�� G " _
                   & "Where A.�ƻ�id = B.ID And A.ҩƷid = C.ҩƷid And C.ҩƷid = D.ID And C.ҩƷid = F.ҩƷid " _
                   & "   and A.�ϴι�Ӧ��=G.����(+) " _
                   & "order by a.�ƻ�id,a.��� "
            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
            'װ������
            DataLoading 2, rsTmp
        Else
            '���ҩƷ�ƻ�����
            strWhere = vsfMain.TextMatrix(Row, vsfMain.ColIndex("no"))
            For i = vsfDetail.rows - 1 To 1 Step -1
                If strWhere = vsfDetail.TextMatrix(i, vsfDetail.ColIndex("planno")) Then
                    vsfDetail.RemoveItem i
                End If
            Next
            
            'ˢ��VSF�����
            For i = 1 To vsfDetail.rows - 1
                vsfDetail.TextMatrix(i, 1) = i
            Next
            'ȷ��vsf��ŵĿ��
            vsfDetail.ColWidth(1) = IIf(vsfDetail.rows > 0, Len(Trim(Str(vsfDetail.rows))) * 130 + 70, 200)
        End If
    End If
    
    With vsfDetail
        For i = 1 To .rows - 1
            If .Cell(flexcpBackColor, i, 1, i, 2) = &HFFC0C0 Then
                bln��¼ = True
                Exit For
            End If
        Next
        If bln��¼ = True Then
            lblInfo.Caption = "��ɫ�����Ǹ�ҩƷ�ڵ���ⷿ��δ���ô洢���ʣ�����������Ӧ�⹺���ݣ�"
            lblInfo.ForeColor = vbRed
            lblInfo.Visible = True
        Else
            lblInfo.Visible = False
        End If
    End With
    
    staThis.Panels(2).Text = CountBuilds
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_RowColChange()
    '��ǰ��¼�ü�ͷָʾ
    vsfDetail.Cell(flexcpText, 0, 0, vsfDetail.rows - 1, 0) = ""
    If vsfDetail.Row > 0 Then
        vsfDetail.Cell(flexcpFontName, , 0) = "Marlett"
        vsfDetail.TextMatrix(vsfDetail.Row, 0) = 4
    End If
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfMain.TextMatrix(Row, vsfMain.ColIndex("whid")) = "" Then
        Cancel = True
        Exit Sub
    End If
        
    If Col = vsfMain.ColIndex("choose") Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub vsfMain_RowColChange()
    '��ǰ��¼�ü�ͷָʾ
    vsfMain.Cell(flexcpText, 0, 0, vsfMain.rows - 1, 0) = ""
    If vsfMain.Row > 0 Then
        vsfMain.Cell(flexcpFontName, , 0) = "Marlett"
        vsfMain.TextMatrix(vsfMain.Row, 0) = 4
    End If
End Sub

Private Function GetComboVSF(ByVal strsql As String)
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    '��ʽ: "#1;Full time|#23;Part time|#65;Contractor|#78;Intern|#0;Other"
    strTmp = " |"
    Do While Not rsTmp.EOF
        strTmp = strTmp & rsTmp.Fields(0) & "|"
        rsTmp.MoveNext
    Loop
    GetComboVSF = strTmp
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CountBuilds()
    Dim i As Integer, j As Integer
'    Dim strOldProvider As String, intOldWHID As Integer
    Dim blnFind As Boolean
    Dim rsProvider As New ADODB.Recordset
    rsProvider.Fields.Append "provider", adVarChar, 1000, adFldIsNullable
    rsProvider.Fields.Append "whid", adInteger, 18, adFldIsNullable
    rsProvider.Open
    '�����������ⵥ�ݣ������ٸ���Ӧ��+�ⷿID
    With vsfDetail
'        If .Rows > 1 Then
'            strOldProvider = Trim(.TextMatrix(1, .ColIndex("provider")))
'            intOldWHID = .TextMatrix(1, .ColIndex("whid"))
'        End If
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("choose")) <> "-1" Then GoTo EndFor
            If Trim(.TextMatrix(i, .ColIndex("provider"))) = "" Then GoTo EndFor
            'If Trim(.TextMatrix(i, .ColIndex("provider"))) = strOldProvider Then GoTo EndFor
            blnFind = False
            If rsProvider.RecordCount > 0 Then rsProvider.MoveFirst
            Do While Not rsProvider.EOF
                If rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider"))) And rsProvider!whid = .TextMatrix(i, .ColIndex("whid")) Then
                    blnFind = True
                    Exit Do
                End If
                rsProvider.MoveNext
            Loop
            If blnFind = False Then
                If chk������ͣ��ҩƷ.Value = 0 Then '��������ͣ��
                    If Not �Ƿ�ͣ��(Val(.TextMatrix(i, .ColIndex("id")))) Then
                        rsProvider.AddNew
                        rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider")))
                        rsProvider!whid = .TextMatrix(i, .ColIndex("whid"))
                        rsProvider.Update
                    End If
                Else
                    rsProvider.AddNew
                    rsProvider!Provider = Trim(.TextMatrix(i, .ColIndex("provider")))
                    rsProvider!whid = .TextMatrix(i, .ColIndex("whid"))
                    rsProvider.Update
                End If
            End If
            
'        strOldProvider = Trim(.TextMatrix(i, .ColIndex("provider")))
'        intOldWHID = .TextMatrix(i, .ColIndex("whid"))
            
EndFor:
        Next
        
        CountBuilds = "��ѡ������ݣ������� " & rsProvider.RecordCount & " ����ⵥ�ݡ�"
        rsProvider.Close
    End With
End Function

Private Function GetProviderID(ByVal strProvider As String) As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    Set rsTmp = zlDatabase.OpenSQLRecord("select ID from ��Ӧ�� where rownum=1 and ����=[1]", Me.Caption, strProvider)
    If Not rsTmp.EOF Then GetProviderID = rsTmp!id
    rsTmp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'���ܣ��ж��Ƿ�ͣ��,true - ͣ��
Private Function �Ƿ�ͣ��(ByVal lngҩƷID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If lngҩƷID = 0 Then Exit Function

    
    '�ж�ҩƷ�Ƿ�ͣ��
    gstrSQL = "select ����,��� from �շ���ĿĿ¼ where ID = [1] and nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ҩƷ�Ƿ�ͣ��", lngҩƷID)
    
    �Ƿ�ͣ�� = rsTemp.RecordCount <> 0  '˵����ҩƷδͣ��

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub SetWarehouse()
'���ÿⷿ��ID
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    Dim i As Integer
    Dim cboTmp As ComboBox
    
    On Error GoTo ErrHand
    
    Set cboTmp = cboָ��
    
    If InStr(1, mstrPrivs, "����ҩ���⹺���") = 0 Then
        strsql = "Select Distinct A.ID �ⷿID, A.���� �ⷿ " _
               & "From ��������˵�� C, �������ʷ��� B, ���ű� A " _
               & "Where C.�������� = B.���� And B.���� In ('H', 'I', 'J') And A.ID = C.����id And To_Char(A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        strsql = "Select Distinct A.ID �ⷿID, A.���� �ⷿ " _
               & "From ��������˵�� C, �������ʷ��� B, ���ű� A " _
               & "Where C.�������� = B.���� And B.���� In ('H', 'I', 'J','K', 'L', 'M','N') And A.ID = C.����id And To_Char(A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    End If
    
    cboTmp.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 0 To rsTmp.RecordCount - 1
            cboTmp.AddItem rsTmp!�ⷿ
            cboTmp.ItemData(i) = rsTmp!�ⷿid
            
            If mlng�ⷿID = Val(rsTmp!�ⷿid) Then cboTmp.ListIndex = i
            
            rsTmp.MoveNext
        Next
        
    End If
    
    If cboTmp.ListIndex = -1 Then cboTmp.ListIndex = 0
    rsTmp.Close
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub ShowCard(FrmMain As Form, ByVal lng�ⷿID As Long)

    mlng�ⷿID = lng�ⷿID

    Me.Show vbModal, FrmMain
End Sub
