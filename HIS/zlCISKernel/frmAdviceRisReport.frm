VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAdviceRisReport 
   AutoRedraw      =   -1  'True
   Caption         =   "��ӡRISԤԼ��"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13260
   Icon            =   "frmAdviceRisReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13260
   StartUpPosition =   2  '��Ļ����
   Begin XtremeReportControl.ReportControl rptAdvice 
      Height          =   1170
      Left            =   2205
      TabIndex        =   21
      Top             =   2415
      Width           =   630
      _Version        =   589884
      _ExtentX        =   1111
      _ExtentY        =   2064
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.TabControl tbcAppend 
      Height          =   1530
      Left            =   3015
      TabIndex        =   24
      Top             =   5325
      Width           =   270
      _Version        =   589884
      _ExtentX        =   476
      _ExtentY        =   2699
      _StockProps     =   64
   End
   Begin VB.Frame fraAdviceUD 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   -405
      MousePointer    =   7  'Size N S
      TabIndex        =   22
      Top             =   4725
      Width           =   6000
   End
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   3855
      ScaleHeight     =   2850
      ScaleWidth      =   4890
      TabIndex        =   14
      Top             =   1980
      Visible         =   0   'False
      Width           =   4920
      Begin VB.CommandButton cmdFindCancle 
         Caption         =   "ȡ��"
         Height          =   270
         Left            =   4200
         TabIndex        =   7
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindOk 
         Caption         =   "ȷ��"
         Height          =   270
         Left            =   3480
         TabIndex        =   8
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����"
         Height          =   270
         Left            =   1740
         TabIndex        =   16
         Top             =   75
         Width           =   615
      End
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   50
         TabIndex        =   15
         Top             =   75
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2280
         Left            =   75
         TabIndex        =   9
         ToolTipText     =   "ȫѡCtrl+A��ȫ��Ctrl+R"
         Top             =   510
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fraFilter 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   12
      Top             =   375
      Width           =   12300
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(F3)"
         Height          =   300
         Left            =   4230
         TabIndex        =   6
         Top             =   765
         Width           =   900
      End
      Begin VB.TextBox txtFilter 
         Height          =   300
         Left            =   2205
         TabIndex        =   5
         Top             =   780
         Width           =   2000
      End
      Begin VB.ComboBox cboFind 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   1320
      End
      Begin VB.OptionButton optType 
         Caption         =   "�Ѵ�ӡ"
         Height          =   240
         Index           =   1
         Left            =   11130
         TabIndex        =   19
         Top             =   405
         Width           =   870
      End
      Begin VB.OptionButton optType 
         Caption         =   "δ��ӡ"
         Height          =   240
         Index           =   0
         Left            =   10260
         TabIndex        =   18
         Top             =   405
         Value           =   -1  'True
         Width           =   850
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "��"
         Height          =   265
         Left            =   4560
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl+D"
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtDept 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   870
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "���п���"
         ToolTipText     =   "���п���"
         Top             =   345
         Width           =   4000
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   5820
         TabIndex        =   2
         Top             =   360
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   179765251
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   8130
         TabIndex        =   3
         Top             =   360
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   179765251
         CurrentDate     =   37953
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "���Ҳ���"
         Height          =   180
         Left            =   90
         TabIndex        =   25
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblTim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԤԼʱ��                        ��"
         Height          =   180
         Left            =   5025
         TabIndex        =   20
         Top             =   405
         Width           =   3060
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   90
         TabIndex        =   11
         Top             =   375
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   90
         TabIndex        =   13
         Top             =   45
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7425
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   635
      SimpleText      =   $"frmAdviceRisReport.frx":014A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceRisReport.frx":0191
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18309
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
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   1395
      Left            =   4320
      TabIndex        =   23
      Top             =   5295
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   2461
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAdviceRisReport.frx":0A25
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
   Begin MSComctlLib.ImageList img16 
      Left            =   660
      Top             =   1875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":0AC2
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":105C
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":15F6
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":1B90
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":212A
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":26C4
            Key             =   "������"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":8F26
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceRisReport.frx":F788
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAdviceRisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng����ID As Long
Private mlngFind As Long
Private mstrMatch As String
Private mstrPrivs As String
Private mintPreDept As Integer
Private mstrFindType As String

Private Enum PatiCol
    COL_ѡ��
    COL_����
    COL_סԺ��
    COL_����
    COL_�Ա�
    COL_����
    COL_����
    COL_ԤԼʱ��
    COL_ԤԼ�豸
    COL_ԤԼ��
    COL_ԤԼ��Ŀ
    COL_��ӡʱ��
    COL_��ӡ��
    
    COL_����ID
    COL_��ҳID
    COL_ҽ��ID
End Enum

Private Enum Ectrl
    eδ��
    e�Ѵ�
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long) As Boolean
    mlng����ID = lng����ID
    Me.Show , frmParent
End Function

Private Sub cboFind_Click()
    mstrFindType = cboFind.Text
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngLW As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraFilter.Top = lngTop
    fraFilter.Left = lngLeft
    fraFilter.Width = Me.ScaleWidth
    
    rptAdvice.Left = lngLeft
    rptAdvice.Top = fraFilter.Top + fraFilter.Height
    rptAdvice.Width = lngRight - lngLeft
    rptAdvice.Height = lngBottom - rptAdvice.Top - fraAdviceUD.Height - tbcAppend.Height
    
    fraAdviceUD.Left = lngLeft
    fraAdviceUD.Top = rptAdvice.Top + rptAdvice.Height
    fraAdviceUD.Width = rptAdvice.Width
    
    tbcAppend.Left = lngLeft
    tbcAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tbcAppend.Width = rptAdvice.Width
    Me.Refresh
End Sub

Private Sub PrintRIS()
    Dim i As Long, j As Long
    Dim lngResult As Long
    Dim lngҽ��ID As Long

    If HaveRIS Then
        '����
        For i = 0 To rptAdvice.Rows.Count - 1
            If Not rptAdvice.Rows(i).GroupRow Then
                If rptAdvice.Rows(i).Record.Tag = "1" Then
                    lngҽ��ID = Val(rptAdvice.Rows(i).Record(COL_ҽ��ID).value)
                    lngResult = -1
                    lngResult = gobjRis.HISPrintOneRisScheduleRpt(lngҽ��ID)
                    j = j + 1
                End If
            End If
        Next
        If j = 0 Then
            MsgBox "δ��ѡ�κ���Ŀ��", vbInformation, gstrSysName
            rptAdvice.SetFocus: Exit Sub
        End If
    End If
End Sub

Private Sub cmdFilter_Click()
    Call ExecuteFindPati
End Sub

Private Sub dtpBegin_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()

    Dim datCur As Date
 
    On Error GoTo errH
    
    mstrPrivs = gMainPrivs
    
    Call InitCommandBar
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "���븽��", rtfAppend.hwnd, 0).Tag = "����"
    End With
    
    Call InitReportColumn
    
    mstrMatch = IIF(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")
    
    datCur = zlDatabase.Currentdate
    
    dtpBegin.value = Format(datCur - 1, "yyyy-MM-dd 00:00:00")
    dtpEnd.value = Format(datCur + 1, "yyyy-MM-dd 23:59:59")
    With cboFind
        .Clear
        .AddItem "����"
        .AddItem "����"
        .AddItem "סԺ��"
        .ListIndex = 0
    End With
    mstrFindType = "����"
    Call InitUnits
    Call LoadDept
    Call LoadAdvice
    mintPreDept = -1
    Call RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptAdvice_SelectionChanged()
'���ܣ���ʾ����
    Dim lngҽ��ID As Long
    If rptAdvice.SelectedRows.Count = 0 Then Exit Sub
    With rptAdvice.SelectedRows(0)
        lngҽ��ID = Val(.Record(COL_ҽ��ID).value)
    End With
    Call ShowAppend(lngҽ��ID)
End Sub

Private Sub LoadDept()
'���ܣ�����ѡ����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objItem As ListItem
    
    On Error GoTo errH
    
    txtDept.Text = "���п���"
    txtDept.ToolTipText = "���п���"
    txtDept.Tag = ""
    picDept.Visible = False
    txtFind.Text = ""
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 1500
        .Add , "����", "����", 900
    End With
    
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    strSql = "select distinct ID,����,����" & _
        " from ���ű� D,��������˵�� T,�������Ҷ�Ӧ a" & _
        " where D.ID=T.����ID and t.��������=[1] and d.id=a.����id and a.����id=[2]" & _
        " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
        " order by d.����"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "�ٴ�", mlng����ID)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsTmp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsTmp!����
        objItem.Checked = False
        rsTmp.MoveNext
    Loop
    
    'û��ʱ�˳�
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    
    Exit Sub
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then 'ȫѡ Ctrl+A
        Call SetSelect(lvwItems, True)
    End If
    
    If KeyCode = vbKeyR And Shift = vbCtrlMask Then 'ȫ�� Ctrl+R
        Call SetSelect(lvwItems, False)
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIF(mstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIF(mstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "û���ҵ������ҵĿ��ҡ�", vbInformation, Me.Caption
        Else
            MsgBox "�Ѿ������һ�������ˡ�", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal blnSelect As Boolean = True)
    Dim i As Integer
    
    With lvwObj
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str���� As String
    Dim str����IDs As String
    Dim strTmp As String
    Dim varArr As Variant
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
  
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked Then
            strTmp = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            If InStr(str����, strTmp) = 0 Then str���� = str���� & ";" & strTmp
        End If
    Next
    If str���� = "" Then
        txtDept.Text = "���п���"
        txtDept.ToolTipText = "���п���"
        txtDept.Tag = ""
        picDept.Visible = False
        txtFind.Text = ""
        Exit Sub
    End If
    str���� = Mid(str����, 2)
    
    varArr = Split(str����, ";"): strTmp = ""
    
    For i = 0 To UBound(varArr)
        strTmp = strTmp & "," & Split(varArr(i), ",")(1)
        str����IDs = str����IDs & "," & Split(varArr(i), ",")(0)
    Next
    
    txtDept.Text = Mid(strTmp, 2)
    txtDept.ToolTipText = txtDept.Text
    txtDept.Tag = Mid(str����IDs, 2)
    picDept.Visible = False
    txtFind.Text = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDFINDCANCLE,LVWITEMS,PICDEPT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    txtFind.Text = ""
    mlngFind = 1
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIF(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdDept_Click()
'���ܣ���ʾ����ѡ����
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim lngTmp  As Long
    Dim i As Integer
    
    With Me.picDept
        .Left = txtDept.Left
        .Width = txtDept.Width + 700
        .Top = txtDept.Top + txtDept.Height + fraFilter.Top
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = 0
        .Top = txtFind.Height + 100
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height - txtFind.Height - 50 - 50
        txtFind.Top = 50
        cmdFind.Top = 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        .SetFocus
        .Refresh
    End With
    
    Call SetSelect(lvwItems, False)
    If txtDept.Tag = "" Then Exit Sub
   
    For i = 1 To lvwItems.ListItems.Count
        lngTmp = Val(Mid(lvwItems.ListItems(i).Key, 2))
        Me.lvwItems.ListItems(i).Checked = InStr("," & txtDept.Tag & ",", "," & lngTmp & ",") > 0
    Next
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
        .IconsWithShadow = True
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
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, " ��ѯ"): objControl.BeginGroup = True
        objControl.ToolTipText = "��ȡRIS������������"
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "ȫѡ")
        objControl.BeginGroup = True
        objControl.ToolTipText = "ѡ�����п��Դ�ӡ�������(Ctrl+A)"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "ȫ��")
        objControl.ToolTipText = "���������ѡ���������ѡ��״̬(Ctrl+R)"
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        objControl.ToolTipText = "���Ѿ���ѡ�ļ�����뵥ִ�д�ӡ�Ĳ���"
        objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print
        .Add 0, vbKeyF5, conMenu_View_Refresh
        .Add FALT, vbKeyX, conMenu_File_Exit
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Refresh
        Call LoadAdvice
    Case conMenu_Edit_SelAll
        Call SelAllCls(True)
    Case conMenu_Edit_ClsAll
        Call SelAllCls(False)
    Case conMenu_File_Print
        Call PrintRIS
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not picDept.Visible Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then
            cbsMain.FindControl(, conMenu_Edit_SelAll).Execute
        ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
            cbsMain.FindControl(, conMenu_Edit_ClsAll).Execute
        End If
        If KeyCode = vbKeyF3 Then
            If txtFilter.Text = "" Then
                txtFilter.SetFocus
            Else
                Call ExecuteFindPati(True)
            End If
        End If
    End If
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If rptAdvice.Height + Y < 1000 Or tbcAppend.Height - Y < 500 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        rptAdvice.Height = rptAdvice.Height + Y
        tbcAppend.Top = tbcAppend.Top + Y
        tbcAppend.Height = tbcAppend.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn
    
    With rptAdvice
        Set objCol = .Columns.Add(COL_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("UnCheck").Index - 1
        Set objCol = .Columns.Add(COL_����, "����", 120, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 100, True)
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 120, True)
        Set objCol = .Columns.Add(COL_ԤԼʱ��, "ԤԼʱ��", 120, True)
        Set objCol = .Columns.Add(COL_ԤԼ�豸, "ԤԼ�豸", 120, True)
        Set objCol = .Columns.Add(COL_ԤԼ��, "ԤԼ��", 60, True)
        Set objCol = .Columns.Add(COL_ԤԼ��Ŀ, "ԤԼ��Ŀ", 120, True)
	Set objCol = .Columns.Add(COL_��ӡʱ��, "��ӡʱ��", 120, True)
        Set objCol = .Columns.Add(COL_��ӡ��, "��ӡ��", 120, True)
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False)
        Set objCol = .Columns.Add(COL_ҽ��ID, "ҽ��ID", 0, False)
        
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

Private Sub LoadAdvice()
'���ܣ����ز���ҽ���б�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strDepts  As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    If dtpBegin.value >= dtpEnd.value Then
        MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
        dtpBegin.SetFocus: Exit Sub
    End If
    
    strDepts = txtDept.Tag
    
    strSql = "select b.����,b.סԺ��,b.��Ժ���� As ����,nvl(c.Ӥ���Ա�,b.�Ա�) as �Ա�,b.����,e.���� as ����,f.ԤԼ��ʼʱ�� as ԤԼʱ��,f.����豸���� as ԤԼ�豸,f.��� as ԤԼ��,a.ҽ������ as ԤԼ��Ŀ," & vbNewLine & _
        "a.id as ҽ��ID,a.����id,a.��ҳid,c.��� as Ӥ��,c.Ӥ������,Round(Decode(c.����ʱ��, Null, Sysdate, c.����ʱ��) - c.����ʱ��)||'��' As Ӥ������,b.��������,to_char(f.��ӡʱ��,'YYYY-MM-DD HH24:MI') as ��ӡʱ��,f.��ӡ��" & vbNewLine & _
        "from ����ҽ����¼ a,������ҳ b,������������¼ c,���ű� e,RIS���ԤԼ f,����ҽ������ g" & vbNewLine & _
        "where a.����id=b.����id and a.��ҳid=b.��ҳid and a.����id=c.����id(+) and a.��ҳid=c.��ҳid(+) and a.Ӥ��=c.���(+) and a.id=g.ҽ��id and nvl(g.ִ�й���,0)<3 " & vbNewLine & _
        "and a.��������id=e.id and a.id=f.ҽ��id And b.��ǰ����id =[1] And f.ԤԼ���� between [2] and [3] And nvl(f.�Ƿ��ӡ,0)=[4]" & _
        IIF(strDepts = "", "", "  and a.��������id in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X)") & _
        " order by a.����id,a.��ҳid,a.Ӥ��,f.���"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, CDate(dtpBegin.value), CDate(dtpEnd.value), IIF(optType(eδ��).value, 0, 1), strDepts)
    
    With rptAdvice
        .Records.DeleteAll
        For i = 1 To rsTmp.RecordCount
            Set objRecord = .Records.Add()
            objRecord.Tag = "0"
            
            Set objItem = objRecord.AddItem("") 'ѡ����
                strTmp = rsTmp!���� & ""
                If rsTmp!Ӥ������ & "" <> "" Then
                    strTmp = strTmp & "֮Ӥ(" & rsTmp!Ӥ������ & ")"
                End If
            Set objItem = objRecord.AddItem(strTmp)
                objItem.Icon = img16.ListImages.Item(IIF(rsTmp!�Ա� & "" = "��", "Man", "Woman")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!סԺ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
            Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
            
                If InStr("," & rsTmp!Ӥ������ & ",", ",��,") > 0 Then
                    strTmp = rsTmp!���� & ""
                Else
                    strTmp = rsTmp!Ӥ������ & ""
                End If
            Set objItem = objRecord.AddItem(strTmp) '����
            
            Set objItem = objRecord.AddItem(rsTmp!���� & "") '����
                 strTmp = Format(rsTmp!ԤԼʱ��, "yyyy-MM-dd HH:mm")
            Set objItem = objRecord.AddItem(strTmp) 'ԤԼʱ��
            Set objItem = objRecord.AddItem(rsTmp!ԤԼ�豸 & "") 'ԤԼ�豸
            Set objItem = objRecord.AddItem(rsTmp!ԤԼ�� & "") 'ԤԼ��
            Set objItem = objRecord.AddItem(rsTmp!ԤԼ��Ŀ & "")  'ԤԼ��Ŀ
	    Set objItem = objRecord.AddItem(rsTmp!��ӡʱ�� & "")
            Set objItem = objRecord.AddItem(rsTmp!��ӡ�� & "")
            
            Set objItem = objRecord.AddItem(rsTmp!����ID & "")
            Set objItem = objRecord.AddItem(rsTmp!��ҳID & "")
            Set objItem = objRecord.AddItem(rsTmp!ҽ��ID & "") 'ҽ��ID
        
            '������ɫ
            objRecord.Item(0).ForeColor = zlDatabase.GetPatiColor(NVL(rsTmp!��������))
            For j = 1 To objRecord.Childs.Count - 1
                objRecord.Item(j).ForeColor = objRecord.Item(0).ForeColor
            Next
            objRecord.Item(COL_ѡ��).Icon = img16.ListImages.Item("UnCheck").Index - 1
            objRecord.Tag = "1"
            rsTmp.MoveNext
        Next
        .Populate
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptAdvice.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptAdvice_RowDblClick(rptAdvice.SelectedRows(0), rptAdvice.SelectedRows(0).Record.Item(COL_ѡ��))
        End If
    End If
End Sub

Private Sub rptAdvice_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptAdvice.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptAdvice.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptAdvice.Columns(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptAdvice.Records.Count - 1
                            rptAdvice.Records(i)(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
                            rptAdvice.Rows(i).Record.Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptAdvice.Columns(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptAdvice.Records.Count - 1
                            rptAdvice.Records(i)(COL_ѡ��).Icon = -1
                            rptAdvice.Rows(i).Record.Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptAdvice_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COL_ѡ��).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COL_ѡ��).Icon = img16.ListImages.Item("UnCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptAdvice.Populate
End Sub

Private Sub SelAllCls(ByVal blnSel As Boolean)
'���ܣ�ȫѡ����ȫ��
'������blnSel true -ѡȫ��false -ȫ��
    Dim i As Long
     
    If blnSel Then
        rptAdvice.Columns(COL_ѡ��).Caption = "1"
        rptAdvice.Columns(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
        For i = 0 To rptAdvice.Records.Count - 1
            rptAdvice.Records(i)(COL_ѡ��).Icon = img16.ListImages("UnCheck").Index - 1
            rptAdvice.Rows(i).Record.Tag = "1"
        Next
    Else
        rptAdvice.Columns(COL_ѡ��).Caption = ""
        rptAdvice.Columns(COL_ѡ��).Icon = img16.ListImages("Check").Index - 1
        For i = 0 To rptAdvice.Records.Count - 1
            rptAdvice.Records(i)(COL_ѡ��).Icon = -1
            rptAdvice.Rows(i).Record.Tag = "0"
        Next
    End If
    rptAdvice.Populate
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim strUnits As String
    
    On Error GoTo errH
    
    strUnits = GetUser����IDs
    
    cboUnit.Clear
    
    If InStr(mstrPrivs, "ȫԺ����") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSql = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=A.����ID)" & _
            " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=A.����ID)" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSql & ") Group by ID,����,���� Order by ����"
    End If
     
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
                If rsTmp!ID = UserInfo.����ID Then 'ֱ����������
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
                If InStr("," & strUnits & ",", "," & rsTmp!ID & ",") > 0 And cboUnit.ListIndex = -1 Then
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            Else '����ȱʡ���������Ŀ����ж��
                If rsTmp!ȱʡ = 1 And cboUnit.ListIndex = -1 Then
                    Call cbo.SetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call cbo.SetIndex(cboUnit.hwnd, 0)
    End If
    
    Call cbo.Locate(cboUnit, mlng����ID, True)
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Sub cboUnit_Click()
'���ܣ�ˢ�½�������
'˵�����Ӹ��¼���ʼ�᲻�ظ�������ص����ݶ�ȡ

    If cboUnit.ListIndex = mintPreDept Then Exit Sub
    mintPreDept = cboUnit.ListIndex
    mlng����ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
    Call LoadDept
    Call LoadAdvice
End Sub

Private Sub ShowAppend(ByVal lngҽ��ID As Long)
'���ܣ���ʾָ��ҽ���ĵ��ݸ�������
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, lngIdx As Long
     
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    If lngҽ��ID = 0 Then Exit Sub
    strSql = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!��Ŀ & "��" & NVL(rsTmp!����)
                lngIdx = .Find(rsTmp!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!��Ŀ & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!��Ŀ & "��")
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    
    If txtFilter.Text = "" Then
        txtFilter.SetFocus
        Exit Sub
    End If
            
    '��ʼ������
    If rptAdvice.SelectedRows.Count > 0 Then
        If Not rptAdvice.SelectedRows(0).GroupRow Then
            If Val(rptAdvice.SelectedRows(0).Record(COL_����ID).value) <> 0 Then blnHave = True
        End If
    End If
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl����������0��ʼ
    Else
        i = rptAdvice.SelectedRows(0).Index + 1
    End If
    
    '���Ҳ���
    For i = i To rptAdvice.Rows.Count - 1
        With rptAdvice.Rows(i)
            If Not .GroupRow Then
                If mstrFindType = "����" Then
                    If UCase(Trim(.Record(COL_����).value)) = UCase(txtFilter.Text) Then Exit For
                ElseIf mstrFindType = "סԺ��" Then
                    If .Record(COL_סԺ��).value = txtFilter.Text Then Exit For
                ElseIf mstrFindType = "����" Then
                    If .Record(COL_����).value Like "*" & txtFilter.Text & "*" Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptAdvice.Rows.Count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptAdvice.FocusedRow = rptAdvice.Rows(i)
        If rptAdvice.Visible Then rptAdvice.SetFocus
    Else
        blnReStart = True
        MsgBox IIF(blnNext, "������", "") & "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
    If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtFilter_GotFocus()
    zlControl.TxtSelAll txtFilter
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    Select Case mstrFindType
        Case "סԺ��"
            If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case "����"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case "����"
    End Select
    If KeyAscii = 13 Then
        Call ExecuteFindPati
        If 13 = KeyAscii Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
