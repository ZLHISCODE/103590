VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAuditResponse 
   AutoRedraw      =   -1  'True
   Caption         =   "������鷴��"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "frmAuditResponse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11385
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "סԺ��"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ƴ������"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox picPati 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2270
         Picture         =   "frmAuditResponse.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����(F4)"
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ɸѡ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   60
         Width           =   840
      End
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   2730
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   1170
   End
   Begin RichTextLib.RichTextBox txtResponse 
      Height          =   765
      Left            =   720
      TabIndex        =   1
      Top             =   6225
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1349
      _Version        =   393217
      BackColor       =   14737632
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAuditResponse.frx":0680
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
      Left            =   1740
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":071D
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":0CB7
            Key             =   "����״̬_�ȴ�����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":1251
            Key             =   "����״̬_�����ݴ�"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":7AB3
            Key             =   "����״̬_�ȴ�����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":804D
            Key             =   "����״̬_����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":85E7
            Key             =   "����_ҽ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":8B81
            Key             =   "����_����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":911B
            Key             =   "����_����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":96B5
            Key             =   "����_��ҳ"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":9C4F
            Key             =   "����_����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":A1E9
            Key             =   "����_�ļ�"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditResponse.frx":A783
            Key             =   "����_·��"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7185
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAuditResponse.frx":10FE5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17171
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
   Begin RichTextLib.RichTextBox txtNote 
      Height          =   765
      Left            =   5670
      TabIndex        =   2
      Top             =   6255
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1349
      _Version        =   393217
      BorderStyle     =   0
      MaxLength       =   255
      Appearance      =   0
      TextRTF         =   $"frmAuditResponse.frx":11877
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
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   75
      ScaleHeight     =   5325
      ScaleWidth      =   11190
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   570
      Width           =   11190
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   5145
         Left            =   60
         TabIndex        =   0
         Top             =   90
         Width           =   11010
         _Version        =   589884
         _ExtentX        =   19420
         _ExtentY        =   9075
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   165
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditResponse.frx":11914
      Left            =   1245
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   630
      Top             =   135
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmAuditResponse.frx":11928
   End
End
Attribute VB_Name = "frmAuditResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Closed(ByVal DataChange As Boolean)
Public Event OpenObject(ByVal PatiID As Long, ByVal PageID As Long, ByVal ObjectType As Integer, ByVal ObjectID As String)

Private Enum ICON_ID
    conIcon_UnCheck = 1
    conIcon_Check = 2
    conIcon_UnSelect = 3
    conIcon_Select = 4
End Enum
Private Enum MENU_ID
    conMenu_FilterLable = 0
    conMenu_Submit = 1
    conMenu_Random = 2
    conMenu_Await = 3
    conMenu_Done = 4
    conMenu_DateLable = 5
    conMenu_DateInput = 6
    conMenu_Pati = 7
    
    conMenu_Refresh = 90
    conMenu_OpenData = 91
    
    conMenu_Save = 92
    conMenu_Commit = 93
    conMenu_CommitOne = 94
    conMenu_Cancel = 95
    
    conMenu_Help = 98
    conMenu_Exit = 99
    
    conMenu_Col = 100
    
    conMenu_AllCollapse = 201
    conMenu_AllExpend = 202
    conMenu_CurCollapse = 203
    conMenu_CurExpend = 204
End Enum

Private Enum COLUMN_PATI
    pcol_סԺ�� = 1
    pcol_���� = 2
    pcol_�Ա� = 3
    pcol_���� = 4
    pcol_���� = 5
    pcol_ƴ������ = 6
End Enum

Private Enum COLUMN_ID
    col_״̬ = 0 '��ʾ״̬ͼ��
    col_���� = 1
    col_סԺ�� = 2
    col_���� = 3
    col_�Ա� = 4
    col_���� = 5
    col_���� = 6
    
    col_�������� = 7 '��ʾ����ͼ��
    col_������� = 8
    col_����˵�� = 9
    col_�������� = 10
    col_������ = 11
    col_����ʱ�� = 12
    
    col_����˵�� = 13
    col_������ = 14
    col_����ʱ�� = 15
    col_��ֵ = 16
    
    col_����Id = 17
    col_��ҳID = 18
    col_����ID = 19
    col_���ID = 20
    col_����ID = 21
    col_���ĵ�ID = 22 '�°没��
End Enum

Private mstrPrivs As String
Private mlngDeptID As Long '����/����ID
Private mintDeptType As Integer '0-��������ʾ��1-��������ʾ
Private mintDataType As Integer '0-ҽ��վ��,1-��ʿվ��
Private mblnICU As Boolean '�Ƿ�Ǳ��Ƶ�ICU����

Private Type FilterCond
    �ύ��� As Boolean
    ������ As Boolean
    δ���� As Boolean
    �Ѵ��� As Boolean
    ��ʼʱ�� As Date
    ����ʱ�� As Date
End Type
Private mvarCond As FilterCond
Private mblnEditing As Boolean
Private mintPreTime As Integer
Private mblnOpen As Boolean
Private mblnOK As Boolean

Public Function ShowMe(frmParent As Object, ByVal lngDeptID As Long, ByVal intDeptType As Integer, _
    ByVal blnICU As Boolean, ByVal intDataType As Integer, ByVal strPrivs As String) As Boolean
    mlngDeptID = lngDeptID
    mintDeptType = intDeptType
    mblnICU = blnICU
    mintDataType = intDataType
    mstrPrivs = strPrivs
        
    If mblnOpen Then
        '����ˢ������
        '###
            
        If Me.WindowState = vbMinimized Then
            Me.WindowState = vbNormal
        End If
    End If
    
    Me.Show , frmParent
End Function

Private Sub cboTime_Click()
    Dim curDate As Date
    
    If cboTime.ListIndex = mintPreTime Then Exit Sub
    
    curDate = zlDatabase.Currentdate
    
    Select Case cboTime.Text
    Case "����"
        mvarCond.��ʼʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "����"
        mvarCond.��ʼʱ�� = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mvarCond.��ʼʱ�� = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mvarCond.��ʼʱ�� = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mvarCond.��ʼʱ�� = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mvarCond.��ʼʱ�� = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[ָ��..]"
        If Not frmSelectTime.ShowMe(Me, mvarCond.��ʼʱ��, mvarCond.����ʱ��, cboTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call Cbo.SetIndex(cboTime.hwnd, mintPreTime)
            rptData.SetFocus: Exit Sub
        Else
            rptData.SetFocus
        End If
    End Select
        
    cboTime.ToolTipText = "��Χ��" & Format(mvarCond.��ʼʱ��, "yyyy-MM-dd") & " �� " & Format(mvarCond.����ʱ��, "yyyy-MM-dd")
    mintPreTime = cboTime.ListIndex
    Me.Refresh
    
    'ˢ������
    Call RefreshData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    
    Select Case Control.ID
        Case conMenu_Submit
            Control.IconId = IIf(mvarCond.�ύ���, conIcon_Check, conIcon_UnCheck)
            Control.Checked = IIf(mvarCond.�ύ���, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Random
            Control.IconId = IIf(mvarCond.������, conIcon_Check, conIcon_UnCheck)
            Control.Checked = IIf(mvarCond.������, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Await
            Control.IconId = IIf(mvarCond.δ����, conIcon_Select, conIcon_UnSelect)
            Control.Checked = IIf(mvarCond.δ����, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_Done
            Control.IconId = IIf(mvarCond.�Ѵ���, conIcon_Select, conIcon_UnSelect)
            Control.Checked = IIf(mvarCond.�Ѵ���, True, False)
            Control.Enabled = Not mblnEditing
        Case conMenu_DateLable, conMenu_DateInput
            Control.Visible = mvarCond.�Ѵ���
            Control.Enabled = Not mblnEditing
        Case conMenu_OpenData '�򿪶�λ
            If InStr(mstrPrivs, "��鷴������") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = False
                If rptData.SelectedRows.Count > 0 Then
                    With rptData.SelectedRows(0)
                        If Not .GroupRow And .Childs.Count = 0 Then
                            blnEnabled = .Record(col_״̬).Value = 1 Or .Record(col_״̬).Value = 2
                        End If
                    End With
                End If
                Control.Enabled = blnEnabled And Not mblnEditing
            End If
        Case conMenu_Refresh
            Control.Enabled = Not mblnEditing
        Case conMenu_Save '�ݴ�
            Control.Enabled = mblnEditing
            Control.Visible = mvarCond.δ���� '�Ѵ���ģ�������ִ���ݴ�
        Case conMenu_CommitOne  '��ɵ���
            If mvarCond.δ���� Then
                Control.Caption = "��ɵ���"
                Control.ToolTipText = "����ǰ�ݴ��еķ��������ύ�ٴ����"
            Else
                Control.Caption = "����"
                Control.ToolTipText = "�����޸ĵĴ������"
            End If
            
            Control.Enabled = mblnEditing
            
            '�ݴ�ģ�����ֱ����ɵ���
            If mblnEditing = False And mvarCond.δ���� Then
                If rptData.SelectedRows.Count > 0 Then
                    With rptData.SelectedRows(0)
                        If Not .GroupRow And .Childs.Count = 0 Then
                            Control.Enabled = .Record(col_״̬).Value = 1 And .Record(col_������).Value <> ""
                        End If
                    End With
                End If
            End If
        Case conMenu_Commit     '��ɣ��ύ�����ݴ棩
            Control.Visible = mvarCond.δ����
            Control.Enabled = Not mblnEditing
        
        Case conMenu_Cancel 'ȡ��
            Control.Enabled = mblnEditing
        Case conMenu_Col + 1 To conMenu_Col + 99 '��ʾ/������
            Control.Checked = rptData.Columns.Find(Val(Control.Parameter)).Visible
            Control.Enabled = Not mblnEditing
        Case conMenu_CurExpend 'չ����ǰ��
            blnEnabled = False
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptData.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled And Not mblnEditing
        Case conMenu_CurCollapse '�۵���ǰ��
            blnEnabled = False
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    blnEnabled = rptData.SelectedRows(0).Expanded
                ElseIf Not rptData.SelectedRows(0).ParentRow Is Nothing Then
                    If rptData.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptData.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled And Not mblnEditing
    End Select
    
    If mblnEditing Then
        txtResponse.Enabled = False
        picData.Enabled = False
    Else
        txtResponse.Enabled = True
        picData.Enabled = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objRow As ReportRow
    
    Select Case Control.ID
        Case conMenu_Submit
            If mvarCond.�ύ��� And Not mvarCond.������ Then Exit Sub
            mvarCond.�ύ��� = Not mvarCond.�ύ���
            Call RefreshData
        Case conMenu_Random
            If mvarCond.������ And Not mvarCond.�ύ��� Then Exit Sub
            mvarCond.������ = Not mvarCond.������
            Call RefreshData
        Case conMenu_Await
            If mvarCond.δ���� Then Exit Sub
            mvarCond.δ���� = True: mvarCond.�Ѵ��� = False
            Call RefreshData
        Case conMenu_Done
            If mvarCond.�Ѵ��� Then Exit Sub
            mvarCond.�Ѵ��� = True: mvarCond.δ���� = False
            Call RefreshData
        Case conMenu_Refresh
            Call RefreshData
        Case conMenu_OpenData '�򿪶�λ
            Call rptData_RowDblClick(rptData.SelectedRows(0), rptData.SelectedRows(0).Record(col_�������))
        Case conMenu_Save   '�ݴ�
            If SaveData(True) Then
                mblnEditing = False
                cbsMain.RecalcLayout
                rptData.SetFocus
            End If
            
        Case conMenu_CommitOne  '��ɵ���
            If Trim(txtNote.Text) = "" And mvarCond.δ���� Then
                Call MsgBox("�����봦��˵����", vbInformation, gstrSysName)
                If txtNote.Enabled And txtNote.Visible Then txtNote.SetFocus
                Exit Sub
            End If
            If SaveData(False) Then
                mblnEditing = False
                cbsMain.RecalcLayout
                rptData.SetFocus
            End If
        Case conMenu_Commit '�������
            
            Call SaveAllPaseData    '���а�������ˢ��
            
        Case conMenu_Cancel
            If txtNote.Text <> rptData.SelectedRows(0).Record(col_����˵��).Value Then
                If MsgBox("ȷʵҪȡ���༭��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            txtNote.Text = rptData.SelectedRows(0).Record(col_����˵��).Value
            mblnEditing = False
            cbsMain.RecalcLayout
            rptData.SetFocus
        Case conMenu_Col + 1 To conMenu_Col + 99 '��ʾ/������
            rptData.Columns.Find(Val(Control.Parameter)).Visible = Not rptData.Columns.Find(Val(Control.Parameter)).Visible
        Case conMenu_CurCollapse '�۵���ǰ��
            If rptData.SelectedRows.Count > 0 Then
                If rptData.SelectedRows(0).GroupRow Then
                    rptData.SelectedRows(0).Expanded = False
                ElseIf Not rptData.SelectedRows(0).ParentRow Is Nothing Then
                    If rptData.SelectedRows(0).ParentRow.GroupRow Then
                        rptData.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '���۵���λ��������,�����Զ�������¼�
            Call rptData_SelectionChanged
        Case conMenu_CurExpend 'չ����ǰ��
            If rptData.SelectedRows.Count > 0 Then
                rptData.SelectedRows(0).Expanded = True
            End If
        Case conMenu_AllCollapse '�۵�������
            For Each objRow In rptData.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '���۵���λ��������,�����Զ�������¼�
            Call rptData_SelectionChanged
        Case conMenu_AllExpend 'չ��������
            For Each objRow In rptData.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_Help
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Exit
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.picData
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = txtResponse.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = txtNote.hwnd
    End If
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    'CommandBars
    '-----------------------------------------------------
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
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = imgMain.Icons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Refresh, "ˢ��")
            objControl.BeginGroup = True
            objControl.ToolTipText = "ˢ�µ�ǰѡ�������"
        Set objControl = .Add(xtpControlButton, conMenu_OpenData, "��")
            objControl.ToolTipText = "�򿪷�������"
            
        Set objControl = .Add(xtpControlButton, conMenu_Save, "�ݴ�")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Commit, "���")
        objControl.ToolTipText = "�����ݴ�ķ�������ȫ���ύ�ٴ����"
        Set objControl = .Add(xtpControlButton, conMenu_CommitOne, "��ɵ���")
        objControl.ToolTipText = "����ǰ�ݴ��еķ��������ύ�ٴ����"
        objControl.ToolTipText = "�����޸ĵĴ������"
        
        Set objControl = .Add(xtpControlButton, conMenu_Cancel, "ȡ��")
            
        Set objControl = .Add(xtpControlButton, conMenu_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
        
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_FilterLable, "��������")
        
        Set objControl = .Add(xtpControlButton, conMenu_Submit, "�ύ���")
        Set objControl = .Add(xtpControlButton, conMenu_Random, "������")
        
        Set objControl = .Add(xtpControlButton, conMenu_Await, "δ����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Done, "�Ѵ���")
        
        Set objControl = .Add(xtpControlLabel, conMenu_DateLable, "����ʱ��")
            objControl.BeginGroup = True
        Set objCustom = .Add(xtpControlCustom, conMenu_DateInput, "����ʱ��")
            objCustom.Handle = cboTime.hwnd
            
        Set objCustom = .Add(xtpControlCustom, conMenu_Pati, "ɸѡ����")
            objCustom.Handle = picPati.hwnd
            picPati.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    
    '�ȼ���:ע�ⲻ�ܺ�ϵͳ���ı��༭�ȼ���ͻ
    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help
        .Add 0, vbKeyF5, conMenu_Refresh
        .Add 0, vbKeyF3, conMenu_OpenData
        .Add 0, vbKeyF2, conMenu_CommitOne
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add 0, vbKeyEscape, conMenu_Cancel
        .Add FALT, vbKeyX, conMenu_Exit
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If txtNote.Enabled And txtNote.Visible And txtNote.Locked = False Then
            txtNote.SetFocus
            Call txtNote_GotFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim curDate As Date
    
    Call InitCommandBar
    
    'ȱʡҽ��ʱ��
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ��..]"
    mintPreTime = 0
    Call Cbo.SetIndex(cboTime.hwnd, 0)
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 320, 100, DockBottomOf, Nothing)
    objPane.Title = "�������"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = Me.dkpMain.CreatePane(2, 320, 100, DockRightOf, objPane)
    objPane.Title = "����˵��"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    '����
    '-----------------------------------------------------
    mblnOpen = True
    mblnOK = False
    
    'ȱʡ����
    curDate = zlDatabase.Currentdate
    With mvarCond
        .�ύ��� = Val(zlDatabase.GetPara("�ύ��鷴��", glngSys, IIf(mintDataType = 0, pסԺҽ��վ, pסԺ��ʿվ), "1")) <> 0
        .������ = Val(zlDatabase.GetPara("�����鷴��", glngSys, IIf(mintDataType = 0, pסԺҽ��վ, pסԺ��ʿվ), "0")) <> 0
        .δ���� = True: .�Ѵ��� = False
        .��ʼʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
        .����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    End With
    
    'ˢ������
    Call RefreshData
    mblnEditing = False
        
    '------------
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnSetup As Boolean

    If mblnEditing Then
        If MsgBox("ȷʵҪȡ���༭���˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("�ύ��鷴��", IIf(mvarCond.�ύ���, 1, 0), glngSys, IIf(mintDataType = 0, pסԺҽ��վ, pסԺ��ʿվ), blnSetup)
    Call zlDatabase.SetPara("�����鷴��", IIf(mvarCond.������, 1, 0), glngSys, IIf(mintDataType = 0, pסԺҽ��վ, pסԺ��ʿվ), blnSetup)
    Call SaveWinState(Me, App.ProductName)

    mblnOpen = False
    RaiseEvent Closed(mblnOK)
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptData
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(col_״̬, "״̬", 75, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.Icon = img16.ListImages("ͼ��").Index - 1
        Set objCol = .Columns.Add(col_����, "����", 60, True)
        Set objCol = .Columns.Add(col_סԺ��, "סԺ��", 62, True)
        Set objCol = .Columns.Add(col_����, "����", 40, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 30, True)
        Set objCol = .Columns.Add(col_����, "����", 30, True)
        Set objCol = .Columns.Add(col_����, IIf(mintDeptType = 0, "����", "����"), 70, True)
        
        Set objCol = .Columns.Add(col_��������, "����", 18, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.AllowRemove = False
            objCol.Icon = img16.ListImages("ͼ��").Index - 1
        Set objCol = .Columns.Add(col_�������, "�������", 200, True)
            objCol.AllowRemove = False
            objCol.Groupable = False
        Set objCol = .Columns.Add(col_����˵��, "����˵��", 120, True)
        Set objCol = .Columns.Add(col_��������, "��������", 80, True)
        Set objCol = .Columns.Add(col_������, "������", 50, True)
        Set objCol = .Columns.Add(col_����ʱ��, "����ʱ��", 80, True)
        
        Set objCol = .Columns.Add(col_����˵��, "����˵��", 200, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(col_������, "������", 50, True)
        Set objCol = .Columns.Add(col_����ʱ��, "����ʱ��", 80, True)
        Set objCol = .Columns.Add(col_��ֵ, "��ֵ", 50, True)
        
        Set objCol = .Columns.Add(col_����Id, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_���ID, "���ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(col_���ĵ�ID, "���ĵ�ID", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�ķ���..."
            .ShadeGroupHeadings = True
        End With
        .ShowGroupBox = True
        .ShowItemsInGroups = False '�Ƿ������Է�������
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_״̬)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_״̬)
        .SortOrder(0).SortAscending = True
                
        .SortOrder.Add .Columns(col_������) '������Ϊ�յģ����������ݴ�ʹ�����
        .SortOrder(1).SortAscending = True
        
        .SortOrder.Add .Columns(col_����Id)
        .SortOrder(2).SortAscending = True
        
        .SortOrder.Add .Columns(col_����ʱ��)
        .SortOrder(3).SortAscending = False
        
    End With
End Sub

Private Sub picData_Resize()
    rptData.Left = 0: rptData.Top = 0
    rptData.Width = picData.ScaleWidth
    rptData.Height = picData.ScaleHeight
End Sub

Private Sub picPati_GotFocus()
    Call txtPati_GotFocus
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHit As ReportHitTestInfo
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim objCol As ReportColumn, lngCount As Long
    
    If Button = 2 Then
        Set objHit = rptData.HitTest(X, Y)
        If objHit.ht = xtpHitTestHeader Then
            Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                lngCount = 1
                For Each objCol In rptData.Columns
                    If objCol.AllowRemove And objCol.Width > 0 Then
                        Set objControl = .Add(xtpControlButton, conMenu_Col + lngCount, objCol.Caption)
                        objControl.Parameter = objCol.ItemIndex
                        lngCount = lngCount + 1
                    End If
                Next
            End With
            objPopup.ShowPopup
        ElseIf objHit.ht = xtpHitTestReportArea Then
            If Not objHit.Row Is Nothing Then
                If objHit.Row.GroupRow Then
                    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
                    With objPopup.Controls
                        Set objControl = .Add(xtpControlButton, conMenu_AllCollapse, "�۵�������")
                        Set objControl = .Add(xtpControlButton, conMenu_AllExpend, "չ��������")
                        Set objControl = .Add(xtpControlButton, conMenu_CurCollapse, "�۵���ǰ��")
                            objControl.BeginGroup = True
                        Set objControl = .Add(xtpControlButton, conMenu_CurExpend, "չ����ǰ��")
                    End With
                    objPopup.ShowPopup
                End If
            End If
        End If
    End If
End Sub

Private Function RefreshData() As Boolean
'���ܣ����ݵ�ǰ���õ�������ȡ��������
Dim rsTmp As ADODB.Recordset, strSQL As String, strReturn As String, rsEmr As New ADODB.Recordset, strSQLEmr As String
Dim strPatis As String, objListItem As ListItem, curDate As Date
Dim objRecord As ReportRecord, objItem As ReportRecordItem, objRow As ReportRow, i As Long
Dim lngPreID As Long, lngPreIdx As Long
    
    Screen.MousePointer = 11
    If lvwPati.Visible = True Then lvwPati.Visible = False
    lvwPati.ListItems.Clear
        
    On Error GoTo errH
    
    '�������ݣ�δ�鵵��ȫ������ʷ�鵵����ʱ��Ϊ׼
    If mintDataType = 0 Then
        strSQL = " And �������� IN(1,2,5,6,7,8,9)" 'ҽ���漰�Ķ���
    ElseIf mintDataType = 1 Then
        strSQL = " And �������� IN(3,4)" '��ʿ�漰�Ķ���
    End If
    If mvarCond.δ���� Then
        strSQL = "Select ID, ���id, ����id, ��ҳid, ��¼����, ��¼״̬, ��������, �ļ�id, �������, ������, ����ʱ��, ��������, ����˵��, ������, ����ʱ��, ����, ��ֵ,����˵��, ���ĵ�id From ����������¼ Where ��¼״̬=1 And Instr([3],��¼����)>0" & strSQL
    ElseIf mvarCond.�Ѵ��� Then
        strSQL = _
            " Select ID, ���id, ����id, ��ҳid, ��¼����, ��¼״̬, ��������, �ļ�id, �������, ������, ����ʱ��, ��������, ����˵��, ������, ����ʱ��, ����, ��ֵ,����˵��, ���ĵ�id From ����������¼ Where ��¼״̬ In(2,3) And Instr([3],��¼����)>0 And ����ʱ�� Between [4] And [5]" & strSQL & _
            " Union ALL" & _
            " Select ID, ���id, ����id, ��ҳid, ��¼����, ��¼״̬, ��������, �ļ�id, �������, ������, ����ʱ��, ��������, ����˵��, ������, ����ʱ��, ����, ��ֵ,����˵��, ���ĵ�id From ����������ʷ Where ��¼״̬ In(2,3) And Instr([3],��¼����)>0 And ����ʱ�� Between [4] And [5]" & strSQL
    End If
    
    'SQL�в��������Ч��,ReportControl��������
    strSQL = _
        " Select NVL(B.����,C.����) ����,NVL(B.�Ա�,C.�Ա�) �Ա�,NVL(B.����,C.����) ����,B.סԺ��,B.��Ժ���� as ����,D.���� as ����," & _
        " A.ID as ����ID,A.���ID,A.����ID,A.��ҳID,A.��¼����,A.��¼״̬," & _
        " A.��������,A.�ļ�ID,A.���ĵ�id,E.�������� as סԺ����,F.���� as �����¼,A.�������,A.����˵��," & _
        " A.������,A.����ʱ��,A.��������,A.����˵��,A.������,A.����ʱ��,Decode(A.����,1,'���ϸ�',A.��ֵ) as ��ֵ" & _
        " From �����ļ��б� F,���Ӳ�����¼ E,���ű� D,������Ϣ C,������ҳ B,(" & strSQL & ") A" & _
        " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And B.����ID=C.����ID And decode(length(a.�ļ�id),32,0,a.�ļ�id)=E.ID(+) And E.�ļ�ID=F.ID(+)" & _
        IIf(mintDeptType = 0, " And B.��Ժ����ID=D.ID", " And B.��ǰ����ID=D.ID") & _
        IIf(mintDeptType = 0, " And B.��Ժ����ID=[1]", " And B.��ǰ����ID=[1]") & _
        IIf(mintDataType = 0, IIf(InStr(mstrPrivs, "���Ʋ���") > 0 Or InStr(mstrPrivs, "ȫԺ����") > 0, "", " And B.סԺҽʦ=[2]"), "") & _
        IIf(mintDataType = 0, IIf(mblnICU And InStr(mstrPrivs, "ȫԺ����") = 0, " And B.סԺҽʦ=[2]", ""), "") & _
        " Order by ����ID,����ʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, UserInfo.����, _
        IIf(mvarCond.������, "1", "") & IIf(mvarCond.�ύ���, "2", ""), mvarCond.��ʼʱ��, mvarCond.����ʱ��)
        
    '��¼����ѡ�еķ���
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow And rptData.SelectedRows(0).Childs.Count = 0 Then
            lngPreIdx = rptData.SelectedRows(0).Index '���ڿ������¶�λ
            lngPreID = rptData.SelectedRows(0).Record(col_����ID).Value
        End If
    End If
    
    txtNote.MaxLength = rsTmp.Fields("����˵��").DefinedSize
    curDate = zlDatabase.Currentdate
    rptData.Records.DeleteAll
    Do While Not rsTmp.EOF
        
        If InStr(strPatis & ",", "," & rsTmp!����ID & ",") = 0 Then
            If strPatis = "" Then
                Set objListItem = lvwPati.ListItems.Add(, "_0_0", "ȫ��")
                objListItem.SubItems(pcol_ƴ������) = "QB"
                txtPati.Text = "ȫ��"
                txtPati.Tag = txtPati.Text
            End If
            strPatis = strPatis & "," & rsTmp!����ID
            
            Set objListItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, rsTmp!����)
            If lvwPati.Tag = "_" & rsTmp!����ID & "_" & rsTmp!��ҳID Then
                objListItem.Selected = True
            End If
            
            objListItem.SubItems(pcol_סԺ��) = Nvl(rsTmp!סԺ��)
            objListItem.SubItems(pcol_����) = Nvl(rsTmp!����)
            objListItem.SubItems(pcol_�Ա�) = Nvl(rsTmp!�Ա�)
            objListItem.SubItems(pcol_����) = Nvl(rsTmp!����)
            objListItem.SubItems(pcol_����) = Nvl(rsTmp!����)
            objListItem.SubItems(pcol_ƴ������) = ZLCommFun.SpellCode(rsTmp!���� & "��0")
        End If
        
        Set objRecord = Me.rptData.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!��¼״̬))
        objItem.Caption = Decode(rsTmp!��¼״̬, 1, IIf(IsNull(rsTmp!������), "�ȴ�����", "�����ݴ�"), 2, "�ȴ�����", 3, "����")
        objItem.Value = Val(rsTmp!��¼״̬)
        objItem.Icon = img16.ListImages("����״̬_" & objItem.Caption).Index - 1

        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!סԺ��, " "))) '��" "��Ϊ���������
        
        Set objItem = objRecord.AddItem(zlStr.Lpad(Nvl(rsTmp!����), 10)) 'Value��������
        objItem.Caption = Nvl(rsTmp!����, " ") 'Ϊ��ʱ�ᱻValue���

        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!�Ա�)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����)))
        Set objItem = objRecord.AddItem(Val(rsTmp!��������))
        objItem.Caption = Decode(rsTmp!��������, 1, "סԺҽ��", 2, "סԺ����", 3, "������", 4, "�����¼", 5, "������ҳ", 6, "ҽ������", 7, "����֤��", 8, "֪���ļ�", 9, "�ٴ�·��")
        objItem.Icon = img16.ListImages("����_" & Decode(rsTmp!��������, 1, "ҽ��", 2, "����", 3, "����", 4, "����", 5, "��ҳ", 6, "����", 7, "�ļ�", 8, "�ļ�", 9, "·��")).Index - 1
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!�������)))
        If rsTmp!�������� = 1 Then
            objItem.Caption = "סԺҽ��"
        ElseIf rsTmp!�������� = 5 Then
            objItem.Caption = "������ҳ"
        ElseIf rsTmp!�������� = 4 Then
            objItem.Caption = Nvl(rsTmp!�����¼)
        ElseIf rsTmp!�������� = 9 Then
            objItem.Caption = "�ٴ�·��"
        ElseIf Len(Nvl(rsTmp!�ļ�ID)) < 32 Then
            objItem.Caption = Nvl(rsTmp!סԺ����)
        ElseIf Len(Nvl(rsTmp!�ļ�ID)) = 32 And Not gobjEmr Is Nothing Then '�°没��
            strSQLEmr = "Select Nvl(b.Subdoc_Title, a.Title) ��������" & vbNewLine & _
                    "From Bz_Doc_Log A, Bz_Doc_Tasks B" & vbNewLine & _
                    "Where a.Id = Hextoraw(:fileid) And a.Id = b.Real_Doc_Id" & IIf(Nvl(rsTmp!���ĵ�ID) = "", "", " And b.Subdoc_Id = :subdocid")
            strReturn = gobjEmr.OpenSQLRecordset(strSQLEmr, rsTmp!�ļ�ID & "^16^fileid" & IIf(Nvl(rsTmp!���ĵ�ID) = "", "", "|" & Nvl(rsTmp!���ĵ�ID) & "^16^subdocid"), rsEmr)
            If strReturn = "" Then
			If rsEmr.EOF THEN
				objItem.Caption = "��ԭʼ�����Ѳ����ڡ�"
			ELSE
                objItem.Caption = Nvl(rsEmr!��������)
			END If
            End If
        End If
        objItem.Caption = IIf(objItem.Caption <> "", objItem.Caption & ":", "") & Nvl(rsTmp!�������)
        If rsTmp!��¼״̬ = 1 Then objItem.Bold = True
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����˵��)))
            objItem.Caption = "" & Nvl(rsTmp!����˵��)
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!��������), "yyyy-MM-dd HH:mm")))
        If Not IsNull(rsTmp!��������) And rsTmp!��¼״̬ = 1 Then
            If curDate > rsTmp!�������� Then objItem.ForeColor = vbRed
        End If
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!������)))
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")))
        
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!����˵��)))
        Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!������)))
        Set objItem = objRecord.AddItem(CStr(Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")))
        Set objItem = objRecord.AddItem("" & rsTmp!��ֵ)
        
        Set objItem = objRecord.AddItem(Val(rsTmp!����ID))
        Set objItem = objRecord.AddItem(Val(rsTmp!��ҳID))
        Set objItem = objRecord.AddItem(Val(rsTmp!����ID))
        Set objItem = objRecord.AddItem(Val(Nvl(rsTmp!���ID, 0)))
        Set objItem = objRecord.AddItem(Nvl(rsTmp!�ļ�ID, "0"))
        Set objItem = objRecord.AddItem(Nvl(rsTmp!���ĵ�ID, ""))
        
        rsTmp.MoveNext
    Loop
    rptData.Populate
    
    '��λ��֮ǰѡ��Ĳ�����
    If Not (lvwPati.Tag = "" Or lvwPati.Tag = "_0_0") Then
        If Not lvwPati.SelectedItem Is Nothing Then
            Call lvwPati_KeyPress(vbKeyReturn)
        End If
    End If
        
    If rptData.Rows.Count = 0 Then
        txtNote.Locked = True '֮ǰ��Lock�Է����ж�
        txtNote.BackColor = txtResponse.BackColor
        txtResponse.Text = "": txtNote.Text = ""
        Me.stbThis.Panels(2).Text = ""
    Else
        If lngPreID <> 0 Then
            '�ȿ��ٶ�λ
            If lngPreIdx <= rptData.Rows.Count - 1 Then
                If Not rptData.Rows(lngPreIdx).GroupRow And rptData.Rows(lngPreIdx).Childs.Count = 0 Then
                    If rptData.Rows(lngPreIdx).Record(col_����ID).Value = lngPreID Then
                        Set objRow = rptData.Rows(lngPreIdx)
                    End If
                End If
            End If
            '�ٽ��в���
            If objRow Is Nothing Then
                For i = 0 To rptData.Rows.Count - 1
                    If Not rptData.Rows(i).GroupRow And rptData.Rows(i).Childs.Count = 0 Then
                        If rptData.Rows(i).Record(col_����ID).Value = lngPreID Then
                            Set objRow = rptData.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        'ȡ��һ���Ƿ�����
        If objRow Is Nothing Then
            For i = 0 To rptData.Rows.Count - 1
                If Not rptData.Rows(i).GroupRow And rptData.Rows(i).Childs.Count = 0 Then Set objRow = rptData.Rows(i): Exit For
            Next
        End If
        Set rptData.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        
        'ѡ��ĳ������ʱ��������¼�����ص�
        If lvwPati.Tag = "" Or lvwPati.Tag = "_0_0" Then
            Me.stbThis.Panels(2).Text = "���� " & rptData.Records.Count & " ��������¼"
        Else
            Me.stbThis.Panels(2).Text = ""
        End If
    End If
        
    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim lngPatID As Long, lngPageID As Long, intObject As Integer, strObjectID As String, strSubdocID As String
    If InStr(mstrPrivs, "��鷴������") = 0 Then Exit Sub
    
    If Not Row.GroupRow And Row.Childs.Count = 0 Then
        If Row.Record(col_״̬).Value = 1 Or Row.Record(col_״̬).Value = 2 Then
            lngPatID = CLng(Row.Record(col_����Id).Value)
            lngPageID = CLng(Row.Record(col_��ҳID).Value)
            intObject = CInt(Row.Record(col_��������).Value)
            strObjectID = CStr(Row.Record(col_����ID).Value)
            If Len(strObjectID) = 32 Then
                strSubdocID = CStr(Row.Record(col_���ĵ�ID).Value)
                If strSubdocID <> "" Then strObjectID = strObjectID & "|" & strSubdocID
            End If
            RaiseEvent OpenObject(lngPatID, lngPageID, intObject, strObjectID)
        End If
    End If
End Sub

Private Sub rptData_SelectionChanged()
    Dim blnData As Boolean, blnModi As Boolean
    
    '��ʾ��ϸ����
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow And rptData.SelectedRows(0).Childs.Count = 0 Then
            blnData = True
        End If
    End If
    
    '�ɷ��޸Ĵ���˵��
    If blnData And InStr(mstrPrivs, "��鷴������") > 0 Then
        With rptData.SelectedRows(0)
            If .Record(col_״̬).Value = 1 Or .Record(col_״̬).Value = 2 And .Record(col_������).Value = UserInfo.���� Then
                blnModi = True
            End If
        End With
    End If
    txtNote.Locked = Not blnModi '֮ǰ��Lock�Է����ж�
    txtNote.BackColor = IIf(blnModi, vbWindowBackground, txtResponse.BackColor)
    
    If blnData Then
        txtResponse.Text = rptData.SelectedRows(0).Record(col_�������).Value
        txtNote.Text = rptData.SelectedRows(0).Record(col_����˵��).Value
    Else
        txtResponse.Text = "": txtNote.Text = ""
    End If
End Sub

Private Sub txtNote_Change()
    If Visible And Not txtNote.Locked And rptData.SelectedRows.Count > 0 And Not mblnEditing Then
        If txtNote.Text <> rptData.SelectedRows(0).Record(col_����˵��).Value Then
            mblnEditing = True
        End If
    End If
End Sub

Private Sub txtNote_GotFocus()
    txtNote.SelStart = 0
    txtNote.SelLength = Len(txtNote.Text)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Function SaveAllPaseData() As Boolean
'���ܣ���������ݴ�ķ�������
    Dim colsql As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, strRows As String
        
    With rptData
        For i = 0 To .Records.Count - 1
            If .Records(i)(col_״̬).Value = 1 And .Records(i)(col_������).Value = UserInfo.���� Then
                strSQL = "Zl_����������¼_Process(" & .Records(i)(col_����ID).Value & ",2,'" & .Records(i)(col_����˵��).Value & "',1)"
                colsql.Add strSQL, "C" & colsql.Count + 1
            End If
        Next
        
        If colsql.Count = 0 Then
            MsgBox "��û���ݴ�ķ��������¼�����������ɲ�����", vbInformation, gstrSysName
            Exit Function
                
        ElseIf MsgBox("ȷʵҪ�����ݴ�����з������������ɲ�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
        
        On Error GoTo errH
        If colsql.Count > 0 Then
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colsql.Count
                    Call zlDatabase.ExecuteProcedure(colsql("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            
            'ˢ�½���
            Call RefreshData
        End If
    End With
        
    SaveAllPaseData = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(ByVal blnPauseSave As Boolean) As Boolean
'������blnPauseSave-True=�ݴ棬-False��ɵ�ǰ��
    Dim curDate As Date
    Dim strSQL As String, blnDel As Boolean
       
    If rptData.SelectedRows.Count = 0 Then Exit Function
    If rptData.SelectedRows(0).GroupRow Or rptData.SelectedRows(0).Childs.Count > 0 Then Exit Function
        
    
    With rptData.SelectedRows(0)
        '�Ǳ༭״̬�¿ɡ���ɵ�����������˵��δ�䣩
        If .Record(col_����˵��).Value <> txtNote.Text Or mvarCond.δ���� And Not blnPauseSave Then
            '������
            If ZLCommFun.ActualLen(txtNote.Text) > txtNote.MaxLength Then
                MsgBox "����˵������̫����������� " & txtNote.MaxLength \ 2 & " �����ֻ� " & txtNote.MaxLength & " ���ַ���", vbInformation, gstrSysName
                txtNote.SetFocus: Exit Function
            End If
            
            '����ȷ��
            curDate = zlDatabase.Currentdate
            If txtNote.Text <> "" Then
                strSQL = "Zl_����������¼_Process(" & .Record(col_����ID).Value & ",2,'" & Replace(txtNote.Text, "'", "''") & "'," & IIf(blnPauseSave, 0, 1) & ")"
                On Error GoTo errH
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                On Error GoTo 0
                                
                If Not blnPauseSave And mvarCond.δ���� Then
                    blnDel = True
                Else
                    .Record(col_����˵��).Value = txtNote.Text
                    .Record(col_����ʱ��).Value = CStr(Format(curDate, "yyyy-MM-dd HH:mm"))
                    .Record(col_������).Value = UserInfo.����
                    .Record(col_�������).Bold = False
                    .Record(col_��������).ForeColor = Me.ForeColor
                    .Record(col_״̬).Value = IIf(blnPauseSave, 1, 2)
                    .Record(col_״̬).Caption = IIf(blnPauseSave, "�����ݴ�", "�ȴ�����")
                    .Record(col_״̬).Icon = img16.ListImages("����״̬_" & IIf(blnPauseSave, "�����ݴ�", "�ȴ�����")).Index - 1
                End If
            Else
                strSQL = "Zl_����������¼_Process(" & .Record(col_����ID).Value & ",1)"
                On Error GoTo errH
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                On Error GoTo 0
                
                If mvarCond.�Ѵ��� Then
                    blnDel = True
                Else
                    .Record(col_����˵��).Value = Empty
                    .Record(col_����ʱ��).Value = Empty
                    .Record(col_������).Value = Empty
                    .Record(col_�������).Bold = True
                    If .Record(col_��������).Value <> "" Then
                        If curDate > CDate(.Record(col_��������).Value) Then
                            .Record(col_��������).ForeColor = vbRed
                        End If
                    End If
                    .Record(col_״̬).Value = 1
                    .Record(col_״̬).Caption = "�ȴ�����" '�ݴ����մ�����������ǵȴ�����
                    .Record(col_״̬).Icon = img16.ListImages("����״̬_�ȴ�����").Index - 1
                End If
            End If
            
            mblnOK = True
        End If
    End With
    
    If mblnOK Then
        If blnDel Then
            Call RefreshData
        Else
            rptData.Populate '������ܱ��ˣ����Բ�����redraw
        End If
        Me.stbThis.Panels(2).Text = "�����ɹ�"
    End If
    
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdPati_Click()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If cmdPati.Tag = "��" Then
        cmdPati.Tag = ""
        lvwPati.Visible = False
        If txtPati.Enabled And txtPati.Visible Then txtPati.SetFocus
    Else
        cmdPati.Tag = "��"
        If lvwPati.Tag <> "" And lvwPati.Tag <> "_0_0" Then
            lvwPati.ListItems(lvwPati.Tag).Selected = True
            lvwPati.SelectedItem.EnsureVisible
        End If
        lvwPati.Left = lngLeft + picPati.Left
        lvwPati.Top = lngTop + txtPati.Top
        lvwPati.ZOrder
        lvwPati.Visible = True
        lvwPati.SetFocus
    End If
End Sub



Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng����ID As Long, lng��ҳID As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng����ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            lng��ҳID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
            
            Call ExecFilterPati(lng����ID, lng��ҳID)
            
            txtPati.Text = lvwPati.SelectedItem.Text
            txtPati.Tag = txtPati.Text
            
            lvwPati.Tag = lvwPati.SelectedItem.Key
            cmdPati.Tag = ""
            lvwPati.Visible = False
        End If
    End If
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
    cmdPati.Tag = ""
End Sub

Private Sub txtPati_GotFocus()
    txtPati.SelStart = 0
    txtPati.SelLength = Len(txtPati.Text)
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtPati_Validate(False)
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    Dim objItem As ListItem
    Dim strInput As String, blnABC As Boolean, blnFind As Boolean
    
    strInput = UCase(txtPati.Text)
    If strInput <> lblPati.Tag Then
        blnABC = ZLCommFun.IsCharAlpha(strInput)
        
        For Each objItem In lvwPati.ListItems
            If objItem.Text <> "ȫ��" Then
                If blnABC Then
                    If objItem.SubItems(pcol_ƴ������) <> "" Then
                        If strInput Like objItem.SubItems(pcol_ƴ������) & "*" Then blnFind = True
                    End If
                Else
                    If strInput Like objItem.Text & "*" Then blnFind = True
                End If
            End If
            
            If blnFind Then
                objItem.Selected = True
                Call lvwPati_KeyPress(vbKeyReturn)
                Exit For
            End If
        Next
        If blnFind = False Then txtPati.Text = txtPati.Tag
    End If
End Sub

Private Sub ExecFilterPati(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ������˹��˷�����¼
'������lng����ID = 0 And lng��ҳID = 0ָȫ����ʾ
    
    Dim i As Long
    
    For i = 0 To rptData.Records.Count - 1
        If lng����ID = 0 And lng��ҳID = 0 Then
            If rptData.Records(i).Visible = False Then rptData.Records(i).Visible = True
        Else
            If rptData.Records(i).Item(col_����Id).Value = lng����ID And rptData.Records(i).Item(col_��ҳID).Value = lng��ҳID Then
                rptData.Records(i).Visible = True
            Else
                rptData.Records(i).Visible = False
            End If
        End If
    Next
    rptData.Populate
    If rptData.Rows.Count > 0 Then Set rptData.FocusedRow = rptData.Rows(1)
End Sub
