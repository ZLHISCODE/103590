VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDoctorShift 
   Caption         =   "ҽ�����Ӱ����"
   ClientHeight    =   11475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17685
   Icon            =   "frmDoctorShift.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11475
   ScaleWidth      =   17685
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picDataIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   0
      ScaleHeight     =   8895
      ScaleWidth      =   6015
      TabIndex        =   11
      Top             =   480
      Width           =   6015
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   2985
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   5805
         _Version        =   589884
         _ExtentX        =   10239
         _ExtentY        =   5265
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picNumBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   5745
         TabIndex        =   34
         Top             =   7560
         Width           =   5775
         Begin VB.PictureBox picNumDown 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            Picture         =   "frmDoctorShift.frx":6852
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   37
            Top             =   1080
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picNumUp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5520
            Picture         =   "frmDoctorShift.frx":7254
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picNum 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   0
            ScaleHeight     =   1095
            ScaleWidth      =   5490
            TabIndex        =   35
            Top             =   0
            Width           =   5490
            Begin VB.Label lblTypeNum 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "��������"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Top             =   120
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.Line lineNumY 
            Visible         =   0   'False
            X1              =   5520
            X2              =   5520
            Y1              =   0
            Y2              =   1320
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "��ѯ����"
         Height          =   3375
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   5820
         Begin VB.CommandButton cmdRef 
            Caption         =   "ˢ��(&R)"
            Height          =   350
            Left            =   960
            TabIndex        =   24
            Top             =   2400
            Width           =   855
         End
         Begin VB.ComboBox cboTime 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtSubject 
            Height          =   300
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdSubject 
            Caption         =   "��"
            Height          =   290
            Left            =   3000
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   1815
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   960
            ScaleHeight     =   1065
            ScaleWidth      =   4665
            TabIndex        =   15
            Top             =   1200
            Width           =   4695
            Begin VB.PictureBox picShift 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   0
               ScaleHeight     =   975
               ScaleWidth      =   4425
               TabIndex        =   18
               Top             =   0
               Width           =   4425
               Begin VB.CheckBox chkType 
                  BackColor       =   &H80000005&
                  Caption         =   "ֵ������"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   19
                  Top             =   120
                  Width           =   1335
               End
            End
            Begin VB.PictureBox picUp 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000E&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4440
               Picture         =   "frmDoctorShift.frx":7C56
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   17
               Top             =   0
               Width           =   255
            End
            Begin VB.PictureBox picDown 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4440
               Picture         =   "frmDoctorShift.frx":8658
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   16
               Top             =   840
               Width           =   255
            End
            Begin VB.Line lineY 
               X1              =   4440
               X2              =   4440
               Y1              =   0
               Y2              =   1200
            End
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   285
            Left            =   1920
            TabIndex        =   25
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CalendarTitleBackColor=   -2147483638
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   209584131
            CurrentDate     =   42675
            MaxDate         =   402133
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   285
            Left            =   3840
            TabIndex        =   26
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            CalendarTitleBackColor=   -2147483638
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   209584131
            CurrentDate     =   42702
            MaxDate         =   402133
         End
         Begin VB.Label lblSubject 
            AutoSize        =   -1  'True
            Caption         =   "ѧ    ��"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "�� ��"
            Height          =   180
            Left            =   3360
            TabIndex        =   30
            Top             =   315
            Width           =   450
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ    ��"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   795
            Width           =   720
         End
         Begin VB.Label lblSplit1 
            AutoSize        =   -1  'True
            Caption         =   "~"
            Height          =   180
            Left            =   3480
            TabIndex        =   28
            Top             =   840
            Width           =   90
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   720
         End
      End
      Begin MSComctlLib.TreeView tvwSubject 
         Height          =   1935
         Left            =   4440
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblAllNum 
         AutoSize        =   -1  'True
         Caption         =   "���˻������"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   7320
         Width           =   1080
      End
      Begin VB.Label lblRecord 
         AutoSize        =   -1  'True
         Caption         =   "���Ӱ��¼"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   900
      End
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   7800
      ScaleHeight     =   4665
      ScaleWidth      =   8385
      TabIndex        =   8
      Top             =   5280
      Width           =   8415
      Begin VB.PictureBox picShow 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   480
         ScaleHeight     =   3135
         ScaleWidth      =   6735
         TabIndex        =   9
         Top             =   600
         Width           =   6735
         Begin XtremeSuiteControls.TabControl tbcSub 
            Height          =   2580
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Width           =   3690
            _Version        =   589884
            _ExtentX        =   6509
            _ExtentY        =   4551
            _StockProps     =   64
         End
      End
   End
   Begin VB.PictureBox picPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8040
      ScaleHeight     =   3225
      ScaleWidth      =   5625
      TabIndex        =   3
      Top             =   720
      Width           =   5655
      Begin VSFlex8Ctl.VSFlexGrid vsDetail 
         Height          =   1575
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   8850
         _cx             =   15610
         _cy             =   2778
         Appearance      =   2
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   600
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDoctorShift.frx":905A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         Caption         =   "���Ӳ����嵥"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "��ɫ"
         ForeColor       =   &H0080C0FF&
         Height          =   180
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "��ʾδ���ɽ�������"
         Height          =   180
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.PictureBox picSplitY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   8280
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   4680
      Width           =   5295
   End
   Begin VB.PictureBox picSplitX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   7320
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6615
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   480
      Width           =   45
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   11115
      Width           =   17685
      _ExtentX        =   31194
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDoctorShift.frx":9240
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   28284
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   7440
      Top             =   0
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
            Picture         =   "frmDoctorShift.frx":9AD4
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":A06E
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":A608
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":10E6A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":176CC
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":180DE
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":1E940
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorShift.frx":1F352
            Key             =   "Down"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   2640
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDoctorShift.frx":1FD64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDoctorShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColData
    cold_��¼id = 0
    cold_����
    cold_����ID
    cold_����ҽ��
    cold_������
    cold_����ʱ�� '���࿪ʼʱ��|�������ʱ��
    cold_�Ӱ�ҽ��
    cold_�Ӱ���
    cold_�Ӱ�ʱ�� '�Ӱ࿪ʼʱ��|�������ʱ��
    cold_����״̬
    cold_�Ӱ�״̬
    cold_���ʱ��
    cold_������
    cold_����ʱ��
    cold_����˵��
    cold_�����ڼ�
End Enum
Private mobjESign As Object '����ǩ���ӿڲ���
Private mstrPriv As String
Private mintCA As Integer
Private mlngRow As Long
Private mrsPati As ADODB.Recordset '��¼ѡ��ʱ���ü�¼id�µ����в���
Private mobjFrom As frmShiftEdit
Private mblnEdit As Boolean
Private mobjMenu As CommandBarPopup
Private mblnLoading As Boolean
Private mlngDeptID As Long '��¼��Ա�����ٴ����ҵ�ID,��û�У�����ݽ���Ŀ���ѡ������ֵ��Ϊ���п���ʱ��Ĭ���ǿ��ҵĵ�һ��
Private mstrDeptId As String '��¼��Ա��Ȩ�޲����Ŀ���id,
Private mblnClick As Boolean '�Ƿ񴥷�cboDept��Click�¼���true ������false-������

Private Sub cboDept_Click()
        
    If cboDept.Text <> "���п���" Then
        mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    End If
    If mblnClick Then
        zlCommFun.PressKey vbKeyReturn
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call LoadType
    End If
End Sub

Private Sub cboTime_Click()
    Dim dtToday As Date
    Dim intDay As Integer
        
    dtToday = zlDatabase.Currentdate
    Select Case cboTime.Text
        Case "����"
            dtpBegin.Value = Format(Date, "yyyy-MM-dd")
            dtpEnd.Value = Format(Date, "yyyy-MM-dd")
        Case "����"
            dtpBegin.Value = Format(Date - 1, "yyyy-MM-dd")
            dtpEnd.Value = Format(Date - 1, "yyyy-MM-dd")
        Case "����"
            dtToday = Format(Date, "yyyy-MM-dd")
            intDay = Weekday(CDate(Format(Date, "yyyy-MM-dd")))
            intDay = IIf(intDay = 1, 7, intDay - 1)
            dtpBegin.Value = Format(DateAdd("d", 0 - intDay + 1, dtToday), "yyyy-MM-dd") & " 00:00:00"
            dtpEnd.Value = Format(DateAdd("d", 7 - intDay, dtToday), "yyyy-MM-dd") & " 23:59:59"
        Case "����"
            dtpBegin.Value = Format(dtToday, "yyyy-MM") & "-01 00:00:00"
            dtpEnd.Value = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(dtToday, "yyyy-MM") & "-01"))), "yyyy-MM-dd") & " 23:59:59"
        Case "�Զ���"
            
    End Select
    If cboTime.Text = "�Զ���" Then
        dtpBegin.Enabled = True
        dtpEnd.Enabled = True
    Else
        dtpBegin.Enabled = False
        dtpEnd.Enabled = False
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Long, j As Long, lngId As Long
    Dim strDept As String, strOutTime As String, strInTime As String
    Dim strOutPer As String, strInPer As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    If rptData.SelectedRows.Count > 0 Then
        If rptData.SelectedRows(0).GroupRow = False Then
            lngId = rptData.SelectedRows(0).Record(cold_��¼id).Value
            strOutPer = rptData.SelectedRows(0).Record(cold_����ҽ��).Value
            strInPer = rptData.SelectedRows(0).Record(cold_�Ӱ�ҽ��).Value
        End If
    End If
    Select Case Control.id
    Case conMenu_File_TypeManage '��ι���
        If frmShiftMange.ShowMe(mstrDeptId, mlngDeptID) Then
            Call LoadType
        End If
    Case conMenu_File_Preview 'Ԥ��
        ReportMode 1
    Case conMenu_File_Print '��ӡ
        ReportMode 2
    Case conMenu_File_Excel '�����Excel
        ReportMode 3
    Case conMenu_Edit_NewItem '����
        If cboDept.List(0) = "���п���" Then j = 1
        For i = j To cboDept.ListCount - 1
            strDept = IIf(strDept = "", "", strDept & "|") & cboDept.List(i) & "," & cboDept.ItemData(i)
        Next
        strDept = cboDept.ListIndex - j & "|" & strDept
        If frmEdit.ShowMe(0, 0, strDept, grsUserInfo!����) Then Call RefreshRecord
    Case conMenu_Edit_Modify '�޸�
        strDept = rptData.SelectedRows(0).Record(cold_����).Value & "|" & rptData.SelectedRows(0).Record(cold_����ID).Value
        strOutTime = rptData.SelectedRows(0).Record(cold_������).Value & "|" & rptData.SelectedRows(0).Record(cold_����ʱ��).Value
        strInTime = rptData.SelectedRows(0).Record(cold_�Ӱ���).Value & "|" & rptData.SelectedRows(0).Record(cold_�Ӱ�ʱ��).Value
        If frmEdit.ShowMe(1, lngId, strDept, strOutPer, strOutTime, strInPer, strInTime) Then Call RefreshRecord
    Case conMenu_Edit_Delete 'ɾ��
        If MsgBox("��ȷ��ɾ������ֵ���¼��", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
        gstrSQL = "Zl_ҽ�����Ӱ��¼_State(0," & lngId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ��ֵ���¼")
        Call RefreshRecord
    Case conMenu_Edit_FinOut '��ɽ���
        Set rsTemp = GetUserInfo(strOutPer)
        If strOutPer = "����Ա" Then
            strTemp = "ZLHIS"
        Else
            strTemp = rsTemp!�û���
        End If
        Call ShiftMange(0, lngId, strTemp, strOutPer, rsTemp!����ID)
    Case conMenu_Edit_FinIn '��ɽӰ�
        Set rsTemp = GetUserInfo(strInPer)
        Call ShiftMange(1, lngId, rsTemp!�û���, strInPer, rsTemp!����ID)
    Case conMenu_Edit_FinRead '�������
        strTemp = frmReview.ShowMe
        If strTemp = "ȡ��JM" Then Exit Sub
        gstrSQL = "Zl_ҽ�����Ӱ��¼_State(3," & lngId & ",'" & grsUserInfo!���� & "','" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
        Call RefreshRecord
    Case conMenu_Edit_CancelOut, conMenu_Edit_CancelIn, conMenu_Edit_CancelRead 'ȡ����ɽ���,ȡ����ɽӰ�,ȡ���������
        Call CancelOper(Control.id, lngId)
    Case conMenu_Edit_CheckOutSign '��֤ǩ��
        Call VerifySign(1, lngId)
    Case conMenu_Edit_CheckInSign
        Call VerifySign(2, lngId)
    Case conMenu_Report_Record '���������ѯ����
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1242_2", Me)
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Control.Checked = Not Control.Checked
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Call Form_Resize
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        Control.Checked = Not Control.Checked
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Not Control.Checked
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Not Control.Checked
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Call Form_Resize
        Me.cbsMain.RecalcLayout
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    End Select
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
Private Sub ReportMode(bytMode As Byte)
'bytMode-1Ԥ����2��ӡ��3�����excel
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1242_1", Me, "��¼id=" & Val(rptData.SelectedRows(0).Record(cold_��¼id).Value), bytMode)
End Sub

Private Sub VerifySign(bytType As Byte, ByVal lngId As Long)
'��֤ǩ��
'bytType��1-���ࣻ2-�Ӱ�
    Dim rsTemp As ADODB.Recordset
    Dim strSource As String
        
    On Error GoTo errH
    If lngId = 0 Then Exit Sub
    gstrSQL = "Select ֤��id From ҽ�����Ӱ�ǩ�� Where ��¼id = [1] And ǩ������ =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId, bytType)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ǩ�������ѱ�ɾ��������ǩ����֤ʧ�ܣ�", vbInformation, Me.Caption
        Exit Sub
    End If
    Call ReadSignSource(lngId, strSource)
    If mobjESign.VerifySignature(strSource, rsTemp!֤��ID, 0) = True Then
        MsgBox "����ǩ����֤�ɹ�!", vbInformation, Me.Caption
    Else
        MsgBox "����ǩ����֤ʧ��!", vbInformation, Me.Caption
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub CancelOper(ByVal lngType As Long, ByVal lngId As Long)
'ȡ����ɽ��ࡢȡ����ɽӰࡢȡ���������
    Dim bytType As Byte
    
    On Error GoTo errH
    bytType = Decode(lngType, conMenu_Edit_CancelOut, 4, conMenu_Edit_CancelIn, 5, conMenu_Edit_CancelRead, 6)
    gstrSQL = "Zl_ҽ�����Ӱ��¼_State(" & bytType & "," & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ�����")
    Call RefreshRecord
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function CheckPati() As Boolean
'����ǰ�Ĳ�����Ϣ��飬��Ҫ��齻�������Ƿ�Ϊ��
    Dim i As Long

    With vsDetail
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("��������")) = "" Then
                MsgBox "���ڽ�������Ϊ�յĲ��ˣ��޷���ɽ��࣬���飡", vbInformation, Me.Caption
                Call .ShowCell(i, .ColIndex("��������"))
                Exit Function
            End If
        Next
    End With
    CheckPati = True
End Function

Private Sub ShiftMange(ByVal bytType As Byte, ByVal lngId As Long, ByVal strPer As String, ByVal strDoc As String, ByVal lngDeptID As Long)
'��ɽ����Ӱ�
'bytType:0-���ࣻ1-�ӰࣻstrPer-�û�����strDoc-����
    Dim lng֤��ID As Long, lngCA As Long
    Dim strSource As String, strTimeStamp As String, strTimeStampCode As String
    Dim strSign As String, strCaInfo As String
    Dim blnBegin As Boolean, blnIndetifi As Boolean '�Ƿ��Ѿ���֤
    Dim rsTemp As ADODB.Recordset

    If bytType = 0 Then
        If Not CheckPati Then Exit Sub
    End If
    If strPer <> grsUserInfo!�û��� Then
        '�����Ӱ�ʱ�����ǰ�û����Ƕ�Ӧ�Ľ����Ӱ�ҽ���û������������֤
        If Not frmUserIdentify.ShowMe(Me, "�����֤������������", glngSys, strPer, True) Then
'            MsgBox "�����֤δͨ�����޷����" & IIf(bytType = 0, "���࣡", "�Ӱ࣡"), vbInformation, Me.Caption
            Exit Sub
        Else
            blnIndetifi = True
        End If
    End If
    If GetCA(strDoc) Then
        On Error Resume Next
        If mobjESign Is Nothing Then
            Set mobjESign = CreateObject("zl9ESign.clsESign")
        End If
        Err.Clear: On Error GoTo errH
        If Not mobjESign Is Nothing Then
            Call mobjESign.Initialize(gcnOracle, glngSys)
        Else
            MsgBox "����ǩ������δ����ȷ��װ����˲������ܼ�����", vbInformation, Me.Caption
            Exit Sub
        End If
        Call ReadSignSource(lngId, strSource)
        strSign = mobjESign.Signature(strSource, strPer, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = zlStr.To_Date(strTimeStamp)
            Else
                strTimeStamp = "NULL"
            End If
        Else
'            MsgBox "����ǩ��ʧ�ܣ��޷����" & IIf(bytType = 0, "���࣡", "�Ӱ࣡"), vbInformation, Me.Caption
            Exit Sub
        End If
        strCaInfo = "," & lng֤��ID & ",'" & strSign & "','" & strTimeStampCode & "'," & strTimeStamp
    Else
        If Not blnIndetifi Then
            If Not frmUserIdentify.ShowMe(Me, "�����֤������������", glngSys, strPer, True) Then
'                MsgBox "�����֤δͨ�����޷����" & IIf(bytType = 0, "���࣡", "�Ӱ࣡"), vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    gcnOracle.BeginTrans: blnBegin = True
    gstrSQL = "Zl_ҽ�����Ӱ�ǩ��_Edit(" & lngId & "," & IIf(bytType = 0, 1, 2) & ",'" & strDoc & "'" & strCaInfo & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽ������ǩ��")
    gstrSQL = "Zl_ҽ�����Ӱ��¼_State(" & IIf(bytType = 0, 1, 2) & "," & lngId & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɽ���")
    blnBegin = True
    gcnOracle.CommitTrans
    
    Call RefreshRecord
    Exit Sub
errH:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngIn As Long, lngHold As Long
    Dim strReadPer As String
    
    If mblnEdit Then
        If Control.id = conMenu_Edit_NewItem Or Control.id = conMenu_Edit_Modify Or Control.id = conMenu_Edit_Delete Or Control.id = conMenu_Edit_FinOut Or Control.id = conMenu_Edit_FinIn Or Control.id = conMenu_Edit_FinRead _
            Or Control.id = conMenu_Edit_CancelOut Or Control.id = conMenu_Edit_CancelIn Or Control.id = conMenu_Edit_CancelRead Or Control.id = conMenu_Edit_CheckOutSign Or Control.id = conMenu_Edit_CheckInSign Or Control.id = conMenu_File_TypeManage _
            Or Control.id = conMenu_File_Preview Or Control.id = conMenu_File_Print Then
            Control.Enabled = False
        Else
            Control.Enabled = True
        End If
        Exit Sub
    End If
    
    lngIn = -1
    lngHold = -1
    strReadPer = "-1"
    If rptData.SelectedRows.Count > 0 Then
        If Not rptData.SelectedRows(0).GroupRow Then
            '��ʾѡ����ʱ������
            lngIn = Val(rptData.SelectedRows(0).Record(cold_����״̬).Value)
            lngHold = Val(rptData.SelectedRows(0).Record(cold_�Ӱ�״̬).Value)
            strReadPer = rptData.SelectedRows(0).Record(cold_������).Value
        End If
    End If
    Select Case Control.id
        'Ȩ������
        Case conMenu_Report_Record
            Control.Enabled = CheckPriv("�ٴ����ҽ��Ӱ������ѯ")
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = CheckPriv("���Ӱ��¼") And lngHold = 1
        Case conMenu_File_TypeManage
            Control.Enabled = CheckPriv("��ι���")
        Case conMenu_Edit_NewItem
            Control.Enabled = CheckPriv("ҽ�����Ӱ�")
        Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_FinOut '�޸�,ɾ������ɽ���
            Control.Enabled = lngIn = 0 And CheckPriv("ҽ�����Ӱ�")
        Case conMenu_Edit_FinIn '��ɽӰ�
            Control.Enabled = False
            Control.Enabled = lngHold = 0 And CheckPriv("ҽ�����Ӱ�") And lngIn = 1
        Case conMenu_Edit_FinRead '�������
            Control.Enabled = strReadPer = "" And CheckPriv("���Ӱ�����") _
                And lngHold = 1 And lngIn = 1
        Case conMenu_Edit_CancelOut 'ȡ����ɽ���
            Control.Enabled = lngIn = 1 And lngHold = 0 And CheckPriv("ҽ�����Ӱ�")
        Case conMenu_Edit_CancelIn 'ȡ����ɽӰ�
            Control.Enabled = lngHold = 1 And strReadPer = "" And CheckPriv("ҽ�����Ӱ�")
        Case conMenu_Edit_CancelRead 'ȡ���������
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("���Ӱ�����")
        Case conMenu_Edit_CheckOutSign '����ҽ������ǩ����֤
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("ҽ�����Ӱ�")
        Case conMenu_Edit_CheckInSign '�Ӱ�ҽ������ǩ����֤
            Control.Enabled = strReadPer <> "" And strReadPer <> "-1" And CheckPriv("ҽ�����Ӱ�")
    End Select
End Sub

Private Sub cmdRef_Click()
    Call RefreshRecord
End Sub

Private Sub cmdSubject_Click()
    
    tvwSubject.Visible = True
    tvwSubject.SetFocus
End Sub

Private Sub Form_Activate()
    With vsDetail
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .AutoSize .ColIndex("��������")
    End With
End Sub

Private Sub Form_Load()
    
    mblnLoading = True
    mstrPriv = gstrPrivs
    Set grsUserInfo = zlDatabase.GetUserInfo
    mintCA = IIf(GetCA(grsUserInfo!����), 1, 0)
    Call InitCommandBar
    Call InitReportColumn
    
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "�Զ���"
    cboTime.ListIndex = 0
        
    Call LoadShowData
    Call LoadType
    Call RefreshRecord
    picSplitX.BackColor = Me.BackColor
    picSplitY.BackColor = Me.BackColor
    picSplitX.Left = 5900
    picSplitY.Top = 4200

    Call GetFrom
    Call RestoreWinState(Me, App.ProductName)
    Call LoadVsfColWidth
    
    cmdSubject.Enabled = CheckPriv("����ѧ��")
    
    mblnLoading = False
    If mlngRow = 0 And rptData.Rows.Count > 1 Then
        Set rptData.FocusedRow = rptData.Rows(1)
    End If
    
End Sub

Private Sub LoadVsfColWidth()
'vsf�����һ�ε��п�
    Dim strCols As String
    Dim varTemp As Variant, varData As Variant
    Dim i As Long
    
    strCols = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbaUser & "\" & gstrProductName & "\ҽ�����Ӱ��¼", "�����嵥�п�")
    If strCols = "" Or InStr(strCols, "����") > 0 Then
        strCols = "���|0;����ID|0;��ҳID|0;����ID|0;����|300;����|300;����|1500;����|900;�Ա�|480;����|720;����|495;��ʶ��|1020;��Ժʱ��|990;��Ժ��ʽ|900"
    End If
    varTemp = Split(strCols, ";")
    With vsDetail
        For i = LBound(varTemp) To UBound(varTemp)
            varData = Split(varTemp(i), "|")
            If varData(0) = .ColKey(i) Then
                .ColWidth(i) = varData(1)
            End If
        Next
    End With
End Sub

Private Sub LoadShowData()
'���ؽ���ѧ�ƺͿ��ҵ�����
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strDept As String, strTemp As String
    Dim objNode As Object
    Dim varTemp As Variant
        
    On Error GoTo errH
    gstrSQL = "Select b.�û���, d.Id ����id, d.���� ����, f.���� ѧ�Ʊ���, f.���� ѧ��" & vbNewLine & _
        "From �ϻ���Ա�� b, ���ű� d, ������Ա e, �ٴ����� f, �ٴ����� g" & vbNewLine & _
        "Where b.�û��� =[1] And e.��Աid = b.��Աid And e.����id = d.Id And e.ȱʡ = 1 And d.Id = g.����id And g.�������� = f.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrDbaUser)
    If rsTemp.RecordCount = 1 Then
        txtSubject.Text = rsTemp!ѧ��
        txtSubject.Tag = rsTemp!ѧ�Ʊ���
        cboDept.Tag = rsTemp!����ID
        strDept = rsTemp!����
        mlngDeptID = rsTemp!����ID
    Else
        txtSubject.Tag = "����ѧ��"
        txtSubject.Text = "����ѧ��"
        strDept = "���п���"
    End If
    If cmdSubject.Enabled Then
        'ֻ�������ٴ����ҵ�ѧ��
        gstrSQL = "Select ����, ���� From �ٴ�����" & vbNewLine & _
            "Where ���� In (Select Distinct a.�������� From �ٴ����� a, ��������˵�� c Where a.����id = c.����id" & vbNewLine & _
            "And c.��������='�ٴ�' And c.������� In (2, 3)) Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With tvwSubject
            .Left = txtSubject.Left + 40
            .Top = txtSubject.Top + txtSubject.Height + 110
            .Width = txtSubject.Width
            .Height = fraFilter.Height
            .Nodes.Clear
            Set objNode = .Nodes.Add(, , "K����ѧ��", "����ѧ��", "Dept")
            Do Until rsTemp.EOF
                varTemp = Split(rsTemp!����, ".")
                For i = LBound(varTemp) To UBound(varTemp) - 1
                    strTemp = IIf(strTemp = "", "", strTemp & ".") & varTemp(i)
                Next
                On Error Resume Next
                Set objNode = .Nodes.Add("K" & strTemp & "", tvwChild, "K" & CStr(rsTemp!����), rsTemp!����, "Dept")
                If Err.Number <> 0 Then
                    Err.Clear: On Error GoTo errH
                    Set objNode = .Nodes.Add(, , "K" & CStr(rsTemp!����), rsTemp!����, "Dept")
                End If
                rsTemp.MoveNext
            Loop
            .ZOrder 0
        End With
    End If
    On Error GoTo errH
    Call LoadDept(txtSubject.Tag)
    
    mblnClick = False
    For i = 0 To cboDept.ListCount - 1
        If cboDept.List(i) = strDept Then
            cboDept.ListIndex = i
            Exit For
        Else
            If i = cboDept.ListCount - 1 Then
                mlngDeptID = cboDept.ItemData(0)
            End If
        End If
    Next
    mblnClick = True
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadDept(ByVal strSubjectCode As String)
'ѡ��ѧ��ʱ��Ӧ�Ĳ���ѡ��仯
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    cboDept.Clear
    mstrDeptId = ""
    If strSubjectCode = "����ѧ��" Then
        gstrSQL = "Select distinct ���� ||'-' || ���� as ����,Id,���� From ���ű�" & vbNewLine & _
            "Where Id In (Select Distinct a.����id From �ٴ����� a, ��������˵�� c Where a.����id = c.����id" & vbNewLine & _
            "And c.������� In (2, 3) And c.��������='�ٴ�') And (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Else
        gstrSQL = "Select distinct b.���� || '-' || b.���� as ����, b.Id,b.���� From �ٴ����� a, ���ű� b,��������˵�� c" & vbNewLine & _
            "Where a.����id = b.Id And a.����id = c.����id And c.������� In (2, 3) And  a.�������� =[1]" & vbNewLine & _
            "And c.��������='�ٴ�'And (b.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or b.����ʱ�� is NULL) Order By ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strSubjectCode)
    If rsTemp.RecordCount > 1 Then
        cboDept.AddItem "���п���"
    End If
    Call zlcontrol.CboAddData(cboDept, rsTemp, False)
    mblnClick = False
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    mblnClick = True
    Do While Not rsTemp.EOF
        mstrDeptId = mstrDeptId & "," & rsTemp!id
        rsTemp.MoveNext
    Loop
    mstrDeptId = Mid(mstrDeptId, 2)
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub RefreshRecord()
'ˢ�¼�¼����
    Dim i As Long
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
        
    vsDetail.Rows = 1
    rptData.Records.DeleteAll
    lblTypeNum(0).Visible = False
    lblAllNum.Caption = "���˻������"
    For i = 1 To lblTypeNum.UBound
        Unload lblTypeNum(i)
    Next
    gstrSQL = ""
    If cboDept.Text = "���п���" Then
        gstrSQL = " And b.id in(Select ����id From �ٴ�����" & IIf(txtSubject.Text = "����ѧ��", ")", " Where �������� =[4])")
    Else
        gstrSQL = " And b.id=[4]"
    End If
    strTemp = ""
    For i = chkType.LBound To chkType.UBound
        If chkType(i).Value = 1 Then
            strTemp = IIf(strTemp = "", "", strTemp & ",") & chkType(i).Caption
        End If
    Next
    gstrSQL = gstrSQL & " order by a.��¼id"
    On Error GoTo errH
    gstrSQL = "Select a.��¼id, a.����ID,b.���� ����, a.����ҽ��, a.������," & vbNewLine & _
        "       To_Char(a.���࿪ʼʱ��, 'MM-DD HH24:Mi') || '��' || To_Char(a.�������ʱ��, 'MM-DD HH24:Mi') As �����ڼ�, a.���࿪ʼʱ��, a.�������ʱ��," & vbNewLine & _
        "       a.�Ӱ�ҽ��, a.�Ӱ���, a.�Ӱ࿪ʼʱ��, a.�Ӱ����ʱ��, a.����״̬, a.�Ӱ�״̬, a.���ʱ��, a.������, a.����ʱ��, a.����˵��" & vbNewLine & _
        "From ҽ�����Ӱ��¼ a, ���ű� b" & vbNewLine & _
        "Where a.����id = b.Id And a.�Ӱ࿪ʼʱ�� >=to_date([1],'yyyy-mm-dd hh24:mi:ss') And a.�Ӱ࿪ʼʱ��<to_date([2],'yyyy-mm-dd hh24:mi:ss')" & _
        " And a.������ in(Select * From Table(f_str2list([3]))) " & gstrSQL

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59"), _
                strTemp, IIf(cboDept.Text = "���п���", Trim(txtSubject.Tag), cboDept.ItemData(cboDept.ListIndex)))
    Do While Not rsTemp.EOF
        Set objRecord = rptData.Records.Add
        Set objItem = objRecord.AddItem(Val(rsTemp!��¼id))
        Set objItem = objRecord.AddItem(CStr(rsTemp!����))
        Set objItem = objRecord.AddItem(CStr(rsTemp!����ID))
        Set objItem = objRecord.AddItem(CStr(rsTemp!����ҽ��))
        Set objItem = objRecord.AddItem(CStr(rsTemp!������))
        Set objItem = objRecord.AddItem(CStr(rsTemp!���࿪ʼʱ��) & "|" & CStr(rsTemp!�������ʱ��))
        Set objItem = objRecord.AddItem(CStr(rsTemp!�Ӱ�ҽ��))
        Set objItem = objRecord.AddItem(CStr(rsTemp!�Ӱ���))
        Set objItem = objRecord.AddItem(CStr(rsTemp!�Ӱ࿪ʼʱ��) & "|" & CStr(rsTemp!�Ӱ����ʱ��))
        Set objItem = objRecord.AddItem(CStr(rsTemp!����״̬ & ""))
        Set objItem = objRecord.AddItem(CStr(rsTemp!�Ӱ�״̬ & ""))
        If IsNull(rsTemp!���ʱ��) Then
            Set objItem = objRecord.AddItem(" ")
        Else
            Set objItem = objRecord.AddItem(Format(rsTemp!���ʱ��, "yyyy-mm-dd") & "")
        End If
        Set objItem = objRecord.AddItem(CStr(rsTemp!������ & ""))
        Set objItem = objRecord.AddItem(Format(rsTemp!����ʱ��, "yyyy-mm-dd") & "")
        Set objItem = objRecord.AddItem(CStr(rsTemp!����˵�� & ""))
        Set objItem = objRecord.AddItem(CStr(rsTemp!�����ڼ�))
        If rsTemp!������ & "" = "" Then
            objRecord.PreviewText = "��δ����"
        Else
            objRecord.PreviewText = "������:" & rsTemp!������ & "" & "  ����ʱ��:" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & "" & "  ����˵��:" & rsTemp!����˵�� & ""
        End If
        rsTemp.MoveNext
    Loop
    rptData.Populate
    If mlngRow > 0 And mlngRow <= rptData.Rows.Count - 1 Then
        Set rptData.FocusedRow = rptData.Rows(mlngRow)
        rptData.SetFocus
        Exit Sub
    End If
    If mlngRow = 0 And rptData.Rows.Count > 1 Then
        Set rptData.FocusedRow = rptData.Rows(1)
'        rptData.SetFocus
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadType()
'��̬����ֵ����
    Dim lngIndex As Long, lngNum As Long, lngHeight As Long
    Dim objChk As Object
    Dim lngMax As Long, lngMaxNum As Long
    Dim rsTemp As ADODB.Recordset
        
    Set rsTemp = GetShiftType(2, IIf(cboDept.Text = "���п���", mstrDeptId, mlngDeptID))
    For lngIndex = 1 To chkType.UBound
        Unload chkType(lngIndex)
    Next
    chkType(0).Visible = False
    lngIndex = 0
    lngMax = rsTemp.RecordCount - 1
    If rsTemp.RecordCount > 1 Then rsTemp.MoveFirst
    For lngIndex = 0 To lngMax
        If lngIndex = 0 Then
            chkType(0).Visible = True
            chkType(0).Value = 1
            chkType(0).Caption = rsTemp!�������
            chkType(0).Width = 1300
        Else
            Load chkType(lngIndex)
            chkType(lngIndex).Caption = rsTemp!�������
            Set objChk = chkType(lngIndex)
            Set chkType(lngIndex).Container = picShift
            lngNum = Fix(lngIndex / 3)
            If lngNum = lngIndex / 3 Then
                chkType(lngIndex).Move chkType(0).Left, chkType(0).Top + (chkType(0).Height + 120) * lngNum, 1300, chkType(0).Height
            Else
                chkType(lngIndex).Move chkType(lngIndex - 1).Left + chkType(lngIndex - 1).Width + 50, chkType(lngIndex - 1).Top, 1300, chkType(0).Height
            End If
            chkType(lngIndex).Visible = True
            chkType(lngIndex).Value = 1
        End If
        rsTemp.MoveNext
    Next
    lngMaxNum = Fix(lngMax / 3) + 1
    picBack.Height = IIf(lngMaxNum > 3, 3, lngMaxNum) * (chkType(0).Height + 120) + 120
    lineY.X1 = picShift.Width
    lineY.X2 = picShift.Width
    lineY.Y1 = 0
    lineY.Y2 = picBack.Height
    lngHeight = lngMaxNum * (chkType(0).Height + 120) + 120
    If lngHeight <= picBack.Height Then lngHeight = picBack.Height
    picShift.Height = lngHeight
    If Fix(lngMax / 3) > 2 Then
        lineY.Visible = True
        picUp.Visible = False
        picDown.Visible = True
        picUp.Top = 0
        picDown.Top = picBack.Height - picDown.Height
    Else
        picBack.BackColor = picShift.BackColor
        lineY.Visible = False
        picUp.Visible = False
        picDown.Visible = False
    End If
    cmdRef.Top = picBack.Top + picBack.Height + 100
    fraFilter.Height = cmdRef.Top + cmdRef.Height + 100
    Call picDataIn_Resize
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_TypeManage, "��ι���(&T)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&C)"): objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mobjMenu.id = conMenu_EditPopup
    With mobjMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinOut, "��ɽ���(&O)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelOut, "ȡ����ɽ���(&J)")
        If mintCA > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckOutSign, "����ǩ����֤(&C)")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinIn, "��ɽӰ�(&I)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelIn, "ȡ����ɽӰ�(&H)")
        If mintCA > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_CheckInSign, "�Ӱ�ǩ����֤(&S)")
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinRead, "�������(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_CancelRead, "ȡ���������(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "����(&R)", -1, False)
    objMenu.id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Report_Record, "�ٴ����ҽ��Ӱ������ѯ(&S)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            objControl.Checked = True
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        objControl.Checked = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
'            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinOut, "��ɽ���"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinIn, "��ɽӰ�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_FinRead, "�������")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    '����һЩ�����Ĳ���������
'    With cbsMain.Options
'        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
'    End With
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPriv, "ZL1_INSIDE_1242_1")
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    On Error Resume Next
    
    If Not cbsMain(2).Visible Then
        lngTop = 500
    End If
    If stbThis.Visible Then
        lngHeight = stbThis.Height
    End If
    picSplitX.Top = 1000 - picSplitY - lngTop
    picSplitX.Height = Me.ScaleHeight - 1000 - lngHeight + lngTop
    
    picDataIn.Move 0, 900 - lngTop, picSplitX.Left, picSplitX.Height
    
    picSplitY.Left = picSplitX.Left + picSplitX.Width
    picSplitY.Width = Me.ScaleWidth - picSplitY.Left
    
    picPatient.Move picSplitY.Left, 1120 - lngTop, picSplitY.Width, picSplitY.Top - 1120 + lngTop
    picSub.Move picPatient.Left, picSplitY.Top + picSplitY.Height, picPatient.Width, picSplitX.Height - picSplitY.Top + 1000 - lngTop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strCols As String
    Dim i As Long

    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjFrom Is Nothing Then
        Unload mobjFrom
        Set mobjFrom = Nothing
    End If
    Set mrsPati = Nothing
    Set mobjESign = Nothing
    Set mobjMenu = Nothing
    Set grsUserInfo = Nothing
    mlngRow = 0
    With rptData
        For i = cold_��¼id To cold_�����ڼ�
            strCols = strCols & ";" & .Columns(i).Caption & "|" & .Columns(i).Width
        Next
    End With
    strCols = Mid(strCols, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbaUser & "\" & gstrProductName & "\ҽ�����Ӱ��¼", "��¼�п�", strCols
    
    strCols = ""
    With vsDetail
        For i = .ColIndex("���") To .ColIndex("���")
            strCols = strCols & ";" & .ColKey(i) & "|" & .ColWidth(i)
        Next
    End With
    strCols = Mid(strCols, 2)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbaUser & "\" & gstrProductName & "\ҽ�����Ӱ��¼", "�����嵥�п�", strCols
End Sub

Private Sub picDataIn_Resize()
    
    On Error Resume Next
    If picDataIn.Width < 3000 Then Exit Sub
    If picDataIn.Height < 3000 Then Exit Sub
    fraFilter.Move 50, 150
    lblRecord.Move fraFilter.Left, fraFilter.Top + fraFilter.Height + 100
    
    picNumBack.Move fraFilter.Left, picDataIn.Height - picNumBack.Height - lblAllNum.Height
    lblAllNum.Move picNumBack.Left, picNumBack.Top - lblAllNum.Height - 50
    
    rptData.Move fraFilter.Left, lblRecord.Top + lblRecord.Height + 20, picDataIn.Width - 40, lblAllNum.Top - rptData.Top - 150
    
End Sub

Private Sub InitReportColumn()
'���ܣ���ʼ�������б���
    Dim objCol As ReportColumn
    Dim strRptCol As String '�ؼ����п���ʽ:����,����;����,����...
    Dim varData As Variant
    Dim i As Long

    strRptCol = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbaUser & "\" & gstrProductName & "\ҽ�����Ӱ��¼", "��¼�п�")
    If strRptCol = "" Then
        strRptCol = "��¼id|0;����|0;����ID|0;����ҽ��|55;������|55;����ʱ��|0;�Ӱ�ҽ��|55;�Ӱ���|55;����ʱ��|0;����״̬|0;" & _
            "�Ӱ�״̬|0;���ʱ��|120;������|55;����ʱ��|120;����˵��|200;�����ڼ�|160"
    End If
    With rptData
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(cold_��¼id, "��¼id", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_����, "����", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_����ҽ��, "����ҽ��", 0, False)
        Set objCol = .Columns.Add(cold_������, "������", 0, False)
        Set objCol = .Columns.Add(cold_����ʱ��, "����ʱ��", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_�Ӱ�ҽ��, "�Ӱ�ҽ��", 0, True)
        Set objCol = .Columns.Add(cold_�Ӱ���, "�Ӱ���", 0, True)
        Set objCol = .Columns.Add(cold_����ʱ��, "����ʱ��", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_����״̬, "����״̬", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_�Ӱ�״̬, "�Ӱ�״̬", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(cold_���ʱ��, "���ʱ��", 0, True)
        Set objCol = .Columns.Add(cold_������, "������", 0, True)
        Set objCol = .Columns.Add(cold_����ʱ��, "����ʱ��", 0, True)
        Set objCol = .Columns.Add(cold_����˵��, "����˵��", 0, True)
        Set objCol = .Columns.Add(cold_�����ڼ�, "�����ڼ�", 0, False)
        varData = Split(strRptCol, ";")
        For i = cold_��¼id To cold_�����ڼ�
            .Columns(i).Width = Split(varData(i), "|")(1)
        Next
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = cold_����
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�ļ�¼..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        
        .GroupsOrder.Add .Columns(cold_����)
        .GroupsOrder(0).SortAscending = True
    End With
End Sub

Private Sub picnumDown_Click()

    picNum.Top = picNum.Top - (lblTypeNum(0).Height + 120) * 3
    If picNum.Top + picNum.Height > picBack.Height Then
        picNumDown.Visible = True
    Else
        picNumDown.Visible = False
    End If
    If picNum.Top < 0 Then
        picNumUp.Visible = True
    Else
        picNumUp.Visible = False
    End If
End Sub

Private Sub picnumup_Click()
    picNum.Top = picNum.Top + (lblTypeNum(0).Height + 120) * 3
    If picNum.Top > 0 Then picNum.Top = 0
    If picNum.Top + picNum.Height > picBack.Height Then
        picNumDown.Visible = True
    Else
        picNumDown.Visible = False
    End If
    If picNum.Top < 0 Then
        picNumUp.Visible = True
    Else
        picNumUp.Visible = False
    End If
End Sub

Private Sub picPatient_Resize()
    lblList.Move 50, 50
    vsDetail.Move 0, 250, picPatient.Width, picPatient.Height - lblList.Height - lblList.Top - 40
    lblColor2.Move vsDetail.Left + vsDetail.Width - lblColor2.Width - 300, lblList.Top
    lblColor.Move lblColor2.Left - lblColor.Width - 50, lblColor2.Top
End Sub


Private Sub picShow_Resize()
    tbcSub.Top = 0: tbcSub.Left = 0
    tbcSub.Width = picShow.Width: tbcSub.Height = picShow.Height
End Sub

Private Sub picSplitX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitX.Tag <> "Draging" Then
        picSplitX.Tag = "Draging"
        picSplitX.BackColor = 0
    End If
    
    sglNew = picSplitX.Left + X
    
    picSplitX.Left = sglNew
End Sub

Private Sub picSplitX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitX.Tag = "Draging" Then
        picPatient.Width = Me.ScaleWidth - picSplitX.Left
        Call Form_Resize
        picSplitX.BackColor = Me.BackColor
        picSplitX.Tag = ""
    End If
End Sub

Private Sub picSplitY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglNew As Single
    
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    If picSplitY.Tag <> "Draging" Then
        picSplitY.Tag = "Draging"
        picSplitY.BackColor = 0
    End If
    
    sglNew = picSplitY.Top + Y
    
    picSplitY.Top = sglNew
End Sub

Private Sub picSplitY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    If picSplitY.Tag = "Draging" Then
        picPatient.Height = Me.ScaleHeight - picSplitY.Top
        Call Form_Resize
        picSplitY.BackColor = Me.BackColor
        picSplitY.Tag = ""
    End If
End Sub

Private Sub picSub_Resize()
    '����tab��ǩ
    picShow.Top = -360: picShow.Left = 0
    picShow.Width = picSub.Width: picShow.Height = picSub.Height + 360
    
    tbcSub.Top = 0: tbcSub.Left = 0
    tbcSub.Width = picShow.Width: tbcSub.Height = picShow.Height
End Sub

Private Sub picDown_Click()

    picShift.Top = picShift.Top - (chkType(0).Height + 120) * 2
    If picShift.Top + picShift.Height > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picShift.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub picUp_Click()
    picShift.Top = picShift.Top + (chkType(0).Height + 120) * 2
    If picShift.Top > 0 Then picShift.Top = 0
    If picShift.Top + picShift.Height > picBack.Height Then
        picDown.Visible = True
    Else
        picDown.Visible = False
    End If
    If picShift.Top < 0 Then
        picUp.Visible = True
    Else
        picUp.Visible = False
    End If
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then
        Call ShowContenMenu.ShowPopup
    End If
End Sub
Private Function ShowContenMenu() As CommandBar
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    Dim cbrControl3 As CommandBarControl
    
    '�����˵�����
    On Error GoTo ErrHand
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In mobjMenu.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.id, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.Visible = cbrControl.Visible
        cbrPopupItem.IconId = cbrControl.IconId
        cbrPopupItem.Checked = cbrControl.Checked
        cbrPopupItem.Style = cbrControl.Style
        If cbrControl.Type = xtpControlPopup Or cbrControl.Type = xtpControlSplitButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrControl3 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.id, cbrControl2.Caption)
                cbrControl3.BeginGroup = cbrControl2.BeginGroup
                cbrControl3.Parameter = cbrControl2.Parameter
                cbrControl3.Visible = cbrControl2.Visible
                cbrControl3.Style = cbrControl2.Style
            Next
        End If
    Next
    Set ShowContenMenu = cbrPopupBar
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub rptData_SelectionChanged()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long, i As Long
    Dim dt��ʼʱ�� As Date, dt����ʱ�� As Date
    Dim btnNoEdit As Boolean
    
    If mblnLoading Then Exit Sub
    If rptData.Records.Count < 1 Then Exit Sub
    vsDetail.Rows = 1
    
    If rptData.SelectedRows(0).GroupRow Then
        Call mobjFrom.zlRefresh(0, 0, 0, 0, 0, 0, 0, False, "") '��ձ༭�б�
        Exit Sub
    End If
    mlngRow = rptData.FocusedRow.Index
    lngId = Val(rptData.SelectedRows(0).Record(cold_��¼id).Value)
    On Error GoTo errH
    gstrSQL = "Select a.���,a.����id,a.����ID,a.��ҳID,a.�������� as ����, a.����," & vbNewLine & _
        "a.�Ա�, a.����, a.����, a.��ʶ��, to_char(a.��Ժʱ��,'yyyy-mm-dd') ��Ժʱ��, a.��Ժ��ʽ, a.�������� From ҽ�����Ӱ����� a Where ��¼id =[1] order by ���"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    
    If mrsPati.EOF And (Not mobjFrom Is Nothing) Then '��ձ༭�б�
        dt��ʼʱ�� = CDate(Split(rptData.SelectedRows(0).Record(cold_����ʱ��).Value, "|")(0))
        dt����ʱ�� = CDate(Split(rptData.SelectedRows(0).Record(cold_����ʱ��).Value, "|")(1))
        btnNoEdit = Val(rptData.SelectedRows(0).Record(cold_����״̬).Value) = 1 Or (Not CheckPriv("ҽ�����Ӱ�"))
        Call mobjFrom.zlRefresh(0, 0, Val(rptData.SelectedRows(0).Record(cold_����ID).Value), 0, Val(rptData.SelectedRows(0).Record(cold_��¼id).Value), dt��ʼʱ��, dt����ʱ��, btnNoEdit, "")
    End If
    
    With vsDetail
        .Redraw = flexRDNone
        .Rows = mrsPati.RecordCount + 1
        Do While Not mrsPati.EOF
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("���")) = mrsPati!���
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����id")) = mrsPati!����ID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����id")) = mrsPati!����ID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("��ҳID")) = mrsPati!��ҳID & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����")) = mrsPati!���� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����")) = mrsPati!���� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("�Ա�")) = mrsPati!�Ա� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����")) = mrsPati!���� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("����")) = mrsPati!���� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("��ʶ��")) = mrsPati!��ʶ�� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("��Ժʱ��")) = mrsPati!��Ժʱ�� & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("��Ժ��ʽ")) = mrsPati!��Ժ��ʽ & ""
            .TextMatrix(mrsPati.AbsolutePosition, .ColIndex("��������")) = mrsPati!�������� & ""
            If IsNull(mrsPati!��������) Then
                .Cell(flexcpBackColor, mrsPati.AbsolutePosition, 0, mrsPati.AbsolutePosition, .Cols - 1) = RGB(255, 239, 219)
            End If
            mrsPati.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > 1 Then
            vsDetail.Row = 1
            vsDetail.ShowCell 1, 0
            If .Rows > 2 Then
                vsDetail.Cell(flexcpPicture, 1, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        End If
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .AutoSize .ColIndex("��������")
    End With
    Call LoadNum(lngId)
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub LoadNum(ByVal lngId As Long)
'��̬��ʾÿ����¼��ʵ�ʻ������
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long, lngMax As Long, lngMaxNum As Long, lngNum As Long
    Dim lngHeight As Long
    Dim objlbl As Object
    
    For lngIndex = 1 To lblTypeNum.UBound
        Unload lblTypeNum(lngIndex)
    Next
    gstrSQL = "Select ��Ŀ, ���� From ҽ�����Ӱ���� Where ��¼id =[1] Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    lblTypeNum(0).Visible = False
    lngIndex = 0
    lngMax = rsTemp.RecordCount - 1
    If rsTemp.RecordCount > 1 Then rsTemp.MoveFirst
    For lngIndex = 0 To lngMax
        If rsTemp!��Ŀ = "סԺ��" Then
            lblAllNum.Caption = "���˻������ ��" & rptData.SelectedRows(0).Record(cold_����).Value & "��  סԺ��������" & rsTemp!����
            lngMax = lngMax - 1
        Else
            If lngIndex = 0 Then
                lblTypeNum(0).Visible = True
                lblTypeNum(0).Caption = rsTemp!��Ŀ & "������" & rsTemp!����
                lblTypeNum(0).Width = 1800
            Else
                Load lblTypeNum(lngIndex)
                lblTypeNum(lngIndex).Caption = rsTemp!��Ŀ & "������" & rsTemp!����
                Set objlbl = lblTypeNum(lngIndex)
                Set lblTypeNum(lngIndex).Container = picNum
                lngNum = Fix(lngIndex / 3)
                If lngNum = lngIndex / 3 Then
                    lblTypeNum(lngIndex).Move lblTypeNum(0).Left, lblTypeNum(0).Top + (lblTypeNum(0).Height + 120) * lngNum, 1800, lblTypeNum(0).Height
                Else
                    lblTypeNum(lngIndex).Move lblTypeNum(lngIndex - 1).Left + lblTypeNum(lngIndex - 1).Width + 50, lblTypeNum(lngIndex - 1).Top, 1800, lblTypeNum(0).Height
                End If
                lblTypeNum(lngIndex).Visible = True
            End If
        End If
        rsTemp.MoveNext
    Next
    lngMaxNum = Fix(lngMax / 3) + 1
    lineNumY.X1 = picNum.Width
    lineNumY.X2 = picNum.Width
    lineNumY.Y1 = 0
    lineNumY.Y2 = picNumBack.Height
    lngHeight = lngMaxNum * (lblTypeNum(0).Height + 120) + 120
    If lngHeight <= picNumBack.Height Then lngHeight = picNumBack.Height
    picNum.Height = lngHeight
    If Fix(lngMax / 3) > 3 Then
        lineNumY.Visible = True
        picNumUp.Visible = False
        picNumDown.Visible = True
        picNumUp.Top = 0
        picNumDown.Top = picNumBack.Height - picNumDown.Height
    Else
        picNumBack.BackColor = picNum.BackColor
        lineNumY.Visible = False
        picNumUp.Visible = False
        picNumDown.Visible = False
    End If
    Call picDataIn_Resize
End Sub

Private Function CheckPriv(ByVal strPri As String) As Boolean
'�ж��Ƿ����ĳ��Ȩ��
    If InStr(";" & mstrPriv & ";", ";" & strPri & ";") > 0 Then
        CheckPriv = True
    End If
End Function

Private Sub tvwSubject_DblClick()
    
    txtSubject.Text = tvwSubject.SelectedItem.Text
    txtSubject.Tag = Mid(tvwSubject.SelectedItem.Key, 2)
    Call LoadDept(Mid(tvwSubject.SelectedItem.Key, 2))
    tvwSubject.Visible = False
    Call LoadType
End Sub

Private Sub tvwSubject_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call tvwSubject_LostFocus
    End If
End Sub

Private Sub tvwSubject_LostFocus()
    tvwSubject.Visible = False
End Sub

Private Sub vsDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim dt��ʼʱ�� As Date, dt����ʱ�� As Date
    Dim btnNoEdit As Boolean
    
    If mblnLoading Then Exit Sub
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsDetail
        If NewRow = 1 Then
            If .Rows = 2 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
        End If

        btnNoEdit = Val(rptData.SelectedRows(0).Record(cold_����״̬).Value) = 1 Or (Not CheckPriv("ҽ�����Ӱ�"))
        With vsDetail
            If (Not mobjFrom Is Nothing) And .TextMatrix(.Row, .ColIndex("����ID")) <> "" Then
                dt��ʼʱ�� = CDate(Split(rptData.SelectedRows(0).Record(cold_����ʱ��).Value, "|")(0))
                dt����ʱ�� = CDate(Split(rptData.SelectedRows(0).Record(cold_����ʱ��).Value, "|")(1))
                Call mobjFrom.zlRefresh(Val(.TextMatrix(.Row, .ColIndex("����ID"))), Val(.TextMatrix(.Row, .ColIndex("��ҳID"))), Val(rptData.SelectedRows(0).Record(cold_����ID).Value), Val(.TextMatrix(.Row, .ColIndex("����ID"))), Val(rptData.SelectedRows(0).Record(cold_��¼id).Value), dt��ʼʱ��, dt����ʱ��, btnNoEdit, .TextMatrix(.Row, .ColIndex("����")))
            End If
        End With
    End With
End Sub

Private Sub vsDetail_Click()
    Dim blnBegin As Boolean
    Dim lngRow As Long, lngId As Long, lngColor As Long
    Dim i As Long
    
    On Error GoTo errH
    
    With vsDetail
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row - 1
            End If
        ElseIf .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row + 1
            End If
        End If
        If lngRow = 0 Then Exit Sub
        '�����ƶ�ʱ������ǲ���ģ���ʾ���˵�˳��
        '���,����,����,�Ա�,����,����,��ʶ��,��Ժʱ�䣬��Ժ��ʽ�����ߣ���ϣ���������
        
        lngId = .TextMatrix(.Row, .ColIndex("����id"))
        lngColor = .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1)
        mrsPati.Filter = "����id=" & .TextMatrix(lngRow, .ColIndex("����id"))
        If mrsPati.RecordCount = 1 Then
            For i = 0 To .Cols - 1
                '��Ų���Ҫ����
                If i <> .ColIndex("���") And i <> .ColIndex("����") And i <> .ColIndex("����") Then
                    .TextMatrix(.Row, i) = mrsPati.Fields(.ColKey(i)).Value & ""
               End If
            Next
            .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1)
        Else
            MsgBox "�����ѱ�ɾ�����޷����ƻ�����!", vbExclamation, Me.Caption
            Exit Sub
        End If

        mrsPati.Filter = "����id=" & lngId
        If mrsPati.RecordCount = 1 Then
            For i = 0 To .Cols - 1
                '��Ų���Ҫ����
                If i <> .ColIndex("���") And i <> .ColIndex("����") And i <> .ColIndex("����") Then
                    .TextMatrix(lngRow, i) = mrsPati.Fields(.ColKey(i)).Value & ""
               End If
            Next
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
        Else
            MsgBox "�����ѱ�ɾ�����޷����ƻ�����!", vbExclamation, Me.Caption
        End If
        '���ݿ��еĽ��Ӱ��������Ӧһ�µ���
        gcnOracle.BeginTrans: blnBegin = True
        gstrSQL = "Zl_ҽ�����Ӱ�����_Edit(3," & .TextMatrix(.Row, .ColIndex("����id")) & ",NULL," & .TextMatrix(.Row, .ColIndex("���")) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
        
        gstrSQL = "Zl_ҽ�����Ӱ�����_Edit(3," & .TextMatrix(lngRow, .ColIndex("����id")) & ",NULL," & .TextMatrix(lngRow, .ColIndex("���")) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�������")
        gcnOracle.CommitTrans
        .Row = lngRow
        .ShowCell lngRow, 1
    End With
    Exit Sub
errH:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub GetFrom()
    Set mobjFrom = New frmShiftEdit
    mobjFrom.BorderStyle = FormBorderStyleConstants.vbBSNone '����Ϊ�ޱ߿�
    mobjFrom.Caption = "���˽��Ӱ����ݱ༭"
    Set mobjFrom.gfrmParent = Me
    
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
        .InsertItem(0, "���Ӱಡ�˱༭", mobjFrom.hWnd, 0).Tag = "���Ӱಡ�˱༭"
        .Item(0).Selected = True
    End With
End Sub

Public Sub SetEnable(Optional intType As Integer)
    '�����������Ƿ�����༭
    'intType 0=����,1=������
    picDataIn.Enabled = intType = 0
    picPatient.Enabled = intType = 0
    mblnEdit = intType <> 0
End Sub

Public Sub RefreshEdit(ByVal lngUPid As Long)
    'ˢ�½��ಡ���б�
    Dim lngFind As Long
    Call rptData_SelectionChanged
    
    '��λ����
    lngFind = vsDetail.FindRow(lngUPid, vsDetail.FixedRows, vsDetail.ColIndex("����ID"))
    If lngFind > 0 Then
        vsDetail.Row = lngFind
        Call vsDetail.ShowCell(lngFind, vsDetail.ColIndex("����"))
    Else
        If vsDetail.Rows - 1 > 0 Then
            vsDetail.Row = vsDetail.Rows - 1
            Call vsDetail.ShowCell(vsDetail.Rows - 1, vsDetail.ColIndex("����"))
        End If
    End If
End Sub

Private Sub vsDetail_DblClick()
    If (Not mobjFrom Is Nothing) And vsDetail.TextMatrix(vsDetail.Row, vsDetail.ColIndex("����ID")) <> "" Then
        Call mobjFrom.EditState
    End If
End Sub


