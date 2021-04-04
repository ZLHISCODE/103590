VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatiStamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˱������"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   Icon            =   "frmPatiStamp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraUnit 
      Height          =   1815
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDays 
         Height          =   300
         Left            =   960
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkSpecial 
         Caption         =   "Ӧ�������ⲡ��ͼ������"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   1170
         Width           =   3045
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0��ʾ������Ч"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ч����"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1905
   End
   Begin VB.Frame fraInfo 
      Height          =   4575
      Left            =   5520
      TabIndex        =   11
      Top             =   945
      Width           =   3975
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   1080
         ScaleHeight     =   2265
         ScaleWidth      =   2625
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
         Begin VB.VScrollBar HScr 
            Height          =   2295
            LargeChange     =   50
            Left            =   2400
            Max             =   100
            SmallChange     =   100
            TabIndex        =   17
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic��� 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   360
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   13
            Top             =   120
            Width           =   1335
            Begin VB.PictureBox picIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Index           =   0
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   615
               TabIndex        =   14
               Top             =   120
               Width           =   615
               Begin VB.Image imgICon 
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmPatiStamp.frx":6852
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Label lblSelect 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  TabIndex        =   16
                  Top             =   120
                  Width           =   300
               End
               Begin VB.Label lblInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "PDA"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  Top             =   450
                  Width           =   270
               End
            End
         End
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdImage 
         Appearance      =   0  'Flat
         Caption         =   "&P"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   720
         Width           =   270
      End
      Begin MSComctlLib.ImageCombo imaCustom 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���ͼ��"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���Ա��"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���˵��"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   22
         Top             =   1260
         Width           =   720
      End
   End
   Begin VB.Frame fraLine 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   100
   End
   Begin VB.Frame fraUd 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4935
      Begin XtremeReportControl.ReportControl UnitReportControl 
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   0
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   420
      Left            =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
      _cx             =   1508
      _cy             =   741
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6135
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiStamp.frx":6F54
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
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
      Left            =   1680
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatiStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FSHIFT = 4
Const FCONTROL = 8
Const FALT = 16

Const VK_DELETE = &H2E
Const VK_F1 = &H70
Const VK_F5 = &H74
Const VK_INSERT = &H2D

Const conMenu_FilePopup = 1    '�ļ�
Const conMenu_EditPopup = 3    '�༭
Const conMenu_ViewPopup = 7    '�鿴
Const conMenu_ToolPopup = 8    '����
Const conMenu_HelpPopup = 9    '����
Const conMenu_File_PrintSet = 101        '*��ӡ����(&S)��
Const conMenu_File_Preview = 102         '*Ԥ��(&V)
Const conMenu_File_Print = 103           '*��ӡ(&P)
Const conMenu_File_Excel = 104           '�����&Excel��
Const conMenu_Edit_Save = 3091        '*����
Const conMenu_File_Exit = 191            '*�˳�(&X)
Const conMenu_Edit_Reuse = 3009      '*����(&U)
Const conMenu_Edit_FileMan = 3047
Const conMenu_Edit_NewParent = 3051   '*�·���(&N)
Const conMenu_Edit_ModifyParent = 3053    '*�޸ķ���(&M)
Const conMenu_Edit_DeleteParent = 3054    '*ɾ������(&D)
Const conMenu_Edit_Leave_Add = 3561    '����
Const conMenu_Edit_NewItem = 3001    '*����Ŀ(&A)
Const conMenu_Edit_Modify = 3003     '*�޸�(&M)
Const conMenu_Edit_Delete = 3004     '*ɾ��(&D)
Const conMenu_View_ToolBar = 701              '������(&T)
Const conMenu_View_ToolBar_Button = 7011         '��׼��ť(&S)
Const conMenu_View_ToolBar_Text = 7012           '�ı���ǩ(&T)
Const conMenu_View_ToolBar_Size = 7013           '��ͼ��(&B)
Const conMenu_View_StatusBar = 702            '״̬��(&S)
Const conMenu_View_Refresh = 791              '*ˢ��(&R)
Const conMenu_Help_Help = 901        '*��������(&H)
Const conMenu_Help_Web = 902         '&WEB�ϵ�����
Const conMenu_Help_Web_Home = 9021       '������ҳ(&H)
Const conMenu_Help_Web_Forum = 9023      '������̳(&F)
Const conMenu_Help_Web_Mail = 9022       '*���ͷ���(&M)
Const conMenu_Help_About = 991       '����(&A)��
Const conMenu_View_Find = 721

Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Const COL_NULL = 0
Const COL_��ע = 1
Const COL_˵�� = 2
Const COL_������� = 3
Const COL_��Ч���� = 4
Const COL_ԭʼ���� = 5
Const COL_ԭʼ��� = 6
Const COL_����˵�� = 7
Const COL_�Ƿ����� = 8
  
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private mRect As RECT

Private Type TYPE_UNIT
    ����ID  As Long
    ������� As Long
    ������ As Long
    ˵�� As String
    ͼ������ As Long
    ��Ч���� As Long
    ԭʼ���� As Long
    ԭʼ��� As Long
End Type

Private mUnit As TYPE_UNIT

Const Enable_Color = &HE0E0E0
Const UnEnable_Color = &H80000005

Private mblnChange As Boolean '��¼������ݱ䶯
Private mstrSubject As String '��Ƿ�������
Private mlngDay As Long '��Ƿ�������
Private mintSpecial As Integer '��Ƿ����Ƿ�Ӧ����������Ⱥ
Private mLngCount As Long  '��ű�Ƿ�����Ŀ

Private m����ID As Long
Private mstr�������� As String

Private mcbrToolBars As CommandBar  '������
Private mcbrMenuBars As CommandBarControl
Const mlngImgIndex As Long = 0 '����ͼƬ�����ӵڼ�����ʼ��ʾ

Private mblnOK As Boolean
Private mrsData As New ADODB.Recordset

Public Function ShowMe(ByVal FrmParent As Form) As Boolean
    mblnOK = False
    Me.Show 1, FrmParent
    ShowMe = mblnOK
End Function

Private Sub cboUnit_Click()
    If cboUnit.ListCount > 0 And m����ID <> Val(cboUnit.ItemData(cboUnit.ListIndex)) Then
        m����ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
        mstr�������� = cboUnit.Text
    
        Call RefreshData
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboUnit.hwnd, KeyAscii)
End Sub

Private Sub cbo���_Click()
'-------------------------------------------------
'����:����ѡ��������Ÿı�������λ��
'-------------------------------------------------
    Dim strTag As String
    Dim lngPreID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim lngRowIndex As Long, lngRow As Long, lngOldID As Long
    Dim strFileds As String, strValues As String
    Dim str��� As String, strCaption As String
    Dim intDay As Integer, intSpecial As Integer
    
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    If cbo���.ListIndex = -1 Or fraInfo.Tag = "����" Or mblnChange = False Then Exit Sub
    If UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count <> 0 Then Exit Sub
    If mrsData Is Nothing Then Exit Sub
    
    strFileds = "�������," & adDouble & ",18|������," & adDouble & ",18|˵��," & adLongVarChar & ",100|ͼ������," & _
        adDouble & ",18|��Ч����," & adDouble & ",18|�Ƿ�����," & adInteger & ",1|ԭʼ�������," & adDouble & ",18|ԭʼ������," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    'A.�������,A.������,A.˵��,A.ͼ������,A.��Ч����,A.�Ƿ�����,A.������� ԭʼ�������,A.������ ԭʼ������
    strFileds = "�������|������|˵��|ͼ������|��Ч����|�Ƿ�����|ԭʼ�������|ԭʼ������"
    
    lngRowIndex = UnitReportControl.FocusedRow.Index
    
    str��� = ""
    mrsData.Filter = ""
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        If Not UnitReportControl.Rows(lngRow).GroupRow Then
            lngOldID = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_�������).Record.Tag, "-")(0))
            mrsData.Filter = "�������=" & lngOldID & " and ������=0"
            If mrsData.RecordCount > 0 Then
                strCaption = Nvl(mrsData!˵��)
                intDay = Val(Nvl(mrsData!��Ч����))
                intSpecial = Val(Nvl(mrsData!�Ƿ�����))
            End If
            
            If UnitReportControl.Rows(lngRow).Index = lngRowIndex Then
                mUnit.������� = Val(cbo���.ItemData(cbo���.ListIndex))
                lngPreID = AgainComputePreId(Val(cbo���.ItemData(cbo���.ListIndex))) '��ȡ������
                mUnit.������ = lngPreID
                
                mrsData.Filter = "�������=" & mUnit.������� & " and ������=0"
                If mrsData.RecordCount > 0 Then
                    mUnit.��Ч���� = Val(Nvl(mrsData!��Ч����))
                End If
                str��� = mUnit.������� & "-" & mUnit.������ & "-" & m����ID & "-" & mUnit.��Ч����
            Else
                mUnit.������� = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_�������).Record.Tag, "-")(0))
                mUnit.������ = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_�������).Record.Tag, "-")(1))
                mUnit.��Ч���� = intDay 'Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_��Ч����).Value, 0))
            End If
                        
            mUnit.˵�� = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_˵��).Value)
            mUnit.ͼ������ = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_��ע).Icon, 0))
            mUnit.ԭʼ���� = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_ԭʼ����).Value, 0)
            mUnit.ԭʼ��� = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_ԭʼ���).Value, 0)
            If mUnit.������� <> mUnit.ԭʼ���� Then '������ű仯ʱ�������Ƿ��Ѿ�ʹ��
                If CheckUseUnit(m����ID, mUnit.ԭʼ����, mUnit.ԭʼ���) Then
                    Call zlControl.CboLocate(cbo���, lngOldID, True)
                    Exit Sub
                End If
            End If
            '�����������Ƿ���� �����ھ����
            rsTemp.Filter = "�������=" & lngOldID & " and ������=0"
            If rsTemp.RecordCount = 0 Then
                strValues = lngOldID & "|" & 0 & "|" & strCaption & "|0|" & _
                    intDay & "|" & intSpecial & "|" & mUnit.ԭʼ���� & "|" & mUnit.ԭʼ���
                Call Record_Add(rsTemp, strFileds, strValues)
            End If
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
                strValues = mUnit.������� & "|" & mUnit.������ & "|" & mUnit.˵�� & "|" & mUnit.ͼ������ & "|" & _
                    mUnit.��Ч���� & "|0|" & mUnit.ԭʼ���� & "|" & mUnit.ԭʼ���
                Call Record_Add(rsTemp, strFileds, strValues)
            End If
        End If
    Next lngRow
    
    rsTemp.Filter = 0
    rsTemp.Sort = "�������,������"
    'Call OutputRsData(rsTemp)
    Call RefreshData(0, str���, rsTemp)
    mblnChange = True
'    With UnitReportControl.FocusedRow.Record(COL_�������)
'        .GroupCaption = "���飺" & cbo���.ItemData(cbo���.ListIndex) & "-" & cbo���.Text
'        strTag = .Record.Tag
'        lngPreID = AgainComputePreId(Val(cbo���.ItemData(cbo���.ListIndex))) '��ȡ������
'        .Record.Tag = cbo���.ItemData(cbo���.ListIndex) & "-" & lngPreID & "-" & Split(strTag, "-")(2)
'    End With
'
'    UnitReportControl.Populate

End Sub

Private Sub cbo���_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cbo���.hwnd, KeyAscii)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cbsMain_Resize()
    Call ResizeState
End Sub

Private Sub cmdImage_Click()
'������ʾ����ͼƬ��Ϣ
    Dim i As Integer, j As Integer
    Dim lngCurXCount As Long
    Dim lngH As Integer, lngW As Integer '��¼picture�ĸ߶ȺͿ��
    Dim lngX1 As Long 'pictrue֮��ļ��
    Dim lngX As Long, lngY As Long  '�趨image�Ķ��������߾�
    Dim lngIndex As Long
    Dim vRect As RECT
    Dim vRect1 As RECT
    
    
    lngIndex = 0
    lngY = 60
    lngX = 60

    imgICon(lngIndex).Top = lngY
    imgICon(lngIndex).Left = lngX
    
    lblSelect(lngIndex).Top = lngY / 2
    lblSelect(lngIndex).Left = lngX / 2
    lblSelect(lngIndex).Width = imgICon(lngIndex).Width + lngX
    lblSelect(lngIndex).Height = imgICon(lngIndex).Height + lngY
    
    lblInfo(lngIndex).FontSize = 8
    lblInfo(lngIndex).Top = lngY + imgICon(lngIndex).Width + lngY / 2
    lblInfo(lngIndex).Caption = zlCommFun.GetPaitSignImageList(1).ListImages(mlngImgIndex + 1).Key
    
    picIcon(lngIndex).Top = 0
    picIcon(lngIndex).Left = 0
    picIcon(lngIndex).Height = imgICon(lngIndex).Height + lngY + lngY / 2 + lblInfo(lngIndex).Height + 10
    picIcon(lngIndex).Width = imgICon(lngIndex).Width + imgICon(lngIndex).Left * 2 + lngX / 2
    
    lngH = picIcon(lngIndex).Height
    lngW = picIcon(lngIndex).Width
    
    lblInfo(lngIndex).Left = (lngW - lblInfo(lngIndex).Width) / 2
    
    '��ȡ����picback��λ�õĿ��
    vRect = zlControl.GetControlRect(imaCustom.hwnd)
    vRect1 = zlControl.GetControlRect(fraInfo.hwnd)
    picBack.Top = vRect.Bottom - vRect1.Top
    picBack.Left = vRect.Left - vRect1.Left
    picBack.Width = vRect1.Right - vRect.Left - 10
    
    pic���.Width = picBack.ScaleWidth - HScr.Width
    
    '����ÿ�пɴ�ŵ�ͼƬ����
    lngCurXCount = (pic���.Width - HScr.Width) \ lngW
    '���¼���λ��
    lngX1 = (pic���.Width - HScr.Width - (lngW * lngCurXCount)) / (lngCurXCount + 1)
    picIcon(lngIndex).Left = lngX1
    
    imgICon(lngIndex).Picture = zlCommFun.GetPaitSignImageList(1).ListImages(mlngImgIndex + 1).Picture
    
    HScr.Top = 0
    HScr.Min = 0
    HScr.Left = pic���.Width
    HScr.Value = 0
    HScr.Height = picBack.ScaleHeight
    
    picBack.Visible = True
    picBack.ZOrder 0
    pic���.Visible = True
    pic���.Top = 0
    pic���.Left = 0
    pic���.SetFocus
    
    For i = 1 To picIcon.Count - 1
        If i < lngCurXCount Then
            picIcon(i).Top = 0
            picIcon(i).Left = lngW * i + (i + 1) * lngX1
        Else
            picIcon(i).Top = lngH * ((i \ lngCurXCount))
            picIcon(i).Left = lngW * (i Mod lngCurXCount) + ((i Mod lngCurXCount) + 1) * lngX1
        End If
        picIcon(i).Width = picIcon(lngIndex).Width
        picIcon(i).Height = picIcon(lngIndex).Height
        
        imgICon(i).Top = imgICon(lngIndex).Top
        imgICon(i).Left = imgICon(lngIndex).Left
        
        lblSelect(i).Top = lblSelect(lngIndex).Top
        lblSelect(i).Left = lblSelect(lngIndex).Left
        lblSelect(i).Width = lblSelect(lngIndex).Width
        lblSelect(i).Height = lblSelect(lngIndex).Height
        
        lblInfo(i).FontSize = lblInfo(lngIndex).FontSize
        lblInfo(i).Top = lblInfo(lngIndex).Top
        lblInfo(i).Left = (lngW - lblInfo(i).Width) / 2
    Next i
    
    pic���.Height = picIcon(i - 1).Top + picIcon(i - 1).Height
    pic���.Refresh
    
    If pic���.ScaleHeight - picBack.ScaleHeight <= 0 Then
        HScr.Max = 0
        HScr.Min = 0
        HScr.Visible = False
    Else
        HScr.Max = pic���.ScaleHeight - picBack.ScaleHeight
        HScr.Visible = True
    End If
    cmdImage.Enabled = False
    
    If Not imaCustom.SelectedItem Is Nothing Then
        lngIndex = imaCustom.SelectedItem.Index
        If lngIndex > 0 And lngIndex <= picIcon.Count Then
            If HScr.Max > 0 Then
                '�������С��ͼƬ��λ�ã�˵��ͼƬ��ʾ����
                If picBack.ScaleHeight < picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height Then
                    If picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height - picBack.ScaleHeight > HScr.Max Then
                        HScr.Value = HScr.Max
                    Else
                        HScr.Value = picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height - picBack.ScaleHeight
                    End If
                End If
            End If
            Call ShowSelect(lngIndex - 1)
        End If
    End If
End Sub

Private Sub LoadICon()
'�����Զ���ͼ��
    Dim i As Integer, j As Integer
    On Error GoTo ErrHand
    i = 1
    For j = mlngImgIndex + 1 To zlCommFun.GetPaitSignImageList(1).ListImages.Count - 1
        Load picIcon(i)
        picIcon(i).Visible = True
        
        '����ͼƬ��Ϣ
        Load imgICon(i)
        imgICon(i).Visible = True
        Set imgICon(i).Container = picIcon(i)
        imgICon(i).Picture = zlCommFun.GetPaitSignImageList(1).ListImages(j + 1).Picture
        
        '����ѡ��ؼ�
        Load lblSelect(i)
        lblSelect(i).Visible = True
        Set lblSelect(i).Container = picIcon(i)
        
        '����ͼƬ˵��
        Load lblInfo(i)
        lblInfo(i).Visible = True
        Set lblInfo(i).Container = picIcon(i)
        lblInfo(i).Caption = zlCommFun.GetPaitSignImageList(1).ListImages(j + 1).Key
        
        i = i + 1
    Next j
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetMarkCount() As Long
    '��ȡ�����Ŀ����
    Dim lngRow As Long
    Dim lngCount As Long
    
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        '������=0��Ϊ���������࣬������ͳ��
        If Not UnitReportControl.Rows(lngRow).GroupRow And UnitReportControl.Rows(lngRow).Childs.Count = 0 Then
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
                lngCount = lngCount + 1
            End If
        End If
    Next lngRow
    
    GetMarkCount = lngCount
End Function

Private Sub RefreshStateInfo()
'------------------------------------------------------------------------------------------------------------------
'���ܣ�ˢ��״̬����ʾ��Ϣ
'-----------------------------------------------------------------------------------------------------------------
    stbThis.Panels(2).Text = "���� " & GetMarkCount & " ��������ݣ�"
End Sub

Private Sub UnLoadImage()
'����:ж��pic��ע�ϵ����пؼ�
    Dim i As Integer
    For i = picIcon.Count - 1 To 1 Step -1
        Unload imgICon(i)
        Unload lblInfo(i)
        Unload lblSelect(i)
        Unload picIcon(i)
    Next i
    picBack.Visible = False
    cmdImage.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
    If KeyCode = 27 And picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    '���ز˵�������
    Call InitCommandBar
    '��ȡ������Ϣ
    Call InitUnits
    '�������������Ϣ
    Call InitReportControl
    '��ȡ����
    Call RefreshData
End Sub

Private Sub AddImage()
'------------------------------------
'����:��������ͼƬ��Ϣ��ImageCombo
'------------------------------------
    Dim objNewItem As ComboItem
    Dim i As Long
 
    imaCustom.ImageList = zlCommFun.GetPaitSignImageList(0)
    For i = 1 To zlCommFun.GetPaitSignImageList(0).ListImages.Count - mlngImgIndex
        Set objNewItem = imaCustom.ComboItems.Add(i, "A" & i, zlCommFun.GetPaitSignImageList(0).ListImages(mlngImgIndex + i).Key, mlngImgIndex + i)
    Next i
    
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(vsfPrint, UnitReportControl) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "������������嵥"
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

Private Sub InitCommandBar()
'����:��ʼ���˵���
    Dim cbrTools As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strProductName As String
    On Error GoTo ErrHand
    
    strProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .ShowTextBelowIcons = False
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False '��ʾͼ��
    End With
    
        '�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBars.ID = conMenu_FilePopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��(&Z)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBars.ID = conMenu_EditPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "��������(&I)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸ķ���(&U) ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ������(&E)")
    
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBars.ID = conMenu_ViewPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBars.ID = conMenu_HelpPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & strProductName)
        
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, strProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, strProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '�����
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Reuse
        .Add FSHIFT, VK_INSERT, conMenu_Edit_NewParent
        .Add FSHIFT, VK_DELETE, conMenu_Edit_DeleteParent
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '--��ӹ�����
    Set mcbrToolBars = cbsMain.Add("������", xtpBarTop)
    mcbrToolBars.EnableDocking xtpFlagStretched
    With mcbrToolBars.Controls
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_FileMan, "����", -1, False)
        cbrTools.IconId = conMenu_Edit_FileMan
        cbrTools.ToolTipText = "��Ƿ���"
        cbrTools.BeginGroup = True
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewParent, "����"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_ModifyParent, "�޸�"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_DeleteParent, "ɾ��"
        
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_Leave_Add, "���", -1, False)
        cbrTools.IconId = conMenu_Edit_NewItem
        cbrTools.ToolTipText = "�������"
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewItem, "����"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Modify, "�޸�"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Delete, "ɾ��"
        

        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        cbrTools.ToolTipText = "����"
        cbrTools.BeginGroup = True
        
        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��")
        cbrTools.ToolTipText = "ȡ��"

        Set cbrTools = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrTools.ToolTipText = "����"
        cbrTools.BeginGroup = True
        Set cbrTools = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")

    End With
    
    For Each cbrControl In mcbrToolBars.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�������Ҳಡ��������ѡ��
    With mcbrToolBars.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Find, "����")
        objControl.flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "����")
        objCustom.Handle = Me.cboUnit.hwnd
        objCustom.flags = xtpFlagRightAlign
        objControl.IconId = conMenu_View_Find
    End With
    
    '����ͼƬ��Ϣ
    Call AddImage
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitReportControl()
'����:��ʼ��ReportControl

    Dim Column As ReportColumn
    
    With UnitReportControl
        Set Column = .Columns.Add(COL_NULL, " ", 10, False)
        Column.Editable = False: Column.Groupable = False: Column.Sortable = False: Column.Alignment = xtpAlignmentCenter
        Set Column = .Columns.Add(COL_��ע, "��ע", 50, True)
        Column.Editable = False: Column.Groupable = False: Column.AllowDrag = False
        
        Set Column = .Columns.Add(COL_˵��, "˵��", 190, True)
        Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_�������, "�������", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = True
        Set Column = .Columns.Add(COL_��Ч����, "��Ч����", 60, True)
        Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_ԭʼ����, "ԭʼ����", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_ԭʼ���, "ԭʼ���", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_����˵��, "����˵��", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_�Ƿ�����, "�Ƿ�����", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .ShadeGroupHeadings = False
            .NoItemsText = "û�п���ʾ�ı�Ƿ���ͱ��������Ϣ..."
        End With
        
        .AllowColumnResize = False
        .ShowItemsInGroups = False '�Ƿ������Է�������
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList zlCommFun.GetPaitSignImageList(0)
            
        .GroupsOrder.Add .Columns(COL_�������)
        .GroupsOrder(0).SortAscending = True
        .GroupsOrder(0).Groupable = True
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(COL_˵��)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_�������)
        .SortOrder(1).SortAscending = True
    End With
    
    Call LoadICon
End Sub

Private Function RefreshData(Optional lngPreIdx As Long, Optional str��� As String = "", Optional ByVal rsTemp As ADODB.Recordset) As Boolean
'-------------------------------------------------------------
'����:��ȡ�������Ի�����
'����:lngPreIdx ѡ��������,str��� ѡ������Ϣ���������ٶ�λ��
'˵�� lngPreIdx=-1ʱ�����в�����Ƿ�����
'-------------------------------------------------------------
    Dim strUnit As String, StrInfo As String, strDay As String, strOldUnit As String
    Dim lngImgIndex As Long
    Dim blnDouble As Boolean
    Dim lngIndex As Long '��ŵ�ǰ���
    Dim blnRead As Boolean
    Dim strSql As String
    'Dim rsTemp As New ADODB.Recordset
    Dim strSubject As String '��ű�Ƿ������Ϣ
    Dim objRow As ReportRow, i As Long
    Dim strFileds As String, strValues As String
    
    mblnChange = False
    Screen.MousePointer = 11
    On Error GoTo ErrHand
    
    mLngCount = CheckUnitSubject(m����ID)
    
    If rsTemp Is Nothing Then blnRead = True
    If blnRead = False Then
        If rsTemp.State = adStateClosed Then blnRead = True
    End If
    If blnRead = True Then
        
        strFileds = "�������," & adDouble & ",18|������," & adDouble & ",18|˵��," & adLongVarChar & ",100|ͼ������," & _
            adDouble & ",18|��Ч����," & adDouble & ",18,|�Ƿ�����," & adInteger & ",1,|ԭʼ�������," & adDouble & ",18|ԭʼ������," & adDouble & ",18"
        Call Record_Init(mrsData, strFileds)
        strFileds = "�������|������|˵��|ͼ������|��Ч����|�Ƿ�����|ԭʼ�������|ԭʼ������"
         '��ȡ������Ϣ
        strSql = _
            " SELECT A.�������,A.������,A.˵��,A.ͼ������,A.��Ч����,A.�Ƿ�����,A.������� ԭʼ�������,A.������ ԭʼ������" & vbNewLine & _
            " FROM ����������� A,����������� B" & vbNewLine & _
            " WHERE  " & IIF(m����ID = 0, " B.����ID IS NULL ", " A.����ID=B.����ID ") & " And A.�������=B.������� And B.������=0 " & IIF(m����ID = 0, " And A.����ID IS NULL ", " And A.����ID=[1] ") & vbNewLine & _
            " ORDER BY A.�������,A.������"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����������Ϣ", m����ID)
    End If
    
    UnitReportControl.Records.DeleteAll
    
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    With rsTemp
        Do While Not .EOF
            If zlCommFun.Nvl(!������) = 0 Then
                If strSubject <> "" Then
                    strUnit = strSubject
                    StrInfo = "�˷�����û�п���ʾ�ı��������Ϣ..."
                    lngImgIndex = 0
                    AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
                    strSubject = ""
                End If
                mstrSubject = zlCommFun.Nvl(!˵��, "���Ա�ע" & zlCommFun.Nvl(!�������))
                mlngDay = Val(zlCommFun.Nvl(!��Ч����, 0))
                mintSpecial = Val(zlCommFun.Nvl(!�Ƿ�����, 0))
                strSubject = zlCommFun.Nvl(!�������) & "-" & zlCommFun.Nvl(!������) & "-" & m����ID
                strOldUnit = zlCommFun.Nvl(!ԭʼ�������) & "-" & zlCommFun.Nvl(!ԭʼ������) & "-" & m����ID
            Else
                strUnit = zlCommFun.Nvl(!�������) & "-" & zlCommFun.Nvl(!������) & "-" & m����ID
                strOldUnit = zlCommFun.Nvl(!ԭʼ�������) & "-" & zlCommFun.Nvl(!ԭʼ������) & "-" & m����ID
                StrInfo = zlCommFun.Nvl(!˵��)
                strDay = zlCommFun.Nvl(!��Ч����, 0)
                lngImgIndex = zlCommFun.Nvl(!ͼ������, 0)
                AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
                strSubject = ""
            End If
            If blnRead = True Then
                strValues = Val(zlCommFun.Nvl(!�������)) & "|" & Val(zlCommFun.Nvl(!������)) & "|" & zlCommFun.Nvl(!˵��) & "|" & Val(zlCommFun.Nvl(!ͼ������)) & "|" & _
                   Val(zlCommFun.Nvl(!��Ч����)) & "|" & Val(zlCommFun.Nvl(!�Ƿ�����)) & "|" & Val(zlCommFun.Nvl(!ԭʼ�������)) & "|" & Val(zlCommFun.Nvl(!ԭʼ������))
                Call Record_Add(mrsData, strFileds, strValues)
            End If
        .MoveNext
        Loop
    End With
    
    If strSubject <> "" Then
        strUnit = strSubject
        StrInfo = "�˷�����û�п���ʾ�ı��������Ϣ..."
        lngImgIndex = 0
        AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
        strSubject = ""
    End If
    
    UnitReportControl.Populate
    
    If UnitReportControl.Rows.Count <> 0 Then
        Call UnitRefresh(lngPreIdx, str���)
    Else
        Call SetFraResize(True)
        txtName.Enabled = False
        txtName.Text = ""
        txtDays.Enabled = False
        txtDays.Text = ""
        txtName.BackColor = Enable_Color
        txtDays.BackColor = Enable_Color
        chkSpecial.Enabled = False
        chkSpecial.Value = 0
        chkSpecial.Visible = (m����ID = 0)
    End If
    
    Call RefreshStateInfo
    
    '����Ƿ����ò�����Ƿ���(-1��������ʾ)
    If lngPreIdx <> -1 Then
        If mLngCount = 0 Then
            'MsgBox "������" & Split(mstr��������, "-")(1) & "����δ���ò�����Ƿ���,�����.", vbInformation, gstrSysName
        End If
    End If
    
    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
        Call SaveErrLog
    End If
End Function


Private Function UnitRefresh(Optional lngPreIdx As Long, Optional str��� As String = "") As Boolean
'-----------------------------------------------
'����:�����Ŀ�������޸ĺ�λ��ѡ��ļ�¼
'����:lngreIdx �ϴ�ѡ���е�����
'     str��� �ϴ�ѡ���е����� ��ʽ:�������-������-����ID
'-----------------------------------------------
    Dim objRow As ReportRow, i As Long, j As Long
    Dim blnRetrun As Boolean, blnChild As Boolean
    Dim arrCode() As String
    Dim lngRow As Long, lngGroup As Long
    
    If lngPreIdx < 0 Then lngPreIdx = 0
    
    If str��� <> "" Then
        
        str��� = str��� & String(3 - UBound(Split(str���, "-")), "-")
        arrCode = Split(str���, "-")
        blnChild = Val(arrCode(1)) <> 0
        
        If blnChild = True Then
            If GetMarkCount = 0 Then blnChild = False
        End If
        
        If blnChild = True Then
            '�ȿ��ٶ�λ
            If lngPreIdx <= UnitReportControl.Rows.Count - 1 Then
                If Not UnitReportControl.Rows(lngPreIdx).GroupRow And UnitReportControl.Rows(lngPreIdx).Childs.Count = 0 Then
                    If UnitReportControl.Rows(lngPreIdx).Record(COL_�������).Record.Tag = str��� Then
                        Set objRow = UnitReportControl.Rows(lngPreIdx)
                    End If
                End If
            End If
            '�ٽ��в���
            If objRow Is Nothing Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                        If UnitReportControl.Rows(i).Record(COL_�������).Record.Tag = str��� Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        Else
            For i = 0 To UnitReportControl.Rows.Count - 1
                   If UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count > 0 Then
                        If Split(UnitReportControl.Rows(i).Childs(0).Record(COL_�������).Record.Tag, "-")(0) = arrCode(0) And arrCode(1) = 0 Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                   End If
            Next i
        End If
    End If
    
    'ȡ��һ���Ƿ�����
    If objRow Is Nothing Then
        For i = 0 To UnitReportControl.Rows.Count - 1
            If blnChild Then
                If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                    If Val(Split(UnitReportControl.Rows(i).Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
                        Set objRow = UnitReportControl.Rows(i): Exit For
                    End If
                End If
            Else
                Set objRow = UnitReportControl.Rows(i)
                If objRow.GroupRow And objRow.Childs.Count > 0 Then
                    For j = 0 To objRow.Childs.Count - 1
                        If Val(Split(objRow.Childs(j).Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
                            Set objRow = UnitReportControl.Rows(i + 1)
                            Exit For
                        End If
                    Next j
                End If
                Exit For
            End If
        Next
    End If
    
    If Not objRow Is Nothing Then
        blnRetrun = True
        If Not objRow.GroupRow Then
            If Val(Split(objRow.Record(COL_�������).Record.Tag, "-")(1)) = 0 Then
                Set objRow = UnitReportControl.Rows(objRow.Index - 1)
            End If
        End If
        Set UnitReportControl.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        UnitReportControl.FocusedRow.Selected = True
        
    End If
    
    UnitRefresh = blnRetrun
End Function

Private Function AddRecord(ByVal strUnit As String, ByVal lngImgIndex As Long, ByVal StrInfo As String, ByVal lngDay As Long, _
    Optional ByVal strUnitOld As String = "") As ReportRecord
'-------------------------------------------------------------------------------------------
'���ܣ���ReportRecord��Ӳ�����Ǽ�¼
'------------------------------------------------------------------------------------------
    Dim blnParent As Boolean
    Dim Record As ReportRecord
    Set Record = UnitReportControl.Records.Add()
    
    If strUnitOld = "" Then strUnitOld = strUnit
    Dim Item As ReportRecordItem
   
    blnParent = Val(Split(strUnit, "-")(1)) = 0
    
    Set Item = Record.AddItem("")
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem("")
    If lngImgIndex >= mlngImgIndex And lngImgIndex <= zlCommFun.GetPaitSignImageList(0).ListImages.Count - 1 And blnParent = False Then
        Item.Icon = lngImgIndex
    End If
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(StrInfo)
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(Val(Split(strUnit, "-")(0)))
    Item.GroupCaption = "���飺" & Val(Split(strUnit, "-")(0)) & "-" & mstrSubject
    '������� & "-" & ������ & "-" & ����Id & "-" & "��Ч����"
    Item.Record.Tag = strUnit & "-" & lngDay
    
    Set Item = Record.AddItem(IIF(blnParent, "", lngDay)) '��Ч����
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    Record.AddItem CInt(Split(strUnitOld, "-")(0))  '��¼ԭʼ�������
    Record.AddItem CInt(Split(strUnitOld, "-")(1)) '��¼ԭʼ������
    Record.AddItem mstrSubject
    Record.AddItem mintSpecial
    
    Set AddRecord = Record
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ������
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim blnTrue As Boolean
    On Error GoTo errH
    
    '114577:֧�����ù�������ͼ��
     strSql = _
         " Select Distinct A.ID,A.����,A.����" & _
         " From ���ű� A,��������˵�� B " & _
         " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
         " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
         " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
         " Order by A.����"

    cboUnit.Clear
    cboUnit.AddItem "0-��������"
    cboUnit.ItemData(cboUnit.NewIndex) = 0
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, glngUserId)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            
            If m����ID = rsTmp!ID Then
                Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                If cboUnit.ListIndex <> -1 Then blnTrue = True
            End If
            
            If Not blnTrue Then
                If rsTmp!ID = glngDeptId Then  'ֱ����������
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboUnit.hwnd, 0)
    End If
    
    If cboUnit.ListIndex <> -1 Then
        m����ID = cboUnit.ItemData(cboUnit.ListIndex)
        mstr�������� = cboUnit.Text
    End If
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Call ResizeState
End Sub

Private Sub SetControlEnable(Optional blnEnable As Boolean = False)
'------------------------------------------------------------------
'����:�����Ƿ���Ա༭
'------------------------------------------------------------------
        Dim blnNone As Boolean
        Dim i As Integer
        cbo���.Enabled = blnEnable
       
        cbo���.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        
        blnNone = IIF(fraInfo.Tag = "����", True, False)
        
        If blnNone = False Then
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    blnNone = False
                Else
                    blnNone = True
                End If
            Else
                blnNone = True
            End If
        End If
        
        If UnitReportControl.Records.Count = 0 Then
            cbo���.ListIndex = -1
        Else
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    cbo���.ListIndex = SetCboIndex(cbo���, Val(Split(UnitReportControl.SelectedRows(0).Record(COL_�������).Record.Tag, "-")(0)))
                Else
                    cbo���.ListIndex = SetCboIndex(cbo���, Val(Split(UnitReportControl.SelectedRows(0).Childs(0).Record(COL_�������).Record.Tag, "-")(0)))
                End If
            End If
        End If
        
        If blnNone = True Then lblSet(9).Tag = "": cbo���.Tag = ""
        txtInfo.Enabled = blnEnable
        txtInfo.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then txtInfo.Text = "": lblSet(8).Tag = "":: txtInfo.Tag = ""
        imaCustom.Enabled = blnEnable
        imaCustom.Locked = True
        imaCustom.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then imaCustom.Text = "": lblSet(7).Tag = "": imaCustom.Tag = ""
        
        cmdImage.Enabled = blnEnable
        
        If blnEnable = True And fraInfo.Visible = True Then cbo���.SetFocus
End Sub

Private Sub ResizeState()
'����:���ô������пؼ�λ��
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim blnGourp As Boolean
    Dim objRow As ReportRow
    Dim i As Integer
    
    If Me.WindowState = 1 Then Exit Sub
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If lngTop = 0 Then lngTop = 600
    
    mRect.Top = lngTop
    mRect.Left = lngLeft
    mRect.Right = lngRight
    mRect.Bottom = lngBottom
    
    fraUd.Top = lngTop
    fraUd.Left = 0
    fraUd.Width = ScaleWidth * 0.6
    fraUd.Height = lngBottom - lngTop
    
    UnitReportControl.Move 0, 100, fraUd.Width - 50, fraUd.Height - 150
    
    fraLine.Width = 50
    fraLine.Top = lngTop
    fraLine.Left = ScaleWidth * 0.6
    fraLine.Height = lngBottom - lngTop

    If InStr(1, ",����,�޸�,", "," & fraInfo.Tag & ",") = 0 And InStr(1, ",����,�޸�,", "," & fraUnit.Tag & ",") = 0 Then
        blnGourp = False
        If UnitReportControl.Rows.Count > 0 Then
            If GetMarkCount > 0 Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If UnitReportControl.Rows(i).Selected = True Then
                        Set objRow = UnitReportControl.Rows(i)
                    End If
                Next i
                
                If Not objRow Is Nothing Then
                    If objRow.GroupRow Then
                        blnGourp = True
                    Else
                        blnGourp = False
                    End If
                Else
                    blnGourp = False
                End If
            Else
                blnGourp = True
            End If
        Else
            blnGourp = True
        End If
    ElseIf InStr(1, ",����,�޸�,", "," & fraInfo.Tag & ",") = 0 Then
        blnGourp = True
    Else
        blnGourp = False
    End If
    
    Call SetFraResize(blnGourp)
End Sub

Private Sub SetFraResize(Optional blnGroup As Boolean = False)
    If blnGroup = True Then
        fraInfo.Visible = False
        fraInfo.Enabled = False
        fraUnit.Visible = True
        fraUnit.Enabled = True
        fraUnit.Top = mRect.Top
        fraUnit.Width = ScaleWidth * 0.4 - fraLine.Width
        fraUnit.Height = mRect.Bottom - mRect.Top
        fraUnit.Left = ScaleWidth * 0.6 + fraLine.Width
    Else
        fraUnit.Visible = False
        fraUnit.Enabled = False
        fraInfo.Visible = True
        fraInfo.Enabled = True
        fraInfo.Top = mRect.Top
        fraInfo.Width = ScaleWidth * 0.4 - fraLine.Width
        fraInfo.Height = mRect.Bottom - mRect.Top
        fraInfo.Left = ScaleWidth * 0.6 + fraLine.Width
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrSubject = ""
    mlngDay = 0
    mintSpecial = 0
    Call UnLoadImage
    mblnOK = (fraUd.Tag = "1")
    If Not (mrsData Is Nothing) Then Set mrsData = Nothing
'    If mblnChange = True Then
'        If MsgBox("������" & Split(mstr��������, "-")(1) & "����������Ѿ������ı䣬��ȷ��Ҫ�˳���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
'    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub HScr_Change()
    pic���.Top = HScr.Top - HScr.Value
    If picBack.Visible = True Then picBack.SetFocus
End Sub

Private Sub HScr_Scroll()
    pic���.Top = HScr.Top - HScr.Value
End Sub

Private Sub imaCustom_Click()
     Call showIcon(imaCustom.SelectedItem.Index - 1)
End Sub

Private Sub imaCustom_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub imaCustom_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii <> vbKeyReturn Then
        Call zlControl.CboMatchIndex(imaCustom.hwnd, KeyAscii)
    Else
        '�����ûس���ImageComboͼ�ζ�ʧ���˴�������ʾͼ��
        If KeyAscii = vbKeyReturn Then
            If imaCustom.Text <> "" Then
                 For i = 1 To zlCommFun.GetPaitSignImageList(0).ListImages.Count - mlngImgIndex
                    If imaCustom.Text = zlCommFun.GetPaitSignImageList(0).ListImages(mlngImgIndex + i).Key Then
                        imaCustom.ComboItems(i).Selected = True
                    End If
                Next i
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub imgIcon_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub showIcon(ByVal Index As Integer)
'����:չʾ�û�ѡ���ͼ��
    If Index < 0 Then Exit Sub
    imaCustom.ComboItems(Index + 1).Selected = True
    picBack.Visible = False
    cmdImage.Enabled = True
    
    If fraInfo.Tag = "�޸�" Then
        With UnitReportControl.FocusedRow.Record(COL_��ע)
            .Icon = Index + mlngImgIndex
        End With
        UnitReportControl.Populate
    End If
    
    If (txtInfo.Text = "" Or txtInfo.Tag <> "�ı�") And IIF(fraInfo.Tag = "�޸�", lblSet(8).Tag = "", True) Then txtInfo.Text = imaCustom.ComboItems(Index + 1).Text
End Sub

Private Sub ShowSelect(ByVal Index As Integer)
'����:ѡ��ͼ��
    Dim i As Integer
    lblSelect(Index).BackColor = &H8000000D
    lblInfo(Index).BackColor = &H8000000D
    For i = 0 To zlCommFun.GetPaitSignImageList(1).ListImages.Count - mlngImgIndex - 1
        If i <> Index Then
            lblSelect(i).BackColor = &H8000000E
            lblInfo(i).BackColor = &H8000000E
        End If
    Next i
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ShowSelect(Index)
End Sub

Private Function AgainComputePreId(ByVal lngPreVId As Long, Optional bln���� As Boolean = False) As Long
'--------------------------------------
'����:�����������
'������lngPreVId���������
'--------------------------------------
    Dim lngTmp As Long
    Dim blnTrue As Boolean
    Dim i As Integer
    For i = 0 To UnitReportControl.Records.Count - 1
        If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_�������).Record.Tag, "-")(0)) Then
            If lngTmp < Val(Split(UnitReportControl.Records(i).Item(COL_�������).Record.Tag, "-")(0)) Then
                lngTmp = Val(Split(UnitReportControl.Records(i).Item(COL_�������).Record.Tag, "-")(0))
            End If
        End If
    Next i
    
    If bln���� = True Then
        '�����ļ�¼ֱ�Ӽ�һ
        lngTmp = lngTmp + 1
    Else
        '���Ա�Ǹı�ʱ�������ǰ��ͬ�����ֱ�Ӽ�һ������ظ�����ǰ������ǰ����Ƿ�ʹ�ã�ʹ�õĻ����»�ȡ�µ����
        If Val(Split(UnitReportControl.FocusedRow.Record(COL_�������).Record.Tag, "-")(0)) = lngPreVId Then
            '���ԭʼ����Ƿ�������¼ʹ��
            For i = 0 To UnitReportControl.Records.Count - 1
                If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_�������).Record.Tag, "-")(0)) Then
                    If UnitReportControl.FocusedRow.Record(COL_ԭʼ���).Value = Val(Split(UnitReportControl.Records(i).Item(COL_�������).Record.Tag, "-")(1)) Then
                        blnTrue = True
                    End If
                End If
            Next i
            
            If blnTrue = True Then
                lngTmp = UnitReportControl.FocusedRow.Record(COL_ԭʼ���).Value
            Else
                lngTmp = lngTmp + 1
            End If
        Else
            lngTmp = lngTmp + 1
        End If
    End If

    AgainComputePreId = lngTmp
    
End Function


Private Function SaveData() As Boolean
'------------------------------------------------------------------
'���ܣ�����������ݱ���
'------------------------------------------------------------------
    Dim lngRowIndex As Long 'ѡ���е�����
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSql As String
    Dim blnTran As Boolean
    Dim strSQLAdd() As String
    Dim StrSQLMod() As String
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    ReDim Preserve strSQLAdd(0 To 0)
    ReDim Preserve StrSQLMod(0 To 0)
    lngRowIndex = 0
    
    If InStr(1, ",����,�޸�,", "," & fraInfo.Tag & ",") <> 0 Then
        If imaCustom.Text = "" Then
            MsgBox "���ͼ�β���Ϊ��,��ѡ����ͼ�κ��ڽ��б������.", vbInformation, gstrSysName
            imaCustom.SetFocus
            Exit Function
        End If
    End If
    
    If InStr(1, ",����,�޸�,", "," & fraUnit.Tag & ",") <> 0 Then
        If Trim(txtName.Text) = "" Then
            MsgBox "������Ʋ���Ϊ��,����.", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        
        If Not zlCommFun.StrIsValid(txtDays.Text, 3, txtDays.hwnd, "��Ч����") Then Exit Function
    End If
    
    '�޸�
    If fraInfo.Tag = "�޸�" Then
        If UnitReportControl.FocusedRow Is Nothing Then Exit Function
        
        lngRowIndex = UnitReportControl.FocusedRow.Index
        mUnit.����ID = m����ID
        mUnit.������� = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_�������).Record.Tag, "-")(0))
        mUnit.������ = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_�������).Record.Tag, "-")(1))
        mUnit.˵�� = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_˵��).Value)
        mUnit.˵�� = Trim(txtInfo.Text)
        mUnit.ͼ������ = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_��ע).Icon, 0))
        mUnit.��Ч���� = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_��Ч����).Value, 0))
        mUnit.ԭʼ���� = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_ԭʼ����).Value, 0)
        mUnit.ԭʼ��� = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_ԭʼ���).Value, 0)
        
        mrsData.Filter = "�������=" & Val(mUnit.�������) & " and ������=0"
        If mrsData.RecordCount > 0 Then
            mUnit.��Ч���� = Val(Nvl(mrsData!��Ч����))
        End If
        
        '�޸ĺ��������κα仯,����������д�����
        If CheckChange Then
            If mUnit.������� <> mUnit.ԭʼ���� Then '������ŷ����ı�
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_�����������_Delete(" & mUnit.����ID & "," & mUnit.ԭʼ���� & "," & mUnit.ԭʼ��� & ")"
                mUnit.������ = GetNewPreID(mUnit.����ID, mUnit.�������)
                
                strTmp1 = mUnit.������� & "-" & mUnit.������
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_�����������_Insert(" & mUnit.����ID & "," & mUnit.������� & "," & _
                mUnit.������ & ",'" & mUnit.˵�� & "'," & mUnit.ͼ������ & "," & mUnit.��Ч���� & ")"
            Else
                strTmp1 = mUnit.������� & "-" & mUnit.ԭʼ���
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_�����������_Update(" & mUnit.����ID & "," & mUnit.������� & "," & _
                    mUnit.ԭʼ��� & ",'" & mUnit.˵�� & "'," & mUnit.ͼ������ & "," & mUnit.��Ч���� & ")"
            End If
            
            If IsEqualInfo(txtInfo.Text, False, strTmp1) = False Then
                If txtInfo.Enabled And txtInfo.Visible Then txtInfo.SetFocus
                Exit Function
            End If
                
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            
            fraUd.Tag = "1"
        Else
            strTmp1 = mUnit.������� & "-" & mUnit.ԭʼ���
        End If
        strTemp = strTmp1 & "-" & mUnit.����ID & "-" & Val(mUnit.��Ч����)
    End If
    
    '����
    If fraInfo.Tag = "����" Then
        If cbo���.ListIndex = -1 Then Exit Function
        If IsEqualInfo(txtInfo.Text, False) = False Then
            If txtInfo.Enabled And txtInfo.Visible Then txtInfo.SetFocus
            Exit Function
        End If
        mUnit.����ID = m����ID
        mUnit.������� = cbo���.ItemData(cbo���.ListIndex)
        mUnit.������ = GetNewPreID(mUnit.����ID, mUnit.�������)
        mUnit.˵�� = txtInfo.Text
        mUnit.ͼ������ = imaCustom.SelectedItem.Index - 1 + mlngImgIndex
        mUnit.��Ч���� = 0
        
        For i = 0 To UnitReportControl.Rows.Count - 1
            If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                If Val(Split(UnitReportControl.Rows(i).Record(COL_�������).Record.Tag, "-")(0)) = cbo���.ItemData(cbo���.ListIndex) Then
                    mUnit.��Ч���� = Val(Split(UnitReportControl.Rows(i).Record(COL_��Ч����).Record.Tag, "-")(3))
                    Exit For
                End If
            End If
        Next i
        
        mrsData.Filter = "�������=" & Val(mUnit.�������) & " and ������=0"
        If mrsData.RecordCount > 0 Then
            mUnit.��Ч���� = Val(Nvl(mrsData!��Ч����))
        End If
        
        strTmp1 = mUnit.������� & "-" & mUnit.������
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_�����������_Insert(" & mUnit.����ID & "," & mUnit.������� & "," & _
            mUnit.������ & ",'" & mUnit.˵�� & "'," & mUnit.ͼ������ & "," & mUnit.��Ч���� & ")"
            
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = strTmp1 & "-" & mUnit.����ID & "-" & Val(mUnit.��Ч����)
        
        mstrSubject = cbo���.Text
        Set Record = AddRecord(mUnit.������� & "-" & mUnit.������ & "-" & mUnit.����ID, mUnit.ͼ������, mUnit.˵��, Val(mUnit.��Ч����))
        fraUd.Tag = "1"
        UnitReportControl.Populate
    End If
                
    '������������
    If fraUnit.Tag = "����" Then
        If IsEqualInfo(txtName.Text, True) = False Then
            If txtName.Enabled And txtName.Visible Then txtName.SetFocus
            Exit Function
        End If
        mUnit.������� = GetNewSubjectId(cboUnit.ItemData(cboUnit.ListIndex))
        If mUnit.������� = 0 Then Exit Function
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_�����������_Insert(" & cboUnit.ItemData(cboUnit.ListIndex) & "," & mUnit.������� & "," & _
            0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & "," & IIF(chkSpecial.Value = 0, "NULL", "1") & ")"
        
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = mUnit.������� & "-0-" & mUnit.����ID & "-" & Val(txtDays.Text)
        
        fraUd.Tag = "1"
    End If
    
    '�޸���������
    If fraUnit.Tag = "�޸�" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Function
        
        mUnit.������� = Val(Split(UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_�������).Record.Tag, "-")(0))
        mUnit.����ID = cboUnit.ItemData(cboUnit.ListIndex)
        
        '��Ƿ��෢���仯������޸Ĳ���
        If CheckChange Then
            If IsEqualInfo(txtName.Text, True, mUnit.�������) = False Then
                If txtName.Enabled And txtName.Visible Then txtName.SetFocus
                Exit Function
            End If
            StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_�����������_Update(" & mUnit.����ID & "," & mUnit.������� & "," & _
                0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & "," & IIF(chkSpecial.Value = 0, "NULL", "1") & ")"
            
            strSql = "select ������,˵��,ͼ������,��Ч���� from ����������� where " & IIF(mUnit.����ID = 0, " ����ID IS NULL ", " ����ID=[1] ") & " and  �������=[2] and ������<>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����������", mUnit.����ID, mUnit.�������)
            '����ӷ���������Ƿ�ͷ�����ͬ����ͬ������޸�
            With rsTmp
                Do While Not .EOF
                    If zlCommFun.Nvl(!��Ч����, 0) <> Val(txtDays.Text) Then
                        StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_�����������_Update(" & mUnit.����ID & "," & mUnit.������� & "," & _
                            zlCommFun.Nvl(!������, 0) & ",'" & Replace(zlCommFun.Nvl(!˵��), "'", "") & "'," & zlCommFun.Nvl(!ͼ������, 0) & "," & Val(txtDays.Text) & ")"
                    End If
                .MoveNext
                Loop
            End With
            
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            fraUd.Tag = "1"
        End If
        strTemp = mUnit.������� & "-0-" & mUnit.����ID & "-" & Val(txtDays.Text)
    End If
    
    mblnChange = False
    
    fraInfo.Tag = ""
    fraUnit.Tag = ""
    UnitReportControl.Tag = ""
    '��λ��Ӧ������
    Call RefreshData(lngRowIndex, strTemp)
    fraUd.Enabled = True
    UnitReportControl.SetFocus
    
    SaveData = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRowIndex As Long 'ѡ���е�����
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSql As String
    Dim blnTran As Boolean
    Dim cbrControl As CommandBarControl
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand

    
    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
                    
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        
        Case conMenu_File_Print
            Call zlRptPrint(1)
        
        Case conMenu_File_Excel
            Call zlRptPrint(3)
    
        Case conMenu_View_ToolBar_Button
            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type <> xtpControlLabel Then
                    cbrControl.Style = IIF(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            cbsMain.RecalcLayout
            
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout
            
        Case conMenu_Edit_NewItem     '*����
            fraInfo.Tag = "����"
            fraUnit.Tag = ""
            Call SetFraResize
            Call SetControlEnable(True)
            mblnChange = True
        Case conMenu_Edit_Modify      '*�޸�(&M)
            fraInfo.Tag = "�޸�"
            fraUnit.Tag = ""
            Call SetControlEnable(True)
            mblnChange = True
            
        Case conMenu_Edit_Delete      '*ɾ��(&D)
            If MsgBox("��ȷ��Ҫɾ��������" & Split(mstr��������, "-")(1) & "�����ݡ�" & UnitReportControl.FocusedRow.Record(COL_˵��).Value & "���ı����Ϣ��?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Record(COL_�������).Record.Tag
            
            mUnit.����ID = CInt(Split(strTemp, "-")(2))
            mUnit.������� = CInt(Split(strTemp, "-")(0))
            mUnit.������ = CInt(Split(strTemp, "-")(1))
            
            '�����������ݸò����Ƿ�����ʹ��
            If CheckUseUnit(mUnit.����ID, mUnit.�������, mUnit.������) = True Then Exit Sub
            
            strSql = "Zl_�����������_Delete(" & mUnit.����ID & "," & mUnit.������� & "," & mUnit.������ & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            '��λ����һ��
            lngRowIndex = UnitReportControl.FocusedRow.Index
            
            Call UnitReportControl.Records.RemoveAt(UnitReportControl.FocusedRow.Record.Index)
            UnitReportControl.Populate
            
            If UnitReportControl.Records.Count > 0 Then
                lngRowIndex = IIF(UnitReportControl.Rows.Count - 1 > lngRowIndex, lngRowIndex, UnitReportControl.Rows.Count - 1)
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow And UnitReportControl.Rows(lngRowIndex).Childs.Count <> 0 Then
                    lngRowIndex = lngRowIndex - 1
                End If
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow Then
                    strTemp = UnitReportControl.Rows(lngRowIndex).Childs.Record(COL_�������).Record.Tag
                Else
                    strTemp = UnitReportControl.Rows(lngRowIndex).Record(COL_�������).Record.Tag
                End If
            End If
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
        Case conMenu_Edit_NewParent '*��������
            fraInfo.Tag = ""
            fraUnit.Tag = "����"
            Call SetFraResize(True)
            txtName.Enabled = True
            txtName.Text = ""
            txtDays.Enabled = True
            txtDays.Text = ""
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            chkSpecial.Visible = (m����ID = 0)
            chkSpecial.Enabled = (m����ID = 0)
            chkSpecial.Value = 0
            If m����ID = 0 Then
                mrsData.Filter = "������=0"
                Do While Not mrsData.EOF
                    If Val("" & mrsData!�Ƿ�����) = 1 Then
                        chkSpecial.Enabled = False
                        Exit Do
                    End If
                    mrsData.MoveNext
                Loop
            End If
            txtName.SetFocus
            UnitReportControl.Tag = ""
            mblnChange = True
            
        Case conMenu_Edit_ModifyParent ' "�޸ķ���(&U)"
            fraInfo.Tag = ""
            fraUnit.Tag = "�޸�"
            txtName.Enabled = True
            txtDays.Enabled = True
            chkSpecial.Visible = (m����ID = 0)
            chkSpecial.Enabled = (m����ID = 0)
            If m����ID = 0 Then
                mrsData.Filter = "������=0 and �������<>" & Val(Split(UnitReportControl.FocusedRow.Childs(0).Record(COL_�������).Record.Tag, "-")(0))
                Do While Not mrsData.EOF
                    If Val("" & mrsData!�Ƿ�����) = 1 Then
                        chkSpecial.Enabled = False
                        Exit Do
                    End If
                    mrsData.MoveNext
                Loop
            End If
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            txtName.SetFocus
            UnitReportControl.Tag = UnitReportControl.FocusedRow.Index
            mblnChange = True

        Case conMenu_Edit_DeleteParent '"ɾ������(&E)"
            If UnitReportControl.FocusedRow Is Nothing Then Exit Sub
            
            If MsgBox("��ȷ��Ҫɾ��������" & Split(mstr��������, "-")(1) & "����Ƿ��ࡾ" & UnitReportControl.FocusedRow.Childs(0).Record(COL_�������).GroupCaption & "������Ϣ��?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Childs(0).Record(COL_�������).Record.Tag
            
            mUnit.����ID = CInt(Split(strTemp, "-")(2))
            mUnit.������� = CInt(Split(strTemp, "-")(0))
            mUnit.������ = 0
            
            '�����������ݸò����Ƿ�����ʹ��
            If CheckUseUnit(mUnit.����ID, mUnit.�������, mUnit.������) = True Then Exit Sub
            
            strSql = "Zl_�����������_Delete(" & mUnit.����ID & "," & mUnit.������� & "," & mUnit.������ & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            Call RefreshData(-1)
            
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Edit_Save     '*����
            picBack.Visible = False
            cmdImage.Enabled = True
            Call SaveData
        Case conMenu_Edit_Reuse    '*ȡ��
            '��¼����ѡ�еı�ע
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '���ڿ������¶�λ
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_�������).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '���ڿ������¶�λ
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_�������).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            Else
                If UnitReportControl.Tag <> "" Then
                    If Not UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then
                        If Not UnitReportControl.Rows(UnitReportControl.Tag).GroupRow And UnitReportControl.Rows(UnitReportControl.Tag).Childs.Count = 0 Then
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTemp = UnitReportControl.Rows(UnitReportControl.Tag).Record(COL_�������).Record.Tag
                        Else
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTmp1 = UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_�������).Record.Tag
                            strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                        End If
                    End If
                End If
            End If
            picBack.Visible = False
            cmdImage.Enabled = True
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_View_Refresh  'ˢ��
            '��¼����ѡ�еı�ע
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '���ڿ������¶�λ
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_�������).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '���ڿ������¶�λ
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_�������).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            End If
            
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hwnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)

        Case conMenu_Help_Web_Mail '����Email
            Call zlMailTo(Me.hwnd)
            
        Case conMenu_Help_Help        '*��������(&H)
             Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit        '*�˳�(&X)
            Unload Me
    End Select
    
    Call RefreshStateInfo
    cbsMain.RecalcLayout
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = (UnitReportControl.Records.Count > 0)
        Case conMenu_Edit_NewItem   '*����(&A)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange
                End If
            Else
                Control.Enabled = mLngCount > 0
            End If
        Case conMenu_Edit_Modify      '*�޸�(&M)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_�������).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Delete      '*ɾ��(&D)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_�������).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        
        Case conMenu_Edit_NewParent '*��������
            Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
            If Control.Enabled = True Then
                Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
            Else
                If UnitReportControl.Rows.Count > 0 Then
                    Control.Enabled = Not mblnChange
                Else
                    Control.Enabled = True And Not mblnChange
                End If
            End If
             
        Case conMenu_Edit_ModifyParent ' "�޸ķ���(&U)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_DeleteParent '"ɾ������(&E)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_Save     '*����
            Control.Enabled = mblnChange
        Case conMenu_Edit_Reuse     '*ȡ��
            Control.Enabled = mblnChange
        Case conMenu_View_Refresh '*ˢ��
            Control.Enabled = Not mblnChange
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsMain(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar
            Control.Checked = Me.stbThis.Visible
    End Select
    
    cboUnit.Enabled = Not mblnChange
    fraUd.Enabled = Not mblnChange
    
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub lblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ShowSelect(Index)
End Sub

Private Sub picBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub picIcon_KeyPress(Index As Integer, KeyAscii As Integer)
    picBack_KeyPress KeyAscii
End Sub

Private Sub pic���_KeyPress(KeyAscii As Integer)
    picBack_KeyPress KeyAscii
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If KeyAscii > 45 And KeyAscii < 58 Then
            If KeyAscii = 46 Then
                If Len(txtDays.Text) = 0 Then
                    KeyAscii = 0
                Else
                    If InStr(1, txtDays.Text, ".") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Else
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtInfo_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
    txtInfo.SelStart = Len(txtInfo.Text)
    Call zlControl.TxtSelAll(txtInfo)
End Sub


Private Sub txtInfo_Change()
    If mblnChange = False Then Exit Sub
    
    If fraInfo.Tag = "�޸�" Then
        With UnitReportControl.FocusedRow.Record(COL_˵��)
            .Value = txtInfo.Text
        End With
        UnitReportControl.Populate
    End If
    
    '�ж�����Ա�Ƿ��ֹ�¼���޸��˱�ע˵��
    If lblSet(8).Tag <> "" And lblSet(8).Tag <> Trim(txtInfo.Text) And Trim(txtInfo.Text) <> cmdImage.Tag Then
        txtInfo.Tag = "�ı�"
    End If
    
    If imaCustom.ComboItems.Count > 0 Then cmdImage.Tag = imaCustom.Text
End Sub

Private Sub txtInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtInfo.Text) <> "" Then
            txtInfo.Tag = "�ı�"
        End If
    Else
        If Chr(KeyCode) = "'" Or Chr(KeyCode) = "|" Then KeyCode = 0
    End If
End Sub


Private Sub txtName_Change()
    Dim i As Integer
    Dim lngPreIdx As Long
    Dim strTemp As String, str��� As String
    If mblnChange = False Then Exit Sub
    
    If fraUnit.Tag = "�޸�" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            lngPreIdx = .Index
            strTemp = .Childs(0).Record(COL_�������).Record.Tag
            str��� = Split(strTemp, "-")(0) & "-0-" & Split(strTemp, "-")(2) & "-" & Split(strTemp, "-")(3)
            
            For i = 0 To .Childs.Count - 1
                .Childs(i).Record(COL_�������).GroupCaption = "���飺" & Split(strTemp, "-")(0) & "-" & Replace(txtName.Text, "'", "")
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = Len(txtName.Text)
    Call zlControl.TxtSelAll(txtName)
End Sub

Private Sub txtDays_GotFocus()
    txtDays.SelStart = Len(txtDays.Text)
    Call zlControl.TxtSelAll(txtDays)
End Sub

Private Sub txtDays_Change()
    Dim i As Integer
    If mblnChange = False Then Exit Sub
    '���ķ�������ʱ���ӷ���ͬ������
    If fraUnit.Tag = "�޸�" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            For i = 0 To .Childs.Count - 1
                If Val(Split(.Childs(i).Record(COL_�������).Record.Tag, "-")(1)) = 0 Then
                    .Childs(i).Record(COL_��Ч����).Value = ""

                Else
                    .Childs(i).Record(COL_��Ч����).Value = IIF(txtDays.Text = "", 0, txtDays.Text)
                End If
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr(KeyCode) = "'" Then KeyCode = 0
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub UnitReportControl_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)
    Call Arrange(Column.Index)
End Sub

Public Sub Arrange(Column As Long)
    UnitReportControl.SortOrder.DeleteAll
    UnitReportControl.SortOrder.Add UnitReportControl.Columns.Find(Column)
    UnitReportControl.SortOrder(0).SortAscending = Not UnitReportControl.SortOrder(0).SortAscending
    UnitReportControl.Populate
End Sub


Private Sub UnitReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         If Not (UnitReportControl.FocusedRow Is Nothing) Then
            If Not UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count = 0 Then
              Call UnitReportControl_RowDblClick(UnitReportControl.FocusedRow, UnitReportControl.FocusedRow.Record.Item(COL_�������))
            End If
        End If
    End If
End Sub

Private Sub UnitReportControl_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
'����:�����ʼ��˵�
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub UnitReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not (Row Is Nothing) Then
        If Not Row.GroupRow And Row.Childs.Count = 0 And Val(Split(Row.Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_Modify, True, True))
        Else
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_ModifyParent, True, True))
        End If
    End If
End Sub


Private Sub UnitReportControl_SelectionChanged()
'-------------------------------------------------
'����:����ReportControl��ѡ���У���ȡ��Ӧ�Ĳ���������Ϣ
'
'--------------------------------------------------
    Dim i As Integer
    
    txtInfo.Text = "": txtInfo.Tag = "": lblSet(7).Tag = "": lblSet(8).Tag = "": imaCustom.Text = "": imaCustom.Tag = ""
    lblSet(9).Tag = "": cbo���.Tag = "": lblSet(1).Tag = "": txtName.Text = "": lblSet(4).Tag = "": txtDays.Text = "": chkSpecial.Value = 0
    
    On Error GoTo ErrHand
        With UnitReportControl.FocusedRow
            If Not UnitReportControl.FocusedRow Is Nothing Then
                If Not .GroupRow And .Childs.Count = 0 Then
                    If Val(Split(.Record(COL_�������).Record.Tag, "-")(1)) <> 0 Then
                        cbo���.ListIndex = SetCboIndex(cbo���, Val(Split(.Record(COL_�������).Record.Tag, "-")(0)))
                        lblSet(9).Tag = .Record(COL_ԭʼ����).Value
                        lblSet(8).Tag = .Record(COL_˵��).Value
                        txtInfo.Text = .Record(COL_˵��).Value
                        lblSet(7).Tag = IIF(Val(.Record(COL_��ע).Icon) <= 0, "0", Val(.Record(COL_��ע).Icon))
                        If lblSet(7).Tag >= mlngImgIndex Then
                            imaCustom.ComboItems(Val(lblSet(7).Tag) - mlngImgIndex + 1).Selected = True
                        End If
                        Call SetControlEnable(fraInfo.Tag <> "")
                        Call SetFraResize
                    Else
                        UnitReportControl.FocusedRow = UnitReportControl.Rows(UnitReportControl.FocusedRow.Index - 1)
                    End If
                Else
                    lblSet(1).Tag = Split(.Childs(0).Record(COL_�������).GroupCaption, "-")(1)
                    txtName.Text = lblSet(1).Tag
                    lblSet(4).Tag = Val(.Childs(0).Record(COL_��Ч����).Value)
                    txtDays.Text = lblSet(4).Tag
                    
                    txtName.Enabled = fraUnit.Tag <> ""
                    txtDays.Enabled = fraUnit.Tag <> ""
                    
                    txtName.BackColor = IIF(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    txtDays.BackColor = IIF(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    chkSpecial.Visible = (m����ID = 0)
                    chkSpecial.Enabled = fraUnit.Tag <> "" And (m����ID = 0)
                    If m����ID = 0 Then
                        chkSpecial.Value = Val(.Childs(0).Record(COL_�Ƿ�����).Value)
                    End If
                    chkSpecial.Tag = chkSpecial.Value
                    
                    Call SetFraResize(True)
                End If
            End If
        End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCboIndex(ByVal objCbo As Object, ByVal intItemData As Integer) As Integer
'------------------------------------------------------------------------
'����:����itemdata��ֵ��ȡcbo��Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim intIndex As Integer
    
    intIndex = -1
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           intIndex = i
           Exit For
        End If
    Next i
    
    SetCboIndex = intIndex
End Function

Private Function GetCboText(ByVal objCbo As Object, ByVal intItemData As Integer) As String
'------------------------------------------------------------------------
'����:����itemdata��ֵ��ȡcbo��Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim strText As String
    
    strText = ""
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           strText = objCbo.Text
           Exit For
        End If
    Next i
    
    GetCboText = strText
End Function

Private Function CheckChange() As Boolean
'-----------------------------------------------------
'����:�޸�ʱ��������Ƿ����仯
'-----------------------------------------------------
    Dim blnChage As Boolean
    If fraInfo.Tag = "�޸�" Then
        If Val(lblSet(9).Tag) <> cbo���.ListIndex Or lblSet(8).Tag <> txtInfo.Text Or _
            Val(lblSet(7).Tag) <> imaCustom.SelectedItem.Index - 1 + mlngImgIndex Then
            blnChage = True
        End If
    ElseIf fraUnit.Tag = "�޸�" Then
        If lblSet(1).Tag <> txtName.Text Or lblSet(4).Tag <> txtDays.Text Or Val(chkSpecial.Tag) <> Val(chkSpecial.Value) Then
            blnChage = True
        End If
    End If
    CheckChange = blnChage
End Function

Private Function CheckUseUnit(ByVal lngUnitID As Long, ByVal lngSubjectID As Long, ByVal lngTracerID As Long) As Boolean
'----------------------------------------------------------
'���ܣ����ı�������Ƿ�����ʹ��
'������lngUnitId ����ID��lngSubjectID ������� ��lngTracerID ������
'----------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrue As Boolean
    Dim strSql
    On Error GoTo ErrHand
    
    If lngTracerID <> 0 Then
        strSql = "Select 1 From ������Ǽ�¼" & _
            "   WHERE  " & IIF(lngUnitID = 0, " ���ⲡ��Id IS NULL ", " ����Id=[1] ") & " and �������=[2] and ������=[3] And RowNum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "������Ǽ�¼", lngUnitID, lngSubjectID, lngTracerID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "�ñ������Ŀǰ�Ĳ�������ʹ��,��ȡ��ʹ�ú���ɾ��.", vbInformation, gstrSysName
        End If
    Else
        strSql = _
            " SELECT 1" & vbNewLine & _
            " FROM ����������� A,������Ǽ�¼ B" & vbNewLine & _
            " WHERE  " & IIF(lngUnitID = 0, " B.���ⲡ��Id IS NULL ", " A.����ID=B.����ID ") & " And A.�������=B.������� And " & IIF(lngUnitID = 0, " A.����ID IS NULL ", " A.����ID=[1] ") & " And A.�������=[2]  " & vbNewLine & _
            " And RowNum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "������Ǽ�¼", lngUnitID, lngSubjectID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "�ñ�Ƿ����µı������Ŀǰ�Ĳ�������ʹ��,��ȡ��ʹ�ú���ɾ��.", vbInformation, gstrSysName
        End If
    End If
    CheckUseUnit = blnTrue
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewPreID(ByVal lng����id As Long, ByVal lngPreVId As Long) As Long
'--------------------------------------------------------------------
'����:��ȡĳ����ĳ�����µı�����
'����:lng����ID������ID �� lngPreVID ���������
'--------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim lngPreID As Long, i As Integer
    Dim arrPreID, blnFind As Boolean
    On Error GoTo ErrHand
    arrPreID = Array()
    strSql = _
        " select ������" & _
        " From �����������" & _
        " Where " & IIF(lng����id = 0, " ����Id IS NULL ", " ����Id=[1] ") & " and �������=[2] order by ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����id, lngPreVId)
    Do While Not rsTemp.EOF
        ReDim Preserve arrPreID(UBound(arrPreID) + 1)
        arrPreID(UBound(arrPreID)) = Val(rsTemp!������ & "")
        rsTemp.MoveNext
    Loop
    For i = 0 To UBound(arrPreID)
        If Val(arrPreID(i)) > i + 1 Then
            lngPreID = i + 1
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lngPreID = i + 1
    End If
    
    GetNewPreID = lngPreID
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewSubjectId(ByVal lng����id As Long) As Long
'------------------------------------------------------------------------
'����:������ע����ʱ����ȡĳ���������������������
'------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim lngSubjectID As Long
    Dim arrSubJectID, i As Integer
    Dim blnFind As Boolean
    
    On Error GoTo ErrHand:
    strSql = _
        " select �������,˵�� from �����������" & _
        " where " & IIF(lng����id = 0, " ����Id IS NULL ", " ����Id=[1] ") & " And ������=0 Order by �������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����������", lng����id)
    
    arrSubJectID = Array()
    With rsTmp
        Do While Not .EOF
            ReDim Preserve arrSubJectID(UBound(arrSubJectID) + 1)
            arrSubJectID(UBound(arrSubJectID)) = Val("" & !�������)
            .MoveNext
        Loop
    End With
    
    For i = 0 To UBound(arrSubJectID)
        If Val(arrSubJectID(i)) > i + 1 Then
            lngSubjectID = i + 1
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lngSubjectID = i + 1
    End If

    GetNewSubjectId = lngSubjectID
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsEqualInfo(strName As String, Optional ByVal blnSubject As Boolean, Optional ByVal strKey As String = "") As Boolean
'ͬһ�����±����Ҫ˵�������ظ�
    Dim StrInfo As String
    Dim blnAdd As Boolean
    On Error GoTo ErrHand
    If blnSubject = True Then
        mrsData.Filter = "������=0"
    Else
        mrsData.Filter = "������>0"
    End If
    Do While Not mrsData.EOF
        blnAdd = False
        If strKey = "" Then
            blnAdd = True
        Else
            If blnSubject = True Then
                If "" & mrsData!������� <> strKey Then blnAdd = True
            Else
                If "" & mrsData!������� & "-" & "" & mrsData!������ <> strKey Then blnAdd = True
            End If
        End If
        If blnAdd = True Then StrInfo = StrInfo & "'" & mrsData!˵��
        mrsData.MoveNext
    Loop
    If Left(StrInfo, 1) = "'" Then StrInfo = Mid(StrInfo, 2)
    '����Ƿ��������Ƿ��ظ�
    If InStr(1, "'" & StrInfo & "'", "'" & strName & "'") <> 0 Then
        If blnSubject = True Then
            MsgBox "�˱�������Ѿ�����,��������д��", vbInformation, gstrSysName
        Else
            MsgBox "�˱��˵���Ѿ�����,��������д��", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    IsEqualInfo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckUnitSubject(ByVal lng����id As Long) As Long
'---------------------------------------------------
'����:����Ƿ���ڱ�ע��������,��������ʾ����Ա��������
'---------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo ErrHand
    
    If lng����id = 0 Then '����ͼ��
        strSql = " select �������,˵�� from �����������  where ����Id is null and  ������=0"
    Else
        strSql = " select �������,˵�� from �����������  where ����Id=[1] and  ������=0"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�����������", lng����id)
    
    cbo���.Clear
    With rsTmp
        Do While Not .EOF
            cbo���.AddItem zlCommFun.Nvl(!˵��, "���Ա�ע" & zlCommFun.Nvl(!�������))
            cbo���.ItemData(cbo���.NewIndex) = Val(zlCommFun.Nvl(!�������))
            If cbo���.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cbo���.hwnd, cbo���.NewIndex)
            End If
        .MoveNext
        Loop
    End With
                
    CheckUnitSubject = rsTmp.RecordCount

    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim strGroupCaption As String
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '�����и���
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = 100 * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                If rptCol.Caption = "��ע" Then rptCol.Width = 10
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '�����и���
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    If rptRow.Record(.ColData(lngCol)).GroupCaption <> "" Then
                        strGroupCaption = Split(rptRow.Record(.ColData(lngCol)).GroupCaption, "��")(1)
                    Else
                        strGroupCaption = rptRow.Record(.ColData(lngCol)).GroupCaption
                    End If
                    .TextMatrix(lngRow, lngCol) = IIF(.TextMatrix(0, lngCol) = "�������", strGroupCaption, rptRow.Record(.ColData(lngCol)).Value)
                    If rptRow.Record(.ColData(lngCol)).Icon > 0 Then
                        '.CellPicture = zlCommFun.GetPaitSignImageList(0).ListImages(rptRow.Record(.ColData(lngCol)).Icon).Picture
                    End If
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Private Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    strTmp = strArray(0)
    lngCount = UBound(strArray) + 1
    GoTo OkHand
InitHand:
    lngCount = 1
OkHand:
    ReDim Preserve strArray(0 To lngCount)
    ReDimArray = lngCount
End Function

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIF(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub
