VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAppforBillSelSample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ѡ��걾"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3840
      ScaleHeight     =   855
      ScaleWidth      =   1485
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   840
      Width           =   1485
      Begin VSFlex8Ctl.VSFlexGrid vsfItem 
         Height          =   500
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   600
         _cx             =   1058
         _cy             =   882
         Appearance      =   0
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
         BackColorSel    =   16706793
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   90
      ScaleHeight     =   1665
      ScaleWidth      =   7905
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Width           =   7935
      Begin VB.PictureBox picExceDept 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3240
         MouseIcon       =   "frmAppforBillSelSample.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBillSelSample.frx":030A
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1260
         Width           =   255
      End
      Begin VB.TextBox txtExecDept 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1230
         Width           =   1725
      End
      Begin VB.PictureBox picGetSampleType 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   7290
         MouseIcon       =   "frmAppforBillSelSample.frx":0D0C
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBillSelSample.frx":1016
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox txtGetSampleType 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5490
         TabIndex        =   15
         Top             =   510
         Width           =   1725
      End
      Begin VB.PictureBox picSampleDept 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   3270
         MouseIcon       =   "frmAppforBillSelSample.frx":1A18
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBillSelSample.frx":1D22
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   540
         Width           =   255
      End
      Begin VB.TextBox txtGetSampleDept 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   12
         Top             =   510
         Width           =   1725
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   3240
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line1 
         X1              =   5430
         X2              =   7230
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line6 
         X1              =   1440
         X2              =   3240
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Label lblExecDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label lblSampleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ɼ���ʽ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4350
         TabIndex        =   14
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblSampleDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   540
         Width           =   1125
      End
   End
   Begin VB.PictureBox picAppend 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   1695
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   780
      Width           =   1725
      Begin VSFlex8Ctl.VSFlexGrid vsfAppend 
         Height          =   465
         Left            =   570
         TabIndex        =   22
         Top             =   300
         Width           =   825
         _cx             =   1455
         _cy             =   820
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
      Begin RichTextLib.RichTextBox rtfAppend 
         Height          =   465
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   820
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmAppforBillSelSample.frx":2724
      End
      Begin VB.CommandButton cmdҽ������ 
         Caption         =   "��"
         Height          =   270
         Left            =   30
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   450
         Width           =   285
      End
      Begin VB.CommandButton cmd�������� 
         Height          =   300
         Left            =   360
         Picture         =   "frmAppforBillSelSample.frx":27C1
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "����ǰ��������Ϊ��������(Ctrl+D)"
         Top             =   480
         Width           =   315
      End
      Begin VB.TextBox cboҽ������ 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   780
         MaxLength       =   100
         TabIndex        =   7
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   810
         TabIndex        =   5
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.PictureBox picSample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1920
      ScaleHeight     =   855
      ScaleWidth      =   1485
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   1485
      Begin VB.CheckBox chkShowall 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ʾ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1305
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFList 
         Height          =   585
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Width           =   645
         _cx             =   1138
         _cy             =   1032
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
   Begin XtremeSuiteControls.TabControl TabMain 
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   8175
      _Version        =   589884
      _ExtentX        =   14420
      _ExtentY        =   7911
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppforBillSelSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnShow As Boolean                         '�����Ƿ���ʾ
Private mstrSample As String                        'ѡ��걾
Private mstrSampleNO As String                      '��Ŀ����
Private mrsAppend As New ADODB.Recordset            '���븽��
Private mstrAppend As String
Private mblnCancel As Boolean
Private mblnNull As Boolean

Private mlng����ID As Long
Private mvar����ID As Variant
Private mstrDiagnosis As String
Private mintӤ�� As Integer
Private mstrAdvItem As String
Private mintPatientType As Integer

Private mlngSampleDept As Long
Private mstrSampleDept As String
Private mlngSampleType As Long
Private mstrSampleType As String
Private mlngExcDept As Long
Private mstrExcDept As String
Private mstrEntrust As String

Private mstrSplieListTag As String                      '�ָ���
Private mstrSplieItemTag As String                      '�ָ���
Private mstrSplieColTag As String                       '�ָ���

Private mstrPosition As String
Private mlngPosition As Long
Private mstrRichText As String
Private mlngSelStart As Long
Private mlngAppForDeptID As Long                        '�������ID

Private mrsReference As ADODB.Recordset                   '�������Ŀ����ָ��Ĳο�
Private mlngGroupItemID As Long                     '�����ĿID
Private mstrSex As String                           '�Ա�
Private mstrAge As String                           '����
Private mstrAppForDept As String                    '�������
Private mstrRef As String                           '�ο�Ҫ��
Private mlngMaxWidth As Long                        '�������п��

Private Const FCONTROL = 8                  'ctrl��ϼ�

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
        Case ConMenu_Browse_Save                '����
            mblnCancel = False
            mstrSample = SelSampe
            mstrAppend = GetAppend
            mlngSampleDept = Val(txtGetSampleDept.Tag)
            mstrSampleDept = txtGetSampleDept.Text
            mlngSampleType = Val(txtGetSampleType.Tag)
            mstrSampleType = txtGetSampleType.Text
            mlngExcDept = txtExecDept.Tag
            mstrExcDept = txtExecDept.Text
            mstrEntrust = cboҽ������.Text
            
            '�����б�����δ��,���˳�
            If Not mblnNull Then Unload Me
        Case ConMenu_Browse_Cancel              'ȡ��
            mblnCancel = True
            Unload Me
    End Select
End Sub

Private Sub chkShowall_Click()
    If chkShowall.value = 1 Then
        Call ReadData(mstrSampleNO, 1)
    Else
        Call ReadData(mstrSampleNO, 0)
    End If
End Sub


Private Sub cmd��������_Click()
          Dim strSQL As String
          Dim rsTmp As Recordset
          Dim int���� As Integer
          
1         On Error GoTo cmd��������_Click_Error

2         If Trim(cboҽ������.Text) = "" Then
3             MsgBox "�������������ݡ�", vbInformation, gSysInfo.ShortName
4             If cboҽ������.Enabled Then cboҽ������.SetFocus
5             Exit Sub
6         End If

7         strSQL = "Select 1 From �������� Where ����=[1] And (��Ա=[2] Or ��Ա is null)"
8         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, Trim(cboҽ������.Text), gUserInfo.Name)
9         If rsTmp.RecordCount > 0 Then
10            MsgBox "�����������Ѿ��ڳ��������С�", vbInformation, gSysInfo.ShortName
11            If cboҽ������.Enabled Then cboҽ������.SetFocus
12            Exit Sub
13        End If
          
14        strSQL = zlGetSymbol(cboҽ������.Text, CByte(int����))
15        strSQL = "zl_��������_Insert('" & Replace(cboҽ������.Text, "'", "''") & "','" & strSQL & "','" & gUserInfo.Name & "')"
16        Call ComExecuteProc(Sel_His_DB, strSQL, Me.Caption)
          
17        AddComboItem cboҽ������.hWnd, CB_ADDSTRING, 0, cboҽ������.Text
18        MsgBox "������Ϊ�������С�", vbInformation, gSysInfo.ShortName
19        If cboҽ������.Enabled Then cboҽ������.SetFocus


20        Exit Sub
cmd��������_Click_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(cmd��������_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear

End Sub

Private Sub cmdҽ������_Click()
    If ReasonSelect("", 2) Then Exit Sub
    cboҽ������.Tag = "1"
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        Call ReadData(mstrSampleNO, 0)
        Call ReadItemData(mstrSampleNO)
        Call Init���븽��
        mblnShow = True
    End If
End Sub

Private Sub Form_Load()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Save, "����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Cancel, "ȡ��")
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, vbKeyS, ConMenu_Appfro_ModifyItem
        .Add 0, vbKeyEscape, ConMenu_Browse_Cancel
    End With
    
    With Me.TabMain
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = True
        .PaintManager.BoldSelected = True
        .InsertItem 0, "��������", picAppend.hWnd, 0
        .InsertItem 1, "�걾", picSample.hWnd, 0
        .InsertItem 2, "ִ��ѡ��", picDept.hWnd, 0
        .InsertItem 3, "��Ŀ��ϸ", picItem.hWnd, 0
    End With
    With vsfList
        .GridLines = flexGridNone
        .Rows = 0
        .Cols = 4
        .ColKey(0) = "�걾1": .ColWidth(0) = 1900
        .ColKey(1) = "�걾2": .ColWidth(1) = 1900
        .ColKey(2) = "�걾3": .ColWidth(2) = 1900
        .ColKey(3) = "�걾4": .ColWidth(3) = 1900
    End With
    
    '�ָ���
    mstrSplieColTag = "<Split A>"
    mstrSplieItemTag = "<Split B>"
    mstrSplieListTag = "<Split C>"
End Sub

Public Function ShowMe(objFrm As Object, strItemNO As String, strSample As String, _
                        lng����ID As Long, var����ID As Variant, strDiagnosis As String, intӤ�� As Integer, _
                        intPatientType As Integer, strAdvItem As String, strAppend As String, _
                        lngSampleDept As Long, strSampleDept As String, lngSampleType As Long, strSampleType As String, _
                        lngExcDept As Long, strExcDept As String, strEntrust As String, ByVal lngAppForDeptID As Long, _
                        ByVal lngGroupItemID As Long, ByVal strSex As String, ByVal strAge As String, _
                        ByVal strAppForDept As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����       ��ʾѡ��걾���ڣ������뵱ǰ�걾
    '����       strSample ��ǰ�걾
    '����       ѡ��ı걾
    '           lngGroupItemID      �����ĿID
    '           strSex              �Ա�
    '           strAge              ����
    '           strAppForDept       �������
    '           blnAllHave          �Ƿ����вο�Ҫ�ض��Ѿ���������
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mstrSampleNO = strItemNO
    mstrSample = strSample
    mlng����ID = lng����ID
    mvar����ID = var����ID
    mstrDiagnosis = strDiagnosis
    mintӤ�� = intӤ��
    mstrAdvItem = strAdvItem
    mstrAppend = strAppend
    mintPatientType = intPatientType
    mlngGroupItemID = lngGroupItemID
    
    
    mlngSampleDept = lngSampleDept
    mstrSampleDept = strSampleDept
    mlngSampleType = lngSampleType
    mstrSampleType = strSampleType
    mlngExcDept = lngExcDept
    mstrExcDept = strExcDept
    mstrEntrust = strEntrust
    mlngAppForDeptID = lngAppForDeptID
    
    mstrSex = strSex
    mstrAge = strAge
    mstrAppForDept = strAppForDept
    
    Me.Show vbModal, objFrm
    If mblnCancel = False Then
        ShowMe = mstrSample & mstrSplieColTag & mstrAppend & mstrSplieColTag & mlngSampleDept & mstrSplieColTag & mstrSampleDept & mstrSplieColTag & _
                                mlngSampleType & mstrSplieColTag & mstrSampleType & mstrSplieColTag & mlngExcDept & mstrSplieColTag & mstrExcDept & _
                                mstrSplieColTag & mstrEntrust
    Else
        ShowMe = ""
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    Set mrsAppend = Nothing
    mlngAppForDeptID = 0
    mvar����ID = vbNullString
End Sub

Private Sub ReadData(strNO As String, intType As Integer)
          '''''''''''''''''''''''''''''''''''''''''''''''''
          '����       ����걾����
          '����       intType 0=�����Ʊ������ 1=�ҵ�����
          '˵��       ��������Ʊ���û���ҵ���¼���ٲ�������
          '''''''''''''''''''''''''''''''''''''''''''''''''
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intCol As Integer
          Dim blnExit As Boolean
          
1         On Error GoTo ReadData_Error
          
2         If intType = 0 Then
3             If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
4                 If gUserInfo.NodeNo <> "-" Then
5                     strSQL = "Select Distinct d.��ʾ���� As �걾����" & vbNewLine & _
                              "From ���������Ŀ A, �������ָ�� B, ����ָ��ο���Χ C, ����ο�Ҫ�ض��� D, ����ָ��ο�Ҫ�� E" & vbNewLine & _
                              "Where a.Id = b.���id And b.��Ŀid = c.ָ��id And c.Id = d.�ο�id And d.Ҫ��id = e.Id And e.Ҫ���� = '�걾����' And d.��ʾ���� Is Not Null And" & vbNewLine & _
                              "      a.���Ʊ��� = [1] And (a.վ�� = [2] Or a.վ�� Is Null)"
6                 Else
7                     strSQL = "Select Distinct d.��ʾ���� As �걾����" & vbNewLine & _
                              "From ���������Ŀ A, �������ָ�� B, ����ָ��ο���Χ C, ����ο�Ҫ�ض��� D, ����ָ��ο�Ҫ�� E" & vbNewLine & _
                              "Where a.Id = b.���id And b.��Ŀid = c.ָ��id And c.Id = d.�ο�id And d.Ҫ��id = e.Id And e.Ҫ���� = '�걾����' And d.��ʾ���� Is Not Null And" & vbNewLine & _
                              "      a.���Ʊ��� = [1]"
8                 End If
9             Else
10                If gUserInfo.NodeNo <> "-" Then
11                    strSQL = "select distinct �걾���� from ���������Ŀ a,�������ָ�� b,����ָ��ο� c" & vbNewLine & _
                               "where a.id = b.���id and b.��Ŀid = c.��Ŀid and a.���Ʊ��� = [1] and (a.վ��=[2] or a.վ�� is null) and c.�걾���� is not null "
12                Else
13                    strSQL = "select distinct �걾���� from ���������Ŀ a,�������ָ�� b,����ָ��ο� c" & vbNewLine & _
                               "where a.id = b.���id and b.��Ŀid = c.��Ŀid and a.���Ʊ��� = [1] and c.�걾���� is not null "
14                End If
15            End If
16            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����", strNO, gUserInfo.NodeNo)
              
17            If rsTmp.RecordCount = 0 Then
18                strSQL = "select ���� �걾���� from ����걾����"
19                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����")
20            End If
21        Else
22            strSQL = "select ���� �걾���� from ����걾����"
23            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����")
24        End If
          
25        With Me.vsfList
26            .Clear
27            .Rows = 1
28            intCol = 0
29            Do Until rsTmp.EOF
30                .Row = .Rows - 1
31                .Col = intCol
32                .TextMatrix(.Row, .Col) = rsTmp("�걾����") & ""
33                If mstrSample = rsTmp("�걾����") & "" Then
34                    blnExit = True
35                    .Cell(flexcpChecked, .Row, intCol, .Row, intCol) = 1
36                Else
37                    .Cell(flexcpChecked, .Row, intCol, .Row, intCol) = 2
38                End If
39                If intCol = 3 Then
40                    .Rows = .Rows + 1
41                    intCol = 0
42                Else
43                    intCol = intCol + 1
44                End If
45                rsTmp.MoveNext
46            Loop
              
              'û��������ĿĬ�ϵı걾����ʱ�Զ�����
47            If Not blnExit And mstrSample <> "" Then
48                .Row = .Rows - 1
49                .Col = intCol
50                .TextMatrix(.Row, .Col) = mstrSample
51                .Cell(flexcpChecked, .Row, intCol, .Row, intCol) = 1
52            End If
53        End With
          
54        txtGetSampleDept.Tag = mlngSampleDept
55        txtGetSampleDept.Text = mstrSampleDept
          
56        txtGetSampleType.Tag = mlngSampleType
57        txtGetSampleType.Text = mstrSampleType
          
58        txtExecDept.Tag = mlngExcDept
59        txtExecDept.Text = mstrExcDept
          
60        cboҽ������.Text = mstrEntrust


61        Exit Sub
ReadData_Error:
62        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(ReadData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
63        Err.Clear
End Sub

Private Sub ReadItemData(strNO As String)
          Dim rsTmp As ADODB.Recordset
          Dim rsBH As ADODB.Recordset
          Dim strSQL As String
          Dim intType As Integer

1         On Error GoTo ReadItemData_Error
          
2         If gUserInfo.NodeNo <> "-" Then
3             strSQL = " select ΢�������� from ���������Ŀ where ���Ʊ��� = [1] and (վ��=[2] or վ�� is null)"
4         Else
5             strSQL = " select ΢�������� from ���������Ŀ where ���Ʊ��� = [1]"
6         End If
7         Set rsBH = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����", strNO, gUserInfo.NodeNo)
8         Do Until rsBH.EOF
9             intType = Val(rsBH("΢��������") & "")
10            rsBH.MoveNext
11        Loop
12        If intType = 1 Then
13            If gUserInfo.NodeNo <> "-" Then
14                 strSQL = "Select Distinct c.���� ��Ŀ����, c.������ ��Ŀ����, c. ϸ����� ��Ŀ���, c.Ĭ��ҩ��,c.Ĭ�Ϸ���" & vbNewLine & _
                           "   From ���������Ŀ A, �������ϸ�� B, ����ϸ����¼ C" & vbNewLine & _
                           "   Where a.Id = b.���id And b.ϸ��id = c.Id And a.���Ʊ��� = [1] and (a.վ��=[2] or a.վ�� is null)"
15            Else
16                strSQL = "Select Distinct c.���� ��Ŀ����, c.������ ��Ŀ����, c. ϸ����� ��Ŀ���, c.Ĭ��ҩ��,c.Ĭ�Ϸ���" & vbNewLine & _
                           "   From ���������Ŀ A, �������ϸ�� B, ����ϸ����¼ C" & vbNewLine & _
                           "   Where a.Id = b.���id And b.ϸ��id = c.Id And a.���Ʊ��� = [1] "
17            End If
18            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����", strNO, gUserInfo.NodeNo)
19            With vsfItem
20                .Rows = 1
21                .Cols = 6
22                .FixedRows = 1
                  
23                .ColKey(0) = "���": .ColWidth(.ColIndex("���")) = 500: .ColAlignment(.ColIndex("���")) = flexAlignCenterCenter
24                .ColKey(1) = "��Ŀ����": .ColWidth(.ColIndex("��Ŀ����")) = 2000: .ColAlignment(.ColIndex("��Ŀ����")) = flexAlignLeftCenter
25                .ColKey(2) = "��Ŀ����": .ColWidth(.ColIndex("��Ŀ����")) = 2000: .ColAlignment(.ColIndex("��Ŀ����")) = flexAlignLeftCenter
26                .ColKey(3) = "��Ŀ���": .ColWidth(.ColIndex("��Ŀ���")) = 2000: .ColAlignment(.ColIndex("��Ŀ���")) = flexAlignLeftCenter
27                .ColKey(4) = "Ĭ��ҩ��": .ColWidth(.ColIndex("Ĭ��ҩ��")) = 1200: .ColAlignment(.ColIndex("Ĭ��ҩ��")) = flexAlignLeftCenter
28                .ColKey(5) = "Ĭ�Ϸ���": .ColWidth(.ColIndex("Ĭ�Ϸ���")) = 1200: .ColAlignment(.ColIndex("Ĭ�Ϸ���")) = flexAlignLeftCenter
               
29                .TextMatrix(0, .ColIndex("���")) = "���"
30                .TextMatrix(0, .ColIndex("��Ŀ����")) = "��Ŀ����"
31                .TextMatrix(0, .ColIndex("��Ŀ����")) = "��Ŀ����"
32                .TextMatrix(0, .ColIndex("��Ŀ���")) = "��Ŀ���"
33                .TextMatrix(0, .ColIndex("Ĭ��ҩ��")) = "Ĭ��ҩ��"
34                .TextMatrix(0, .ColIndex("Ĭ�Ϸ���")) = "Ĭ�Ϸ���"
35                .Row = 0
          
36                Do Until rsTmp.EOF
37                    .Rows = .Rows + 1
38                    .TextMatrix(.Rows - 1, .ColIndex("���")) = .Rows - 1
39                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = rsTmp("��Ŀ����") & ""
40                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = rsTmp("��Ŀ����") & ""
41                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")) = rsTmp("��Ŀ���") & ""
42                    .TextMatrix(.Rows - 1, .ColIndex("Ĭ��ҩ��")) = rsTmp("Ĭ��ҩ��") & ""
43                    .TextMatrix(.Rows - 1, .ColIndex("Ĭ�Ϸ���")) = rsTmp("Ĭ�Ϸ���") & ""
44                    rsTmp.MoveNext
45                Loop
          
46            End With
47        Else
48            If gUserInfo.NodeNo <> "-" Then
49                strSQL = "Select Distinct c.ָ����� ��Ŀ����, c.������ ��Ŀ����, decode( c.�������,1,'����',2,'����',3,'�붨��') ��Ŀ���, c.��λ ��Ŀ��λ,c.�������" & vbNewLine & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C" & vbNewLine & _
                           "   Where a.Id = b.���id And b.��Ŀid = c.Id And a.���Ʊ��� = [1] and (a.վ��=[2] or a.վ�� is null) order by c.�������"
50            Else
51                strSQL = "Select Distinct c.ָ����� ��Ŀ����, c.������ ��Ŀ����, decode( c.�������,1,'����',2,'����',3,'�붨��') ��Ŀ���, c.��λ ��Ŀ��λ,c.�������" & vbNewLine & _
                           "   From ���������Ŀ A, �������ָ�� B, ����ָ�� C" & vbNewLine & _
                           "   Where a.Id = b.���id And b.��Ŀid = c.Id And a.���Ʊ��� = [1] order by c.�������"
52            End If
53            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����", strNO, gUserInfo.NodeNo)
54            With vsfItem
55                .Rows = 1
56                .Cols = 6
57                .FixedRows = 1
                  
58                .ColKey(0) = "���": .ColWidth(.ColIndex("���")) = 500: .ColAlignment(.ColIndex("���")) = flexAlignCenterCenter
59                .ColKey(1) = "��Ŀ����": .ColWidth(.ColIndex("��Ŀ����")) = 2000: .ColAlignment(.ColIndex("��Ŀ����")) = flexAlignLeftCenter
60                .ColKey(2) = "��Ŀ����": .ColWidth(.ColIndex("��Ŀ����")) = 2000: .ColAlignment(.ColIndex("��Ŀ����")) = flexAlignLeftCenter
61                .ColKey(3) = "��Ŀ���": .ColWidth(.ColIndex("��Ŀ���")) = 1500: .ColAlignment(.ColIndex("��Ŀ���")) = flexAlignLeftCenter
62                .ColKey(4) = "��Ŀ��λ": .ColWidth(.ColIndex("��Ŀ��λ")) = 1500: .ColAlignment(.ColIndex("��Ŀ��λ")) = flexAlignLeftCenter
63                .ColKey(5) = "�������": .ColWidth(.ColIndex("�������")) = 0: .ColAlignment(.ColIndex("�������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�������")) = True
               
64                .TextMatrix(0, .ColIndex("���")) = "���"
65                .TextMatrix(0, .ColIndex("��Ŀ����")) = "��Ŀ����"
66                .TextMatrix(0, .ColIndex("��Ŀ����")) = "��Ŀ����"
67                .TextMatrix(0, .ColIndex("��Ŀ���")) = "��Ŀ���"
68                .TextMatrix(0, .ColIndex("��Ŀ��λ")) = "��Ŀ��λ"
69                .TextMatrix(0, .ColIndex("�������")) = "�������"
70                .Row = 0
          
71                Do Until rsTmp.EOF
72                    .Rows = .Rows + 1
73                    .TextMatrix(.Rows - 1, .ColIndex("���")) = .Rows - 1
74                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = rsTmp("��Ŀ����") & ""
75                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = rsTmp("��Ŀ����") & ""
76                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ���")) = rsTmp("��Ŀ���") & ""
77                    .TextMatrix(.Rows - 1, .ColIndex("��Ŀ��λ")) = rsTmp("��Ŀ��λ") & ""
78                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsTmp("�������") & ""
79                    rsTmp.MoveNext
80                Loop
          
81            End With
82        End If


83        Exit Sub
ReadItemData_Error:
84        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(ReadItemData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
85        Err.Clear

End Sub


Private Sub picAppend_Resize()
    With rtfAppend
        .Top = 10
        .Left = 5
        .Height = Me.picAppend.ScaleHeight - cboҽ������.Height - 50 - 15
        .Width = Me.picAppend.ScaleWidth - 10
    End With
    With vsfAppend
        .Top = 10
        .Left = 5
        .Height = Me.picAppend.ScaleHeight - cboҽ������.Height - 50 - 15
        .Width = Me.picAppend.ScaleWidth - 10
    End With
    With lbl
        .Top = vsfAppend.Top + vsfAppend.Height + 25
        .Left = 20
    End With
    With cboҽ������
        .Top = lbl.Top
        .Left = lbl.Left + lbl.Width + 25
        .Width = picAppend.ScaleWidth - .Left - cmd��������.Width - 25
    End With
    With cmdҽ������
        .Top = lbl.Top
        .Left = cboҽ������.Left + cboҽ������.Width - .Width - 25
    End With
    With cmd��������
        .Top = lbl.Top
        .Left = cmdҽ������.Left + cmdҽ������.Width + 15
    End With
    With lbl
        .Top = .Top + 60
    End With
    With cmdҽ������
        .Top = .Top + 20
    End With
End Sub

Private Sub picExceDept_Click()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItem As String
          Dim astrItem() As String

1         On Error GoTo picExceDept_Click_Error

2         If gUserInfo.NodeNo <> "-" Then
3             strSQL = "select id,����,����,HIS���ű��� from ����С���¼ where վ��=[1] or վ�� is null"
4         Else
5             strSQL = "select id,����,����,HIS���ű��� from ����С���¼ "
6         End If
7         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "ѡ��ִ��С��", gUserInfo.NodeNo)
8         If rsTmp.RecordCount = 0 Then Exit Sub
9         strItem = SeletItemFromRsOld(Me, rsTmp, "")
10        If strItem <> "" Then
11            astrItem = Split(strItem, ",")
12            strSQL = "select id,���� from ���ű� where ���� = [1] "
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ѡ����", CStr(astrItem(3)))
14            If rsTmp.RecordCount > 0 Then
15                txtExecDept.Tag = rsTmp("ID")
16                txtExecDept.Text = rsTmp("����")
17            Else
18                MsgBox "С��<" & astrItem(2) & "û�к�HIS���Ҷ��룡", vbInformation, "С��ѡ��"
19            End If
20        End If


21        Exit Sub
picExceDept_Click_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(picExceDept_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
          
End Sub

Private Sub picGetSampleType_Click()
           '��ѡ��ɼ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
          
1         On Error GoTo picGetSampleType_Click_Error

2         Set rsTmp = GetSampleTypeRS()
         
          
3         strVal = SeletItemFromRsOld(Me, rsTmp, "")
4         astrItem = Split(strVal, ",")
5         If strVal <> "" Then
6             If UBound(astrItem) >= 2 Then
7                 If astrItem(2) <> "" Then
8                     astrItem = Split(strVal, ",")
9                     txtGetSampleType.Tag = astrItem(0)
10                    txtGetSampleType.Text = astrItem(2)
11                End If
12            End If
13        Else
14            If txtGetSampleType.Tag = "" Then
15                MsgBox "û��ѡ��ɼ���ʽ���ܱ��棡", vbInformation, "�ɼ���ʽѡ��"
16            End If
17        End If


18        Exit Sub
picGetSampleType_Click_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(picGetSampleType_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear
End Sub

Private Sub picItem_Resize()
    With vsfItem
        .Top = 5
        .Left = 5
        .Width = Me.picItem.ScaleWidth - 10
        .Height = Me.picItem.ScaleHeight - 50
    End With
End Sub

Private Sub picSample_Resize()
    With vsfList
        .Top = 5
        .Left = 5
        .Width = Me.picSample.ScaleWidth - 10
        .Height = Me.picSample.ScaleHeight - Me.chkShowall.Height - 50
    End With
    
    With chkShowall
        .Top = Me.vsfList.Top + Me.vsfList.Height + 25
        .Left = 50
    End With
End Sub

Private Sub picSampleDept_Click()
          '��ѡ��ɼ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
         
1         On Error GoTo picSampleDept_Click_Error

2         Set rsTmp = GetSampleDeptRS()
          
3         strVal = SeletItemFromRsOld(Me, rsTmp, "")
4         astrItem = Split(strVal, ",")
5         If strVal <> "" Then
6             If UBound(astrItem) >= 2 Then
7                 If astrItem(2) <> "" Then
8                     txtGetSampleDept.Tag = astrItem(0)
9                     txtGetSampleDept.Text = astrItem(2)
10                End If
11            End If
12        Else
13            If txtGetSampleDept.Tag = "" Then
14                MsgBox "û��ѡ��������Ҳ��ܱ��棡", vbInformation, "��������ѡ��"
15            End If
16        End If


17        Exit Sub
picSampleDept_Click_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(picSampleDept_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
19        Err.Clear
End Sub

Private Sub rtfAppend_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim varItem As Variant
          Dim i As Integer
1         On Error GoTo rtfAppend_KeyDown_Error

2         mlngPosition = Len(Left(mstrRichText, mlngSelStart))
3         If InStr(mstrPosition, ";") > 0 Then
4             varItem = Split(mstrPosition, ";")
5             For i = 0 To UBound(varItem)
6                 If varItem(i) <> "" Then
7                     If Split(varItem(i), ",")(1) = 1 Then 'ֻ����ɱ༭
8                         If i = UBound(varItem) Then
9                             If mlngPosition >= Val(InStr(mstrRichText, Split(varItem(i), ",")(0))) Then
10                                KeyCode = 0: Exit For
11                            End If
12                        Else
13                            If mlngPosition >= Val(InStr(mstrRichText, Split(varItem(i), ",")(0))) And mlngPosition <= Val(InStr(mstrRichText, Split(varItem(i + 1), ",")(0))) Then
14                                KeyCode = 0: Exit For
15                            End If
16                        End If
17                    End If
18                End If
19            Next
20        End If


21        Exit Sub
rtfAppend_KeyDown_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(rtfAppend_KeyDown)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
End Sub

Private Sub rtfAppend_SelChange()
    mstrRichText = rtfAppend.Text
    mlngSelStart = rtfAppend.SelStart
End Sub

Private Sub txtExecDept_GotFocus()
    Me.txtExecDept.SelStart = 0
    Me.txtExecDept.SelLength = Len(Me.txtExecDept)
End Sub

Private Sub txtGetSampleDept_DblClick()
          '��ѡ��ɼ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
         
1         On Error GoTo txtGetSampleDept_DblClick_Error

2         Set rsTmp = GetSampleDeptRS()
          
3         strVal = SeletItemFromRsOld(Me, rsTmp, "")
4         astrItem = Split(strVal, ",")
5         If strVal <> "" Then
6             If UBound(astrItem) >= 2 Then
7                 If astrItem(2) <> "" Then
                  
8                     txtGetSampleDept.Tag = astrItem(0)
9                     txtGetSampleDept.Text = astrItem(2)
10                End If
11            End If
12        Else
13            If txtGetSampleDept.Tag = "" Then
14                MsgBox "û��ѡ��������Ҳ��ܱ��棡", vbInformation, "��������ѡ��"
15            End If
16        End If


17        Exit Sub
txtGetSampleDept_DblClick_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(txtGetSampleDept_DblClick)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
19        Err.Clear
          
End Sub

Private Sub txtGetSampleDept_GotFocus()
    txtGetSampleDept.SelStart = 0
    txtGetSampleDept.SelLength = Len(txtGetSampleDept)
End Sub

Private Sub txtGetSampleDept_KeyPress(KeyAscii As Integer)
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
         
1         On Error GoTo txtGetSampleDept_KeyPress_Error

2         If KeyAscii = 13 Then
3             Set rsTmp = GetSampleDeptRS()
              
4             strVal = SeletItemFromRsOld(Me, rsTmp, txtGetSampleDept.Text)
5             astrItem = Split(strVal, ",")
6             If strVal = "" Then
7                 If UBound(astrItem) >= 2 Then
8                     If astrItem(2) <> "" Then
9                         astrItem = Split(strVal, ",")
10                        txtGetSampleDept.Tag = astrItem(0)
11                        txtGetSampleDept.Text = astrItem(2)
12                    End If
13                End If
14            Else
15                If txtGetSampleDept.Tag = "" Then
16                    MsgBox "û��ѡ��������Ҳ��ܱ��棡", vbInformation, "��������ѡ��"
17                Else
18                    txtGetSampleType.SetFocus
19                End If
20            End If
21        End If


22        Exit Sub
txtGetSampleDept_KeyPress_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(txtGetSampleDept_KeyPress)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear
End Sub

Private Sub txtGetSampleType_DblClick()
          '��ѡ��ɼ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
          
1         On Error GoTo txtGetSampleType_DblClick_Error

2         Set rsTmp = GetSampleTypeRS()
         
          
3         strVal = SeletItemFromRsOld(Me, rsTmp, "")
4         astrItem = Split(strVal, ",")
5         If strVal <> "" Then
6             If UBound(astrItem) >= 2 Then
7                 If astrItem(2) <> "" Then
8                     astrItem = Split(strVal, ",")
9                     txtGetSampleType.Tag = astrItem(0)
10                    txtGetSampleType.Text = astrItem(2)
11                End If
12            End If
13        Else
14            If txtGetSampleType.Tag = "" Then
15                MsgBox "û��ѡ��ɼ���ʽ���ܱ��棡", vbInformation, "�ɼ���ʽѡ��"
16            End If
17        End If


18        Exit Sub
txtGetSampleType_DblClick_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(txtGetSampleType_DblClick)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear

End Sub

Private Sub txtGetSampleType_GotFocus()
    Me.txtGetSampleType.SelStart = 0
    Me.txtGetSampleType.SelLength = Len(Me.txtGetSampleType)
End Sub

Private Sub txtGetSampleType_KeyPress(KeyAscii As Integer)
          Dim rsTmp As New ADODB.Recordset
          Dim strVal As String
          Dim astrItem() As String
          
1         On Error GoTo txtGetSampleType_KeyPress_Error

2         If KeyAscii = 13 Then
3             Set rsTmp = GetSampleTypeRS()
4             strVal = SeletItemFromRsOld(Me, rsTmp, txtGetSampleType.Text)
5             astrItem = Split(strVal, ",")
6             If strVal <> "" Then
7                 If UBound(astrItem) >= 2 Then
8                     If astrItem(2) <> "" Then
9                         astrItem = Split(strVal, ",")
10                        txtGetSampleType.Tag = astrItem(0)
11                        txtGetSampleType.Text = astrItem(2)
12                    End If
13                End If
14            Else
15                If txtGetSampleType.Tag = "" Then
16                    MsgBox "û��ѡ��ɼ���ʽ���ܱ��棡", vbInformation, "�ɼ���ʽѡ��"
17                Else
18                    txtExecDept.SetFocus
19                End If
20            End If
             
21        End If


22        Exit Sub
txtGetSampleType_KeyPress_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(txtGetSampleType_KeyPress)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear
End Sub

Private Sub vsfAppend_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strHave As String
          Dim strTime As String
          Dim strVal As String
          Dim i As Integer

1         On Error GoTo vsfAppend_CellButtonClick_Error

2         With Me.vsfAppend
3             If Row < 0 Or Col <> .ColIndex("����") Then Exit Sub

4             If .TextMatrix(Row, .ColIndex("��������")) = "�ַ�" Then   '�ַ���
5                 mrsReference.Filter = "Ҫ��ID=" & Val(.TextMatrix(Row, .ColIndex("Ҫ��ID"))) & " and ��������<>''"
6                 Set rsTmp = gobjHisDatabase.CopyNewRec(mrsReference, True)
                  'ȥ���ظ���ѡ��
7                 Do While Not mrsReference.EOF
8                     If InStr("<SP>" & strHave & "<SP>", "<SP>" & mrsReference("Ҫ��ID") & "<S>" & mrsReference("��������") & "<SP>") = 0 Then
9                         rsTmp.AddNew
10                        For i = 0 To rsTmp.Fields.Count - 1
11                            rsTmp.Fields(i).value = mrsReference(rsTmp.Fields(i).Name).value
12                        Next

13                        strHave = strHave & "<SP>" & mrsReference("Ҫ��ID") & "<S>" & mrsReference("��������")
14                    End If
15                    mrsReference.MoveNext
16                Loop
17            ElseIf .TextMatrix(Row, .ColIndex("��������")) = "ʱ���" Then  'ʱ�����
                  '���ɲ�ѯʱ����SQL
18                For i = 1 To 24
19                    strTime = strTime & "," & i & ":00"
20                Next
21                If strTime <> "" Then strTime = Mid(strTime, 2)
22                strSQL = "Select /*+cardinality(b,10)*/ '0' ID1,'0' ID2,Column_Value ��ʾ����,Column_Value �������� From Table(Cast(f_Str2list([1]) As zltools.t_strlist)) b"
23                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "ʱ���", strTime)
24            End If


25            strVal = SeletItemFromRs(Me, rsTmp, , , 3)
26            If strVal <> "" Then
27                .TextMatrix(Row, Col) = Split(strVal, "<SP2>")(2)
28                .Cell(flexcpData, Row, Col) = Split(strVal, "<SP2>")(3)
29            End If
30        End With


31        Exit Sub
vsfAppend_CellButtonClick_Error:
32        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(vsfAppend_CellButtonClick)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
33        Err.Clear
End Sub

Private Sub vsfAppend_Click()
          Dim lngRow As Long
          Dim lngCol As Long
          Dim blnNull As Boolean

1         On Error GoTo vsfAppend_Click_Error

2         With Me.vsfAppend
3             lngRow = .MouseRow
4             lngCol = .MouseCol

5             If lngRow < 0 Or lngCol <> .ColIndex("����") Then Exit Sub

6             If .TextMatrix(lngRow, .ColIndex("��������")) = "�ַ�" And .TextMatrix(lngRow, .ColIndex("����")) <> "�ٴ���ϣ�" Then    '�ַ���
                  '��Ϊ�յ�Ҫ����������
7                 mrsReference.Filter = "Ҫ��ID=" & Val(.TextMatrix(lngRow, .ColIndex("Ҫ��ID")))
8                 If mrsReference.RecordCount > 0 Then

9                     blnNull = True
10                Else
11                    blnNull = False
12                End If

13                If blnNull Then
14                    .ColComboList(lngCol) = "|..."
15                Else
16                    .ColComboList(lngCol) = ""
17                End If
18                .Editable = flexEDKbdMouse
19                .EditCell
20            Else
21                .ColComboList(lngCol) = ""
22                .Editable = flexEDNone
23            End If
24        End With


25        Exit Sub
vsfAppend_Click_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(vsfAppend_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear
End Sub



Private Sub VSFList_Click()
          Dim intRow As Integer, intCol As Integer
1         On Error GoTo VSFList_Click_Error

2         With Me.vsfList
3             If .MouseRow >= 0 And .MouseCol >= 0 Then
4                 intRow = .MouseRow
5                 intCol = .MouseCol
6                 If .TextMatrix(intRow, intCol) = "" Then Exit Sub
7                 If .Cell(flexcpChecked, intRow, intCol) = 1 Then
8                     .Cell(flexcpChecked, intRow, intCol) = 1
9                 Else
10                    Call ClearAllSel
11                    .Cell(flexcpChecked, intRow, intCol) = 1
12                End If
13            End If
14        End With


15        Exit Sub
VSFList_Click_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(VSFList_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Sub
Private Sub ClearAllSel()
          Dim intRow As Integer
          Dim intCol As Integer
1         On Error GoTo ClearAllSel_Error

2         With Me.vsfList
3             For intRow = 0 To .Rows - 1
4                 For intCol = 0 To .Cols - 1
5                     If .TextMatrix(intRow, intCol) <> "" Then
6                         .Cell(flexcpChecked, intRow, intCol) = 2
7                     End If
8                 Next
9             Next
10        End With


11        Exit Sub
ClearAllSel_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(ClearAllSel)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear
End Sub
Private Function SelSampe() As String
          Dim intRow As Integer
          Dim intCol As Integer
1         On Error GoTo SelSampe_Error

2         With Me.vsfList
3             For intRow = 0 To .Rows - 1
4                 For intCol = 0 To .Cols - 1
5                     If .Cell(flexcpChecked, intRow, intCol) = 1 Then
6                         SelSampe = .TextMatrix(intRow, intCol)
7                         Exit Function
8                     End If
9                 Next
10            Next
11        End With


12        Exit Function
SelSampe_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(SelSampe)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear
End Function

Private Function Init���븽��() As Boolean
      '���ܣ���ȡ��Ŀ�ĵ������븽��
      '���أ���Ӧ�ĵ��ݶ��������븽��ʱ����True
          Dim strSQL As String, lngIdx As Long
          Dim arrData As Variant, strData As String
          Dim strNoneAppend As String, strHaveAppend As String
          Dim arrSub As Variant, i As Long
          Dim lng�Һ�ID As Long
          Dim rsTmp As ADODB.Recordset


1         On Error GoTo Init���븽��_Error

          'ͨ���Һŵ���ѯ�Һ�ID
2         If mintPatientType = 1 Then
3             strSQL = "Select ID From ���˹Һż�¼ Where no = [1]"
4             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�Һ�ID", CStr(mvar����ID))
5             If Not rsTmp.EOF Then
6                 lng�Һ�ID = Val(rsTmp("ID") & "")
7             End If
8         Else
9             lng�Һ�ID = mvar����ID
10        End If

11        If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
12            rtfAppend.Visible = False
13            vsfAppend.Visible = True

14            mlngMaxWidth = 0
15            With Me.vsfAppend
16                .FixedRows = 0
17                .FixedCols = 0
18                .Rows = 0
19                .Cols = 6
20                .ExtendLastCol = True
21                .GridLines = flexGridNone
22                .AutoSizeMode = flexAutoSizeRowHeight
23                .WordWrap = True

24                .ColKey(0) = "����"
25                .ColKey(1) = "����"
26                .ColKey(2) = "ֻ��": .ColHidden(.ColIndex("ֻ��")) = True
27                .ColKey(3) = "��������": .ColHidden(.ColIndex("��������")) = True
28                .ColKey(4) = "Ҫ��ID": .ColHidden(.ColIndex("Ҫ��ID")) = True
29                .ColKey(5) = "����": .ColHidden(.ColIndex("����")) = True

30                strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����,C.ֻ��,d.������,E.id " & _
                           " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D,������ĿĿ¼ E" & _
                           " Where a.������ĿID = E.id and E.����=[1] And A.Ӧ�ó���=[2]" & _
                           " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
                           " Order by C.����"
31                Set mrsAppend = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, mstrSampleNO, 2)
32                arrData = Split(mstrAppend, "<Split1>")

33                Do While Not mrsAppend.EOF
                      'ȷ����������
34                    strData = ""
                      '��ȡ�°没���е����븽��
35                    If intEMR_Setup = 1 Then
36                        If Not gobjEmrInterface.IsInited Or gobjEmrInterface.IsOffline Then

37                        Else
38                            On Error Resume Next
39                            strData = gobjEmrInterface.GetOrderInspectInfoEX(mintPatientType, mlng����ID, lng�Һ�ID, mrsAppend("������") & "")
40                            If Err.Description <> "" Then
41                                Err.Clear: On Error GoTo Init���븽��_Error
42                                strData = gobjEmrInterface.GetOrderInspectInfo(mlng����ID, mrsAppend("������") & "")
43                            End If
44                        End If
45                    Else
46                        If mstrAppend <> "" Then
                              '�޸�ʱ������ԭ������
47                            For i = 0 To UBound(arrData)
48                                arrSub = Split(arrData(i), "<Split2>")
49                                If arrSub(0) = mrsAppend!��Ŀ Then
50                                    strData = arrSub(3)
51                                    If strData = "" And UBound(arrSub) >= 4 Then
                                          '���Ը��ƻ���ײ�����ҽ�������޸�ʱ�����븽��ҲҪȡȱʡֵ
52                                        If Val(arrSub(4)) = 1 Then
53                                            If Not IsNull(mrsAppend!����) Then
54                                                strData = mrsAppend!����
55                                            ElseIf mlng����ID <> 0 Then
56                                                strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mlng����ID, mvar����ID, _
                                                                               mstrDiagnosis, mintӤ��, mstrAdvItem)
57                                            End If
58                                        End If
59                                    End If

                                      '���ڵĸ���
60                                    strHaveAppend = strHaveAppend & "," & arrSub(0)
61                                    strNoneAppend = Replace(strNoneAppend & ",", "," & arrSub(0) & ",", ",")
                                      '                        If Right(strNoneAppend, 1) = "," Then strNoneAppend = Left(strNoneAppend, Len(strNoneAppend) - 1)
62                                ElseIf InStr(strNoneAppend & ",", "," & arrSub(0) & ",") = 0 _
                                         And InStr(strHaveAppend & ",", "," & arrSub(0) & ",") = 0 Then
63                                    strNoneAppend = strNoneAppend & "," & arrSub(0)    '�ȼǵ�û�еĸ�����
64                                End If
65                            Next
66                        Else
                              '����ʱ��ʹ��Ԥ�������ݻ�Ӳ�����������ȡ
67                            If Not IsNull(mrsAppend!����) Then
68                                strData = mrsAppend!����
69                            ElseIf mlng����ID <> 0 Then
70                                strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mlng����ID, mvar����ID, _
                                                               mstrDiagnosis, mintӤ��, mstrAdvItem)
71                            End If
72                        End If
73                    End If

                      '��������ʾ��RTF��:�����ı����һ��λ�ò���ֱ��¼�뺺��,����ȶ��һ���������Ŀո�
74                    .Rows = .Rows + 1
75                    .TextMatrix(.Rows - 1, .ColIndex("����")) = mrsAppend!��Ŀ & "��"
76                    If mlngMaxWidth < Len(.TextMatrix(.Rows - 1, .ColIndex("����"))) * 220 Then
77                        mlngMaxWidth = Len(.TextMatrix(.Rows - 1, .ColIndex("����"))) * 220
78                    End If
79                    .TextMatrix(.Rows - 1, .ColIndex("����")) = strData
80                    .TextMatrix(.Rows - 1, .ColIndex("ֻ��")) = NVL(mrsAppend!ֻ��, 0)
81                    .TextMatrix(.Rows - 1, .ColIndex("Ҫ��ID")) = NVL(mrsAppend!Ҫ��ID, 0)
82                    .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(mrsAppend!����, 0)
83                    .TextMatrix(.Rows - 1, .ColIndex("��������")) = "�ַ�"

84                    mrsAppend.MoveNext
85                Loop

                  '�ο�����
86                If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then Call GreatCrl

87                .AutoSize 1
88                .ColWidth(.ColIndex("����")) = mlngMaxWidth
89                If .Rows > 0 Then
90                    .Cell(flexcpAlignment, 0, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = flexAlignRightTop
91                    .Cell(flexcpAlignment, 0, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
92                    .Cell(flexcpFontBold, 0, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = True
93                End If

94            End With
95        Else
96            rtfAppend.Visible = True
97            vsfAppend.Visible = False

98            rtfAppend.Text = "": rtfAppend.SelStart = 0: mstrPosition = "": mlngPosition = 0

99            strSQL = "Select C.��Ŀ,C.����,C.Ҫ��ID,C.����,C.ֻ��,d.������,E.id " & _
                       " From ��������Ӧ�� A,�����ļ��б� B,�������ݸ��� C,����������Ŀ D,������ĿĿ¼ E" & _
                       " Where a.������ĿID = E.id and E.����=[1] And A.Ӧ�ó���=[2]" & _
                       " And A.�����ļ�ID=B.ID And B.����=7 And B.ID=C.�ļ�ID And c.Ҫ��id=d.id(+)" & _
                       " Order by C.����"

100           Set mrsAppend = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, mstrSampleNO, 2)
101           If Not mrsAppend.EOF Then
102               arrData = Split(mstrAppend, "<Split1>")
103               With rtfAppend
104                   Do While Not mrsAppend.EOF
                          'ȷ����������
105                       strData = ""
                          '��ȡ�°没���е����븽��
106                       If intEMR_Setup = 1 Then
107                           If Not gobjEmrInterface.IsInited Or gobjEmrInterface.IsOffline Then

108                           Else
109                               On Error Resume Next
110                               strData = gobjEmrInterface.GetOrderInspectInfoEX(mintPatientType, mlng����ID, lng�Һ�ID, mrsAppend("������") & "")
111                               If Err.Description <> "" Then
112                                   Err.Clear: On Error GoTo Init���븽��_Error
113                                   strData = gobjEmrInterface.GetOrderInspectInfo(mlng����ID, mrsAppend("������") & "")
114                               End If
115                           End If
116                       Else
117                           If mstrAppend <> "" Then
                                  '�޸�ʱ������ԭ������
118                               For i = 0 To UBound(arrData)
119                                   arrSub = Split(arrData(i), "<Split2>")
120                                   If arrSub(0) = mrsAppend!��Ŀ Then
121                                       strData = arrSub(3)
122                                       If strData = "" And UBound(arrSub) >= 4 Then
                                              '���Ը��ƻ���ײ�����ҽ�������޸�ʱ�����븽��ҲҪȡȱʡֵ
123                                           If Val(arrSub(4)) = 1 Then
124                                               If Not IsNull(mrsAppend!����) Then
125                                                   strData = mrsAppend!����
126                                               ElseIf mlng����ID <> 0 Then
127                                                   strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mlng����ID, mvar����ID, _
                                                                                   mstrDiagnosis, mintӤ��, mstrAdvItem)
128                                               End If
129                                           End If
130                                       End If

                                          '���ڵĸ���
131                                       strHaveAppend = strHaveAppend & "," & arrSub(0)
132                                       strNoneAppend = Replace(strNoneAppend & ",", "," & arrSub(0) & ",", ",")
133                                       If Right(strNoneAppend, 1) = "," Then strNoneAppend = Left(strNoneAppend, Len(strNoneAppend) - 1)
134                                   ElseIf InStr(strNoneAppend & ",", "," & arrSub(0) & ",") = 0 _
                                             And InStr(strHaveAppend & ",", "," & arrSub(0) & ",") = 0 Then
135                                       strNoneAppend = strNoneAppend & "," & arrSub(0)    '�ȼǵ�û�еĸ�����
136                                   End If
137                               Next
138                           Else
                                  '����ʱ��ʹ��Ԥ�������ݻ�Ӳ�����������ȡ
139                               If Not IsNull(mrsAppend!����) Then
140                                   strData = mrsAppend!����
141                               ElseIf mlng����ID <> 0 Then
142                                   strData = GetAppendItemValue(mrsAppend!��Ŀ, NVL(mrsAppend!Ҫ��ID, 0), mlng����ID, mvar����ID, _
                                                                   mstrDiagnosis, mintӤ��, mstrAdvItem)
143                               End If
144                           End If
145                       End If

                          '��������ʾ��RTF��:�����ı����һ��λ�ò���ֱ��¼�뺺��,����ȶ��һ���������Ŀո�
146                       .SelText = IIf(.Text = "", "", vbCrLf) & mrsAppend!��Ŀ & "�� " & strData
147                       lngIdx = .Find(mrsAppend!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
148                       If lngIdx <> -1 Then
149                           .SelStart = lngIdx
150                           .SelLength = Len(mrsAppend!��Ŀ & "��")
151                           .SelBold = True
152                           .SelIndent = 100
153                           .SelProtected = True
154                       End If
155                       .SelStart = Len(.Text)

156                       mstrPosition = mstrPosition & ";" & mrsAppend!��Ŀ & "��" & "," & NVL(mrsAppend!ֻ��, 0)

157                       mrsAppend.MoveNext
158                   Loop

                      '��궨λ�ڵ�һ�����븽��
159                   mrsAppend.MoveFirst
160                   lngIdx = .Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
161                   If lngIdx <> -1 Then
162                       .SelStart = lngIdx + Len(mrsAppend!��Ŀ & "��") + 1
163                       mlngPosition = InStr(.Text, mrsAppend!��Ŀ & "��")
164                       mstrPosition = Mid(mstrPosition, 2)
165                   End If
166               End With

167               rtfAppend.Visible = True
168               Init���븽�� = True
169           End If

              '�Ѳ����ڵ�������Ŀ��ʾ
170           If strNoneAppend <> "" Then
171               MsgBox "���¸�������Ŀ��Ӧ�ĵ����������Ѳ����ڣ�" & vbCrLf & Mid(strNoneAppend, 2), vbInformation, "100"
172           End If
173       End If

174       Exit Function
Init���븽��_Error:
175       Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(Init���븽��)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
176       Err.Clear
End Function


Private Function GetAppend() As String
      '����           ȡ�ø���
          Dim i As Integer
          Dim strData As String
          Dim lngEnd As Long
          Dim lngBegin As Long
          Dim strAppend As String

          '��鲢�ռ������������
1         On Error GoTo GetAppend_Error
2         If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
3             With Me.vsfAppend
4                 For i = 0 To .Rows - 1
5                     strAppend = strAppend & "<Split1>" & Replace(.TextMatrix(i, .ColIndex("����")), "��", "") & "<Split2>" & .TextMatrix(i, .ColIndex("����")) & "<Split2>" & .TextMatrix(i, .ColIndex("Ҫ��ID")) & "<Split2>" & .TextMatrix(i, .ColIndex("����"))
6                 Next
7             End With

8             GetAppend = Mid(strAppend, Len("<Split1>") + 1)
9         Else
10            mblnNull = False
11            If mrsAppend.RecordCount = 0 Then Exit Function

12            mrsAppend.MoveFirst
13            For i = 1 To mrsAppend.RecordCount
14                strData = ""
15                lngEnd = -1
16                lngBegin = rtfAppend.Find(mrsAppend!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
17                If lngBegin <> -1 Then
18                    lngBegin = lngBegin + Len(mrsAppend!��Ŀ & "��")
19                    If i = mrsAppend.RecordCount Then
20                        lngEnd = Len(rtfAppend.Text)
21                    Else
22                        mrsAppend.MoveNext
23                        lngEnd = rtfAppend.Find(vbCrLf & mrsAppend!��Ŀ & "��", lngBegin, , rtfNoHighlight Or rtfMatchCase)
24                        mrsAppend.MovePrevious
25                    End If
26                End If
27                If lngBegin <> -1 And lngEnd <> -1 Then
                      'MID��������1Ϊ������rtf����0Ϊ����
28                    lngBegin = lngBegin + 1
29                    lngEnd = lngEnd + 1
30                    strData = Mid(rtfAppend.Text, lngBegin, lngEnd - lngBegin)
                      'ȥ��Ϊ��������ı����һ��λ�ò���ֱ��¼�뺺������ӵĿո�
31                    If Left(strData, 1) = " " Then strData = Mid(strData, 2)
32                    If Right(strData, 1) = " " Then strData = Left(strData, Len(strData) - 1)

33                    If Trim(strData) = "" And NVL(mrsAppend!����, 0) = 1 Then
34                        MsgBox "���ݸ���""" & mrsAppend!��Ŀ & """������û����д��", vbInformation, "LIS���뵥"
35                        If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
36                            rtfAppend.SelStart = lngBegin
37                        Else
38                            rtfAppend.SelStart = lngBegin - 1
39                        End If
40                        If rtfAppend.Visible = True Then
41                            mblnNull = True: rtfAppend.SetFocus: Exit Function
42                        End If
43                    ElseIf ActualLen(strData) > 4000 Then
44                        MsgBox "���ݸ���""" & mrsAppend!��Ŀ & """�����ݹ������������2000�����ֻ�4000���ַ���", vbInformation, "LIS���뵥"
45                        If Mid(rtfAppend.Text, lngBegin, 1) = " " Then
46                            rtfAppend.SelStart = lngBegin
47                        Else
48                            rtfAppend.SelStart = lngBegin - 1
49                        End If
50                        If rtfAppend.SelText = " " Then rtfAppend.SelStart = lngBegin
51                        If rtfAppend.Visible = True Then
52                            mblnNull = True: rtfAppend.SetFocus: Exit Function
53                        End If
54                    End If
55                End If

                  'û���������ݵĸ���Ҳ�����˱���
56                strAppend = strAppend & "<Split1>" & mrsAppend!��Ŀ & "<Split2>" & NVL(mrsAppend!����, 0) & "<Split2>" & NVL(mrsAppend!Ҫ��ID) & "<Split2>" & strData

57                mrsAppend.MoveNext
58            Next
59            GetAppend = Mid(strAppend, Len("<Split1>") + 1)
60        End If


61        Exit Function
GetAppend_Error:
62        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GetAppend)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
63        Err.Clear

End Function
Private Function ReasonSelect(ByVal strFind As String, ByVal intType As Integer) As Boolean
      '�������кͿ�����ҩ����ѡ����
      'intType  1-������ҩ���ɣ�2-��������
          Dim blnCancle As Boolean
          Dim strRetrun As String
          Dim lngLeft As Long, lngTop As Long
          
1         On Error GoTo ReasonSelect_Error

2         lngLeft = IIf(intType = 1, 0, cboҽ������.Left) + cboҽ������.Left + Me.Left
3         lngTop = IIf(intType = 1, 0, cboҽ������.Top) + cboҽ������.Top + Me.Top - 2600
4         strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, intType)
5         If Not blnCancle Then
6             If strRetrun = "" Then
7                 If strFind = "" Then
8                     MsgBox "û���ҵ����õ�" & IIf(intType = 1, "������ҩ���ɡ�", "�������С�"), vbInformation, Me.Caption
9                 End If
10            Else
11                If intType = 1 Then
                      
12                ElseIf intType = 2 Then
13                    cboҽ������.Text = strRetrun
14                End If
15            End If
16        End If
17        ReasonSelect = blnCancle


18        Exit Function
ReasonSelect_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(ReasonSelect)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear
End Function
Public Function GetSampleDeptRS(Optional strErr As String) As ADODB.Recordset
          '����       ȡ�òɼ����ҵ����ݼ�
          '����       �ҵ��Ĳɼ��������ݼ�

          Dim strSQL As String
1         On Error GoTo GetSampleDeptRS_Error

2         strSQL = "Select Distinct C.Id, C.����, C.����" & vbNewLine & _
                      "From ������ĿĿ¼ A, ����ִ�п��� B, ���ű� C" & vbNewLine & _
                      "Where A.��� = 'E' And A.�������� = '6' And A.Id = B.������Ŀid And B.ִ�п���id = C.Id and c.����ʱ��=to_date('3000/1/1','yyyy/mm/dd')"
3         Set GetSampleDeptRS = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����")


4         Exit Function
GetSampleDeptRS_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GetSampleDeptRS)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear

End Function

Public Function GetSampleTypeRS(Optional strErr As String) As ADODB.Recordset
      '����       ȡ�òɼ���Ŀ�����ݼ�
      '����       �ҵ��Ĳɼ���Ŀ���ݼ�

          Dim strSQL As String
          Dim strPatientType As String

1         On Error GoTo GetSampleTypeRS_Error

2         Select Case mintPatientType
          Case 1
3             strPatientType = "3,1"
4         Case 2
5             strPatientType = "3,2"
6         Case 3
7             strPatientType = "1"
8         Case 4
9             strPatientType = "4"
10        End Select
          
11        strSQL = "Select /*+ rule */" & vbCrLf & _
                   "    A.ID , A.����, A.����" & vbCrLf & _
                   "   From ������ĿĿ¼ A, �������ÿ��� B" & vbCrLf & _
                   "   Where a.id = b.��ĿID And a.��� = 'E' And a.�������� = '6' And b.����id = [1] And" & vbCrLf & _
                   "         Nvl(a.����ʱ��, To_Date('3000-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) > Sysdate And" & vbCrLf & _
                   "         a.������� In (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist)))" & vbCrLf & _
                   "   Union all" & vbCrLf & _
                   "   Select /*+ rule */" & vbCrLf & _
                   "    A.ID , A.����, A.����" & vbCrLf & _
                   "   From ������ĿĿ¼ A" & vbCrLf & _
                   "   Where a.��� = 'E' And a.�������� = '6' And a.������� In (Select * From Table(Cast(F_Num2list([2]) As Zltools.T_Numlist))) And" & vbCrLf & _
                   "         Nvl(a.����ʱ��, To_Date('3000-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) > Sysdate And Not Exists" & vbCrLf & _
                   "    (Select ��ĿID From �������ÿ��� b Where a.id = b.��ĿID)"
12        Set GetSampleTypeRS = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����", mlngAppForDeptID, strPatientType)
13        If GetSampleTypeRS.RecordCount <= 0 Then
14            strSQL = "select id,����,���� from ������ĿĿ¼ where ��� = 'E' and �������� = '6' "
15            Set GetSampleTypeRS = ComOpenSQL(Sel_His_DB, strSQL, "�ɼ�����")
16        End If

17        Exit Function
GetSampleTypeRS_Error:
18        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GetSampleTypeRS)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
19        Err.Clear


End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-06
'��    ��:  ��̬�����ο�Ҫ�ؿؼ�
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub GreatCrl()
          Dim strSQL As String
          Dim rsRef As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strRefValue As String
          Dim strNum As String
          Dim strUnit As String
          Dim strNum1 As String
          Dim strUnit1 As String
          Dim i As Integer

          '��ȡ��ǰ�����Ŀ����ָ��Ĳο�
1         On Error GoTo GreatCrl_Error

2         strSQL = "Select distinct c.�ο�ID,  c.Ҫ��ID,c.��ʾ����,c.��������, d.Ҫ����" & vbCrLf & _
                 "   From �������ָ�� A, ����ָ��ο���Χ B, ����ο�Ҫ�ض��� C, ����ָ��ο�Ҫ�� D" & vbCrLf & _
                 "   Where a.��Ŀid = b.ָ��id And b.id = c.�ο�id And c.Ҫ��id = d.id And a.���id = [1]"
3         Set mrsReference = ComOpenSQL(Sel_Lis_DB, strSQL, "����ָ��ο�", mlngGroupItemID)

          '��ȡҪ�ص�ѡ��ֵ
          'ƥ���Ա�
4         Select Case mstrSex
          Case "��"
5             mstrSex = 1
6         Case "Ů"
7             mstrSex = 2
8         Case "������"
9             mstrSex = 0
10        Case "δ֪"
11            mstrSex = 9
12        End Select
13        Call GetSeleList(0, mstrSex, "��ֵ", , "�Ա�")

          '����
14        strNum = GetAgeMid(0, mstrAge)
15        strUnit = GetAgeMid(1, mstrAge)
16        strNum1 = GetAgeMid(0, strUnit)
17        strUnit1 = GetAgeMid(1, strUnit)
18        Call GetSeleList(0, CalcAgeUnit(Val(strNum), strUnit) + CalcAgeUnit(Val(strNum1), strUnit1), "��ֵ", , "����")

          '�������
19        Call GetSeleList(0, mstrAppForDept, "�ַ�", , "�������")

          '�걾����
20        Call GetSeleList(0, mstrSample, "�ַ�", , "�걾����")

21        strSQL = "select ID,Ҫ����,�����ֶ���,ֵ������,ֵ����Դ,ֵ�� from ����ָ��ο�Ҫ�� where �����ֶ��� is null"
22        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ָ��ο�Ҫ��")
23        With Me.vsfAppend
24            Do While Not rsTmp.EOF
25                If InStr(",����ʱ��,", "," & rsTmp("Ҫ����") & ",") = 0 Then    '����ʱ���ڳ����������⴦������Ҫҽ��������д
26                    mrsReference.Filter = "Ҫ��ID=" & rsTmp("ID") & " and �������� <>''"
27                    If mrsReference.RecordCount > 0 Then
28                        If .FindRow(rsTmp("Ҫ����") & "��", , .ColIndex("����")) < 0 Then
29                            .Rows = .Rows + 1
30                            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("Ҫ����") & "��"
31                            .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsTmp("ֵ������") & ""
32                            .TextMatrix(.Rows - 1, .ColIndex("Ҫ��ID")) = rsTmp("ID") & ""
33                            If mlngMaxWidth < Len(.TextMatrix(.Rows - 1, .ColIndex("����"))) * 220 Then
34                                mlngMaxWidth = Len(.TextMatrix(.Rows - 1, .ColIndex("����"))) * 220
35                            End If

36                        End If
37                    End If
38                End If

39                rsTmp.MoveNext
40            Loop
41        End With



          '���ֻ��һ��ѡ���Ĭ����ʾ
42        Call ShowDefText

43        Exit Sub
GreatCrl_Error:
44        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GreatCrl)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
45        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-09
'��    ��:  ���ֻ��һ��ѡ���Ĭ����ʾ
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub ShowDefText()
          Dim i As Integer
          Dim J As Integer
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strHave As String
          Dim lngRowFind As Long
          Dim strArr() As String


1         On Error GoTo ShowDefText_Error

2         With Me.vsfAppend
3             For i = 0 To .Rows - 1
4                 If .TextMatrix(i, .ColIndex("Ҫ��ID")) <> "" And .TextMatrix(i, .ColIndex("��������")) = "�ַ�" Then
5                     mrsReference.Filter = "Ҫ��ID=" & Val(.TextMatrix(i, .ColIndex("Ҫ��ID"))) & " and ��������<>''"
6                     Set rsTmp = gobjHisDatabase.CopyNewRec(mrsReference, True)
                      'ȥ���ظ���ѡ��
7                     Do While Not mrsReference.EOF
8                         If InStr("<SP>" & strHave & "<SP>", "<SP>" & mrsReference("Ҫ��ID") & "<S>" & mrsReference("��������") & "<SP>") = 0 Then
9                             rsTmp.AddNew
10                            For J = 0 To rsTmp.Fields.Count - 1
11                                rsTmp.Fields(J).value = mrsReference(rsTmp.Fields(J).Name).value
12                            Next

13                            strHave = strHave & "<SP>" & mrsReference("Ҫ��ID") & "<S>" & mrsReference("��������")
14                        End If
15                        mrsReference.MoveNext
16                    Loop

17                    If rsTmp.RecordCount = 1 Then
18                        .TextMatrix(i, .ColIndex("����")) = rsTmp("��ʾ����") & ""
19                        .Cell(flexcpData, i, .ColIndex("����")) = rsTmp("��������") & ""
20                    End If
21                End If
22            Next

23            If mstrAppend <> "" Then
24                strArr = Split(mstrAppend, "<Split1>")
25                For i = 0 To UBound(strArr)
26                    lngRowFind = .FindRow(Split(strArr(i), "<Split2>")(2), , .ColIndex("Ҫ��ID"))
27                    If lngRowFind > -1 Then
28                        If .TextMatrix(lngRowFind, .ColIndex("����")) = "�ٴ���ϣ�" Then
29                            If .TextMatrix(lngRowFind, .ColIndex("����")) = "" Then
30                                .TextMatrix(lngRowFind, .ColIndex("����")) = Split(strArr(i), "<Split2>")(3)
31                            End If
32                        Else
33                            .TextMatrix(lngRowFind, .ColIndex("����")) = Split(strArr(i), "<Split2>")(3)
34                        End If
35                    End If
36                Next
37            Else

                  '��ȡ�ϴ������ֵ
38                strSQL = "Select a.Ҫ��ID, ��Ŀ, ����" & vbCrLf & _
                         " From ����ҽ������ A," & vbCrLf & _
                         "     (Select ID" & vbCrLf & _
                         "       From (Select ID From ����ҽ����¼ Where ����ID = [1] And ���ID Is Null Order By ����ʱ�� Desc)" & vbCrLf & _
                         "       Where rownum = 1) B" & vbCrLf & _
                         " Where A.ҽ��id = B.ID" & vbCrLf & _
                         " Order By a.����"
39                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ������", mlng����ID)
40                Do While Not rsTmp.EOF
41                    If rsTmp("Ҫ��ID") & "" <> "" Then
42                        lngRowFind = .FindRow(rsTmp("Ҫ��ID") & "", , .ColIndex("Ҫ��ID"))
43                        If lngRowFind > -1 Then
44                            If .TextMatrix(lngRowFind, .ColIndex("����")) = "�ٴ���ϣ�" Then
45                                If .TextMatrix(lngRowFind, .ColIndex("����")) = "" Then
46                                    .TextMatrix(lngRowFind, .ColIndex("����")) = rsTmp("����") & ""
47                                End If
48                            Else
49                                .TextMatrix(lngRowFind, .ColIndex("����")) = rsTmp("����") & ""
50                            End If

51                        End If
52                    Else
53                        lngRowFind = .FindRow(rsTmp("��Ŀ") & "", , .ColIndex("����"))
54                        If lngRowFind > -1 Then
55                            .TextMatrix(lngRowFind, .ColIndex("����")) = rsTmp("����") & ""
56                        End If
57                    End If
58                    rsTmp.MoveNext
59                Loop
60            End If

61        End With

62        Exit Sub
ShowDefText_Error:
63        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(ShowDefText)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
64        Err.Clear
End Sub



'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-09-07
'��    ��:  ͨ������Ҫ�أ���ȡ����Ҫ�ص�ѡ��ֵ
'��    ��:
'           lngRefItemID        ����Ҫ��ID
'           strRefItemVal       ����Ҫ��ֵ
'           strValeType         ֵ������
'           lngReturnID         ��Ҫ��ȡѡ���Ҫ�ص�ID
'           strRefItemName      ����Ҫ�ص����ƣ�����Ҫ��ID�����Ʊ��봫������һ��
'��    ��:
'��    ��:  Ҫ��ѡ���¼��
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Function GetSeleList(ByVal lngRefItemID As Long, ByVal strRefItemVal As String, ByVal strValeType As String, _
                             Optional ByVal lngReturnID As Long, Optional ByVal strRefItemName As String) As ADODB.Recordset
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsRef As ADODB.Recordset
          Dim i As Integer
          Dim blnFind As Boolean
          Dim strRefID As String

1         On Error GoTo GetSeleList_Error

2         If lngRefItemID = 0 And strRefItemName = "" Then
3             MsgBox "Ҫ��ID��Ҫ�������봫��һ��", vbInformation, gSysInfo.AppName
4             Exit Function
5         End If

          '��ʼ����¼��
6         If Not mrsReference Is Nothing Then mrsReference.Filter = ""
7         Set rsTmp = gobjHisDatabase.CopyNewRec(mrsReference, True)
8         Set rsRef = gobjHisDatabase.CopyNewRec(mrsReference)

9         If lngRefItemID > 0 Then
10            mrsReference.Filter = "Ҫ��ID=" & lngRefItemID
11        Else
12            mrsReference.Filter = "Ҫ����='" & strRefItemName & "'"
13        End If
14        If mrsReference.RecordCount > 0 Then mrsReference.MoveFirst
15        Do While Not mrsReference.EOF
16            blnFind = True
17            If strValeType = "��ֵ" Or strValeType = "ʱ���" Then
18                If InStr(mrsReference("��������") & "", ">") > 0 Or InStr(mrsReference("��������") & "", "<") > 0 Or InStr(mrsReference("��������") & "", "=") > 0 Then
19                    If strValeType = "ʱ���" Then
20                        strRefItemVal = Format(strRefItemVal, "hh:mm:ss")
21                        blnFind = CalcNumExpress(Replace(mrsReference("��������") & "", mrsReference("Ҫ����") & "", "cdate(""" & strRefItemVal & """)"))
22                    Else
23                        blnFind = CalcNumExpress(Replace(mrsReference("��������") & "", mrsReference("Ҫ����") & "", strRefItemVal))
24                    End If
25                Else
26                    If Not (strRefItemVal = mrsReference("��������") & "" Or mrsReference("��������") & "" = "") Then
27                        blnFind = False
28                    End If
29                End If
30            Else
31                If Not (strRefItemVal = mrsReference("��������") & "" Or mrsReference("��������") & "" = "") Then
32                    blnFind = False
33                End If
34            End If
35            If blnFind Then
36                If InStr("," & strRefID & ",", "," & mrsReference("�ο�ID") & ",") = 0 Then
37                    rsRef.Filter = "�ο�ID=" & mrsReference("�ο�ID")
38                    Do While Not rsRef.EOF
39                        rsTmp.AddNew
40                        For i = 0 To rsTmp.Fields.Count - 1
41                            rsTmp.Fields(i).value = rsRef(rsTmp.Fields(i).Name).value
42                        Next
43                        rsRef.MoveNext
44                    Loop
45                    strRefID = strRefID & "," & mrsReference("�ο�ID")
46                End If
47            End If
48            mrsReference.MoveNext
49        Loop

50        Set mrsReference = rsTmp
51        Set GetSeleList = gobjHisDatabase.CopyNewRec(rsTmp, True)

52        Exit Function
GetSeleList_Error:
53        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GetSeleList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
54        Err.Clear

End Function


Public Function GetAgeMid(ByVal intType As Integer, ByVal strAge As String) As String
      '����           ת������
      '����           0=ȡ�������� 1=ȡ���䵥λ
          Dim strNO As String
          Dim lngCount As Long

1         On Error GoTo GetAgeMid_Error

2         If intType = 0 Then
3             Do While Len(strAge) > 0
4                 strNO = Mid(strAge, 1, 1)
5                 If IsNumeric(strNO) Then
6                     GetAgeMid = GetAgeMid & strNO
7                 Else
8                     If GetAgeMid <> "" Then Exit Function
9                 End If
10                strAge = Mid(strAge, 2)
11            Loop
12        Else
13            Do While Len(strAge) > 0
14                lngCount = lngCount + 1
15                strNO = Mid(strAge, 1, 1)
16                If Not IsNumeric(strNO) Then
17                    If lngCount > 1 Then
18                        GetAgeMid = strAge
19                        Exit Function
20                    End If
21                End If
22                strAge = Mid(strAge, 2)
23            Loop
24        End If


25        Exit Function
GetAgeMid_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(GetAgeMid)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Function


Public Function CalcNumExpress(strExpress As String, Optional ByRef strErr As String) As String
    '����               ������ʽ
    '����               strExpress = ������ʽ
    '����               ������
    Dim sc
    
    On Error GoTo errH
    
    Set sc = CreateObject("ScriptControl")
    sc.Language = "VBScript"
    CalcNumExpress = sc.Eval(Trim(strExpress))
    
    Exit Function
errH:
    If InStr(Err.Description, "�����") > 0 Or InStr(Err.Description, "���") > 0 Or InStr(strExpress, "0") > 0 Then
        CalcNumExpress = 0
    Else
        strErr = "������(CalcNumExpress)��������Ϣ:" & Err.Number & " " & Err.Description
    End If
    Err.Clear
End Function

Private Function CalcAgeUnit(intAge As Integer, strUnit As String) As Long
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '����                           �Ѵ��������͵�λ����Ϊ��С��λСʱ
          '����
          '   intAge                      ��������
          '   strUnit                     ���䵥λ
          '����
          '                               ������䵥λ����С��λ����
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          
1         On Error GoTo CalcAgeUnit_Error

2         Select Case Mid(strUnit, 1, 1)
              Case "��"
3                 CalcAgeUnit = DateDiff("n", DateAdd("yyyy", intAge * -1, Now()), Now())
4             Case "��"
5                 CalcAgeUnit = DateDiff("n", DateAdd("m", intAge * -1, Now()), Now())
6             Case "��"
7                 CalcAgeUnit = DateDiff("n", DateAdd("y", intAge * -1, Now()), Now())
8             Case "ʱ", "Сʱ"
9                 CalcAgeUnit = DateDiff("n", DateAdd("h", intAge * -1, Now()), Now())
10            Case "��", "����"
11                CalcAgeUnit = intAge
12            Case Else
                  'û����д�������
13                CalcAgeUnit = CLng(intAge) * 365 * 24 * 60
14        End Select


15        Exit Function
CalcAgeUnit_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmAppforBillSelSample", "ִ��(CalcAgeUnit)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Function

