VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSystemParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ĳ�������"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmSystemParaSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8460
      TabIndex        =   0
      Top             =   375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8460
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8490
      TabIndex        =   2
      Top             =   5520
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   -315
      Top             =   5865
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystemParaSet.frx":000C
            Key             =   "Limit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystemParaSet.frx":045E
            Key             =   "bm"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stbPage 
      Height          =   7365
      Left            =   75
      TabIndex        =   5
      Top             =   90
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   12991
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "��������(&0)"
      TabPicture(0)   =   "frmSystemParaSet.frx":09F8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl�����㷨"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl��������ǰ׺"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl����ǰ׺��ʾ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl���۵�λ"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmb�����㷨"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt��������ǰ׺"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "vsfParameter"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optUnit(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optUnit(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "�����������(&1)"
      TabPicture(1)   =   "frmSystemParaSet.frx":0A14
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(23)"
      Tab(1).Control(1)=   "Image1(0)"
      Tab(1).Control(2)=   "vsf����"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "�����(&2)"
      TabPicture(2)   =   "frmSystemParaSet.frx":0A30
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1(1)"
      Tab(2).Control(1)=   "lbl��ʾ"
      Tab(2).Control(2)=   "vsf�ⷿ���"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "����ⷿ����(&3)"
      TabPicture(3)   =   "frmSystemParaSet.frx":0A4C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(1)=   "Image1(3)"
      Tab(3).Control(2)=   "vsf����"
      Tab(3).ControlCount=   3
      Begin VB.OptionButton optUnit 
         Caption         =   "ɢװ��λ"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   17
         Top             =   6705
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "��װ��λ"
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   16
         Top             =   6705
         Width           =   1425
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf���� 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   8055
         _cx             =   14208
         _cy             =   10398
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
      Begin VSFlex8Ctl.VSFlexGrid vsf���� 
         Height          =   6135
         Left            =   -74640
         TabIndex        =   13
         Top             =   1080
         Width           =   7815
         _cx             =   13785
         _cy             =   10821
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
         FormatString    =   $"frmSystemParaSet.frx":0A68
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
      Begin VSFlex8Ctl.VSFlexGrid vsfParameter 
         Height          =   5055
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   8055
         _cx             =   14208
         _cy             =   8916
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
         AllowUserResizing=   1
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
         FormatString    =   $"frmSystemParaSet.frx":0B6C
         ScrollTrack     =   0   'False
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
      Begin VB.TextBox txt��������ǰ׺ 
         Height          =   300
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   10
         Top             =   6240
         Width           =   2145
      End
      Begin VB.ComboBox cmb�����㷨 
         Height          =   300
         ItemData        =   "frmSystemParaSet.frx":1096
         Left            =   1800
         List            =   "frmSystemParaSet.frx":1098
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5797
         Width           =   2145
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf�ⷿ��� 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   8055
         _cx             =   14208
         _cy             =   10821
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
         FormatString    =   $"frmSystemParaSet.frx":109A
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
      Begin VB.Label lbl���۵�λ 
         Caption         =   "����ָ�������۶��۵�λ"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label lbl����ǰ׺��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����¼��2-8λ���ֻ���ĸ��"
         Height          =   180
         Left            =   3960
         TabIndex        =   11
         Top             =   6300
         Width           =   2250
      End
      Begin VB.Label lbl��������ǰ׺ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������ǰ׺"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   6270
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   -74520
         Picture         =   "frmSystemParaSet.frx":1190
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "    ���������ѡ�����ķ��ϲ��ţ����Ĳֿ⣬��������ⷿ���߶�Ӧ�Ĺ�ϵ��"
         Height          =   315
         Left            =   -73920
         TabIndex        =   8
         Top             =   600
         Width           =   7080
      End
      Begin VB.Label lbl�����㷨 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ĳ��������㷨"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5820
         Width           =   1440
      End
      Begin VB.Label lbl��ʾ 
         Caption         =   "    ���������ѡ����ⷿ�Ƿ����漰����鷽ʽ�����ⷿѡ��ʱ˫���򰴡�C�����ɸı�ⷿ�ļ�鷽ʽ��"
         Height          =   435
         Left            =   -74040
         TabIndex        =   4
         Top             =   525
         Width           =   7080
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   -74700
         Picture         =   "frmSystemParaSet.frx":1A5A
         Top             =   450
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   -74670
         Picture         =   "frmSystemParaSet.frx":20DB
         Stretch         =   -1  'True
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ʋ����ڲ�ͬ�ⷿ�����ͨ����"
         Height          =   180
         Index           =   23
         Left            =   -74025
         TabIndex        =   3
         Top             =   720
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmSystemParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnChange As Boolean
Private mblnLoad As Boolean
Private mblnChkClick As Boolean     '�Ƿ��ǳ�������chk�ؼ���ֵ.
Private mintOldChkValue As Integer      '����¿������ľ�ֵ.
Private mrs���� As New ADODB.Recordset
Private mstrPrivs As String
Private Const mlngColor As Long = &H8000000F        '�����޸ĵ��н�������ɫ�ĳɻ�ɫ
Private Const MCON_LNGCOLOR As Long = &H80000005    '���޸ĵ��б�����ɫ
Private mstrOld�ӳ����� As String         '��¼�ɵ� ʱ��������ⰴ��ǰ�ӳ�����
    
Private Enum mPara
    mintʱ�����������ԼӼ������ = 1
    mint������������������ = 2
    mint������¿��ÿ��
    mint�������ⰴ���һ�����ĳɱ��ۼ�����
    mint���İ��ֶμӳ������
    mint���ϸ��������ָ�����ۺ�ָ���ۼ�
    mintʱ��������ⰴ��ǰ�ӳ�����
    mint�������ϲ���������������
    mintʱ������ֱ��ȷ���ۼ�
    mint�⹺��ⵥ��Ҫ�˲�
    mintʱ���������ȡ�ϴ��ۼ�
    mintCount = 12
End Enum

Private Enum m����
    mint���ڿⷿ = 1
    mint�Է��ⷿ = 2
    mint����
    mint���ڿⷿid
    mint�Է��ⷿid
    mintCount = 6
End Enum

Private Enum m�ⷿ���
    mintid = 0
    mint����
    mint����
    mint��鷽ʽ
    mintCount = 4
End Enum

Private Enum m�ⷿ����
    mint����id = 0
    mint���ϲ��� = 1
    mint�ⷿid = 2
    mint���Ĳֿ� = 3
    mint����ⷿid
    mint����ⷿ
    mint����
    mintCount = 7
End Enum

'Private Sub bill����_AfterAddRow(Row As Long)
'    bill����.TextMatrix(bill����.Rows - 1, 6) = "��"
'End Sub

'Private Sub bill����_cboClick(ListIndex As Long)
'    Dim intRow As Integer
'    Dim lng����id As Long
'
'    With vsf����
'        If ListIndex < 0 Then Exit Sub
'        If .Col = 1 Then
'            .TextMatrix(.Row, 0) = .ItemData(ListIndex)
'        ElseIf .Col = 3 Then
'            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
'        ElseIf .Col = 5 Then
'            .TextMatrix(.Row, 4) = .ItemData(ListIndex)
'        End If
'
'        lng����id = Val(.TextMatrix(.Row, 0))
'
'        For intRow = 1 To .Rows - 1
'            If Val(.TextMatrix(intRow, 0)) > 0 Then
'                If Val(.TextMatrix(intRow, 0)) = lng����id And intRow <> .Row Then
'                    .TextMatrix(.Row, 0) = ""
'                    .TextMatrix(.Row, 1) = ""
'                    .TextMatrix(.Row, 2) = ""
'                    .TextMatrix(.Row, 3) = ""
'                    .TextMatrix(.Row, 4) = ""
'                    .TextMatrix(.Row, 5) = ""
'                    .TextMatrix(.Row, 6) = ""
'                    Exit For
'                End If
'            End If
'        Next
'    End With
'
'    mblnChange = True
'End Sub

Private Sub bill����_cboKeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim lng����id As Long
    
    With vsf����
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .Col = 1 Then
                .TextMatrix(.Row, 0) = .ItemData(.ListIndex)
            ElseIf .Col = 3 Then
                .TextMatrix(.Row, 1) = .ItemData(.ListIndex)
            ElseIf .Col = 5 Then
                .TextMatrix(.Row, 3) = .ItemData(.ListIndex)
            End If
            
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, 0)) > 0 Then
                    If Val(.TextMatrix(intRow, 0)) = lng����id And intRow <> .Row Then
                        .TextMatrix(.Row, 0) = ""
                        .TextMatrix(.Row, 1) = ""
                        .TextMatrix(.Row, 2) = ""
                        .TextMatrix(.Row, 3) = ""
                        .TextMatrix(.Row, 4) = ""
                        .TextMatrix(.Row, 5) = ""
                        .TextMatrix(.Row, 6) = ""
                        Exit For
                    End If
                End If
            Next
        End If
        mblnChange = True
    End With
End Sub

Private Sub vsf����_ChangeEdit()
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
    
    If InStr(1, vsf����.EditText, "-") <= 0 Then Exit Sub
    strID = Mid(vsf����.EditText, 1, InStr(1, vsf����.EditText, "-") - 1)
    str���� = Mid(vsf����.EditText, InStr(1, vsf����.EditText, "-") + 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ų�ѯ", strID, str����)
    If rsTemp.RecordCount > 0 Then
        With vsf����
            If .Col = m�ⷿ����.mint���ϲ��� Then
                .TextMatrix(.Row, m�ⷿ����.mint����id) = rsTemp!Id
            ElseIf .Col = m�ⷿ����.mint���Ĳֿ� Then
                .TextMatrix(.Row, m�ⷿ����.mint�ⷿid) = rsTemp!Id
            ElseIf .Col = m�ⷿ����.mint����ⷿ Then
                .TextMatrix(.Row, m�ⷿ����.mint����ⷿid) = rsTemp!Id
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf����_DblClick()
    With vsf����
        If .Col = m�ⷿ����.mint���� Then
            If .TextMatrix(.Row, m�ⷿ����.mint����) = "" Then
                .TextMatrix(.Row, m�ⷿ����.mint����) = "��"
            Else
                .TextMatrix(.Row, m�ⷿ����.mint����) = ""
            End If
        End If
    End With
End Sub



Private Sub bill����_cboClick(ListIndex As Long)
   
    With vsf����
        If ListIndex < 0 Then Exit Sub
        If .Col = 0 Then
            .RowData(.Row) = .ItemData(ListIndex)
        ElseIf .Col = 1 Then
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
        End If
'        .TextMatrix(.Row, .Col) = .CboText
        
        If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
    End With
    mblnChange = True
End Sub

Private Sub bill����_cboKeyDown(KeyCode As Integer, Shift As Integer)
  With vsf����
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                .RowData(.Row) = .ItemData(.ListIndex)
            End If
            If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
        End If
        mblnChange = True
    End With
End Sub

'Private Sub bill����_DblClick(Cancel As Boolean)
'    '�������һ�еı仯
'    With vsf����
'        If .MouseRow = 0 Then Exit Sub
'        If .MouseCol <> .Cols - 1 Then Exit Sub
'        Select Case Left(.TextMatrix(.Row, .Col), 1)
'            Case "1"
'                .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
'            Case "2"
'                .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
'            Case Else
'                .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
'        End Select
'        mblnChange = True
'End With
'End Sub

Private Sub bill����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)

    With vsf����
            If .Col = 2 Then
                '����ֵ��ֻ����س���
                If KeyCode <> vbKeyReturn Then Exit Sub
                If .TextMatrix(.Row, 2) = "" Then
                    '����һ���ؼ�
                    OS.PressKey vbKeyTab
                End If
            ElseIf .Col >= 3 Then
                If KeyCode <> vbKeyReturn And KeyCode <> vbKeyDelete Then KeyCode = 0: Cancel = True
            End If
    End With
End Sub

Private Sub bill����_KeyPress(KeyAscii As Integer)
    With vsf����
            If .Col = 3 Then
                Select Case KeyAscii
                    Case Asc(" ")
                        '�л������־
                        Select Case Left(.TextMatrix(.Row, .Col), 1)
                            Case "1"
                                .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                            Case "2"
                                .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                            Case Else
                                .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                        End Select
                        mblnChange = True
                    Case vbKey1
                        .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                        mblnChange = True
                    Case vbKey2
                        .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                        mblnChange = True
                    Case vbKey3
                        .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                        mblnChange = True
                End Select
                mblnChange = True
            End If
    End With
End Sub

Private Function Check�ƿⵥ() As Boolean
    '����:����ƿⵥ�Ƿ����δ��˵ĵ���
    Dim rsTemp As New ADODB.Recordset
    Dim blnTemp As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From ҩƷ�շ���¼ where ����=19 and ������� is null and rownum<=3 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then
        blnTemp = True
    Else
        blnTemp = rsTemp.RecordCount = 0
    End If
    If blnTemp = False Then
        ShowMsgBox "���������쵥���ƿⵥ��δ�󵥾�,���ȴ����������!"
    End If
    Check�ƿⵥ = blnTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check��ⵥ() As Boolean
    '����:����ƿⵥ�Ƿ����δ��˵ĵ���
    Dim rsTemp As New ADODB.Recordset
    Dim blnTemp As Boolean
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID From ҩƷ�շ���¼ where ����=15 and ������� is null and rownum<=3 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    
    If rsTemp.EOF Then
        blnTemp = True
    Else
        blnTemp = rsTemp.RecordCount = 0
    End If
    If blnTemp = False Then
        ShowMsgBox "������δ��˵��⹺��ⵥ,�봦���������!"
    End If
    Check��ⵥ = blnTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub chkcheck_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
'End Sub
'
'Private Sub chk�����¿��_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))

End Sub

Private Sub cmdOk_Click()
    If ISValid() = False Then Exit Sub
    If Save����() = False Then Exit Sub
    Call InitSystemPara
    mblnChange = False
    Unload Me
End Sub
Private Function ISValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    Dim lngIndex As Long
    Dim lngCol1 As Long, lngCol2 As Long
    Dim i As Integer
    Dim j As Integer
    
    With vsf����
        
        For lngRow = 1 To .Rows - 1
            If (.TextMatrix(lngRow, 1) = "" Or .TextMatrix(lngRow, 2) = "" Or .TextMatrix(lngRow, 3) = "") And lngRow <> .Rows - 1 Then
                MsgBox "��" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 1
                stbPage.Tab = 1
                Exit Function
            End If
'            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
'                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
'                .Row = lngRow
'                .Col = 0
'                stbPage.Tab = 1
'                Exit Function
'            End If
            If .TextMatrix(lngRow, 1) = .TextMatrix(lngRow, 2) And lngRow <> .Rows - 1 Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 1
                stbPage.Tab = 1
                Exit Function
            End If
            For j = 1 To .Rows - 1
                If .TextMatrix(i, 1) = .TextMatrix(j, 1) And .TextMatrix(i, 2) = .TextMatrix(j, 2) And i <> j Then
                    MsgBox "��" & i & "�����" & j & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = i
                    .Col = 1
                    stbPage.Tab = 1
                    Exit Function
                End If
            Next
            
'            For lngTemp = lngRow + 1 To .Rows - 1
'                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
'                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
'                    .Row = lngTemp
'                    .Col = 1
'                    stbPage.Tab = 1
'                    Exit Function
'                End If
'            Next
        Next
    End With
    
    With vsfParameter
        If mintOldChkValue <> Val(.TextMatrix(mPara.mint������¿��ÿ��, 1)) Then
            If Check�ƿⵥ = False Then Exit Function
        End If
        If IIf(.TextMatrix(mPara.mintʱ��������ⰴ��ǰ�ӳ�����, 1) = "", 0, .TextMatrix(mPara.mintʱ��������ⰴ��ǰ�ӳ�����, 1)) <> mstrOld�ӳ����� Then '��¼ԭ������ⰴ��ǰ�ӳ����ۺ����ڵ�ֵ�Ƿ�һ��
            '��Ҫ��֤�Ƿ�����˵�
            If Check��ⵥ = False Then Exit Function
        End If
    End With
    ISValid = True
End Function

Private Function Save����() As Boolean
    Dim str����  As String
    Dim i As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim str����id As String
    Dim arr�ⷿ As Variant
    Dim str���ڿⷿid As String
    Dim str�Է��ⷿid As String
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim bln���� As Boolean
    Dim arrSQL  As Variant

    On Error GoTo ErrHandle

    gcnOracle.BeginTrans
    
    With vsfParameter
        Call zlDatabase.SetPara(82, IIf(.TextMatrix(mPara.mintʱ�����������ԼӼ������, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(83, IIf(.TextMatrix(mPara.mint������������������, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(95, IIf(.TextMatrix(mPara.mint������¿��ÿ��, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(120, IIf(.TextMatrix(mPara.mint�������ⰴ���һ�����ĳɱ��ۼ�����, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(121, IIf(.TextMatrix(mPara.mint���İ��ֶμӳ������, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(123, IIf(.TextMatrix(mPara.mint���ϸ��������ָ�����ۺ�ָ���ۼ�, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(127, IIf(.TextMatrix(mPara.mintʱ��������ⰴ��ǰ�ӳ�����, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(132, IIf(.TextMatrix(mPara.mint�������ϲ���������������, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(136, IIf(.TextMatrix(mPara.mintʱ������ֱ��ȷ���ۼ�, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(140, IIf(.TextMatrix(mPara.mint�⹺��ⵥ��Ҫ�˲�, 1) = "1", 1, 0), glngSys, 0)
        Call zlDatabase.SetPara(229, IIf(.TextMatrix(mPara.mintʱ���������ȡ�ϴ��ۼ�, 1) = "1", 1, 0), glngSys, 0)
    End With
    
    Call zlDatabase.SetPara(88, IIf(optUnit(0).Value = True, 0, 1), glngSys, 0)
    Call zlDatabase.SetPara(156, IIf(cmb�����㷨.ListIndex = -1, 0, cmb�����㷨.ListIndex), glngSys, 0)

    If zlStr.IsHavePrivs(mstrPrivs, "��������ǰ׺") = True Then
        Call zlDatabase.SetPara(159, IIf(Trim(txt��������ǰ׺.Text) = "", "", UCase(Trim(txt��������ǰ׺.Text))), glngSys, 0)
    End If

    strTemp = ""
    arrSQL = Array()
    With vsf����
        For lngRow = 1 To .Rows - 1
            str���� = Left(.TextMatrix(lngRow, m����.mint����), 1)
            If str���� = "" Then str���� = "3"
            
            str���ڿⷿid = ""
            str�Է��ⷿid = ""
            
            If .TextMatrix(lngRow, m����.mint���ڿⷿid) = "" And lngRow <> .Rows - 1 Then
                gstrSQL = "select id from ���ű� where ����=[1]"
                strID = Mid(.TextMatrix(lngRow, m����.mint���ڿⷿ), 1, InStr(1, .TextMatrix(lngRow, m����.mint���ڿⷿ), "-") - 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID)
                If rsTemp.RecordCount > 0 Then
                    str���ڿⷿid = rsTemp!Id
                End If
            Else
                str���ڿⷿid = .TextMatrix(lngRow, m����.mint���ڿⷿid)
            End If
            
            If .TextMatrix(lngRow, m����.mint�Է��ⷿid) = "" And lngRow <> .Rows - 1 Then
                strID = Mid(.TextMatrix(lngRow, m����.mint�Է��ⷿ), 1, InStr(1, .TextMatrix(lngRow, m����.mint�Է��ⷿ), "-") - 1)
                gstrSQL = "select id from ���ű� where ����=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Է��ⷿ��ѯ", strID)
                If rsTemp.RecordCount > 0 Then
                    str�Է��ⷿid = rsTemp!Id
                End If
            Else
                str�Է��ⷿid = .TextMatrix(lngRow, m����.mint�Է��ⷿid)
            End If
            If str���ڿⷿid <> "" Or str�Է��ⷿid <> "" Then
                If LenB(StrConv(strTemp & str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ",", vbFromUnicode)) >= 4000 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strTemp
                    strTemp = str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ","
                    bln���� = True
                Else
                    strTemp = strTemp & str���ڿⷿid & "," & str�Է��ⷿid & "," & str���� & ","
                End If
            End If
        Next
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strTemp
    End With
    
    For i = 0 To UBound(arrSQL)
        If bln���� = True Then
            If i = 0 Then
                Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "ɾ�����ۼ�¼")
            Else
                Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',1" & ")", "ɾ�����ۼ�¼")
            End If
        Else
            Call zlDatabase.ExecuteProcedure("zl_�����������_Modify('" & CStr(arrSQL(i)) & "',0" & ")", "ɾ�����ۼ�¼")
        End If
    Next

    '����ⷿ���
    gstrSQL = ""
    With vsf�ⷿ���
        For i = 1 To .Rows - 1
            gstrSQL = gstrSQL & .TextMatrix(i, m�ⷿ���.mintid) & "," & Switch(.TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "0-�����", "0", .TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "1-��飬��������", "1", .TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = "2-��飬�����ֹ", "2") & ","
        Next
    End With

    gstrSQL = "Zl_���ϳ�����_insert('" & gstrSQL & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption


    '��������ⷿ����
    If zlStr.IsHavePrivs(mstrPrivs, "��������ⷿ����") = True Then
        strTemp = ""
        With vsf����
            For i = 1 To .Rows - 1
                If .TextMatrix(i, m�ⷿ����.mint����) = "��" And Val(.TextMatrix(i, m�ⷿ����.mint����id)) > 0 And Val(.TextMatrix(i, m�ⷿ����.mint�ⷿid)) > 0 And Val(.TextMatrix(i, m�ⷿ����.mint����ⷿid)) > 0 Then
                    If InStr(1, "," & str����id & ",", "," & Val(.TextMatrix(i, m�ⷿ����.mint����id)) & ",") = 0 Then
                        str����id = IIf(str����id = "", "", str����id & ",") & .TextMatrix(i, m�ⷿ����.mint����id)
                        strTemp = IIf(strTemp = "", "", strTemp & "|") & .TextMatrix(i, m�ⷿ����.mint����id) & "," & .TextMatrix(i, m�ⷿ����.mint�ⷿid) & "," & .TextMatrix(i, m�ⷿ����.mint����ⷿid)
                    End If
                End If
            Next
        End With

        gstrSQL = "Zl_����ⷿ����_Update('" & strTemp & "')"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If

    '������ϣ������ύ
    gcnOracle.CommitTrans
    Save���� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    mblnLoad = True
    mblnChkClick = False
    mstrPrivs = gstrPrivs

    '���г�ʼ��
    Call InitCtrl
    Call Load�ⷿ���
    Call Load��������
    Call Load����ⷿ����
    Call initVsfPara
    
'    Me.lvwCheckMed.Sorted = True
    
    '������������
    Call LoadPara
    Call SetColor
    
    If zlStr.IsHavePrivs(mstrPrivs, "��������ⷿ����") = False Then
        stbPage.TabVisible(3) = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "��������ǰ׺") = False Then
        lbl��������ǰ׺.Visible = False
        txt��������ǰ׺.Visible = False
        lbl����ǰ׺��ʾ.Visible = False
        lbl���۵�λ.Top = lbl��������ǰ׺.Top
        optUnit(0).Top = lbl���۵�λ.Top
        optUnit(1).Top = optUnit(0).Top
    End If
    
    '�ָ��п�
    RestoreFlexState vsf����, App.ProductName & "\" & Me.Name & "\��������"
    
    '��ʼ���ɹ�
    mblnChange = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPara()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnChkClick = True
'    chkCheck(5).Tag = ""
    Set rsTemp = ReturnParaData(glngSys, "82,83,88,95,120,121,123,127,132,136,140,156,159,229")
    With rsTemp
        Do While Not .EOF
            Select Case zlStr.Nvl(!������, 0)
                Case 88
                    optUnit(0).Value = IIf(zlStr.Nvl(!����ֵ, 0) = 0, True, False)
                    optUnit(1).Value = IIf(zlStr.Nvl(!����ֵ, 0) = 0, False, True)
                Case 156
                    '���ĳ����㷨
                    If zlStr.Nvl(!����ֵ, 0) = 1 Then
                        cmb�����㷨.ListIndex = 1
                    Else
                        cmb�����㷨.ListIndex = 0
                    End If
                Case 159
                    '��������ǰ׺
                    txt��������ǰ׺.Text = IIf(IsNull(!����ֵ), "", !����ֵ)
            End Select
            With vsfParameter
                For i = 1 To .Rows - 1
                    If zlStr.Nvl(rsTemp!������, 0) = .TextMatrix(i, 0) And rsTemp!����ֵ = 1 Then
                        .TextMatrix(i, 1) = 1
                    End If
                    If zlStr.Nvl(rsTemp!������, 0) = "95" Then
                        mintOldChkValue = Val(rsTemp!����ֵ)
                    End If
                    If zlStr.Nvl(rsTemp!������, 0) = "127" Then
                        mstrOld�ӳ����� = Val(rsTemp!����ֵ)
                    End If
                Next
            End With
            
            .MoveNext
        Loop
    End With
    mblnChkClick = False
End Function
Private Sub InitCtrl()
    Dim lngIndex As Long
    
    With vsf����
        .Cols = m����.mintCount  '����һ��������
        .ColWidth(0) = 0
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    With vsf����
        .Cols = m�ⷿ����.mintCount
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, m�ⷿ����.mint����id) = "����id"
        .TextMatrix(0, m�ⷿ����.mint���ϲ���) = "���ϲ���"
        .TextMatrix(0, m�ⷿ����.mint�ⷿid) = "���Ĳֿ�ID"
        .TextMatrix(0, m�ⷿ����.mint���Ĳֿ�) = "���Ĳֿ�"
        .TextMatrix(0, m�ⷿ����.mint����ⷿid) = "����ⷿID"
        .TextMatrix(0, m�ⷿ����.mint����ⷿ) = "����ⷿ"
        .TextMatrix(0, m�ⷿ����.mint����) = "����"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 0
        .ColWidth(5) = 2000
        .Editable = flexEDKbdMouse
    End With
    
    With cmb�����㷨
        .Clear
        .AddItem "0-�������Ƚ��ȳ�"
        .AddItem "1-��Ч������ȳ�"
    End With
End Sub

Private Sub Load��������()
    '����:װ�������������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTemp As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With vsf����
        '����װ��ⷿ
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('���Ŀ�','�Ƽ���','����ⷿ','���ϲ���') " & _
                   "   and  b.����ID=a.ID and " & Where����ʱ��("A") & _
                   " order by ����"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        .Rows = 1
        .Rows = 2
        strTemp = ""
        If Not rsTemp.EOF Then
            rsTemp.MoveFirst
            For i = 1 To rsTemp.RecordCount
                strTemp = strTemp & rsTemp!���� & "-" & rsTemp!���� & "|"
                rsTemp.MoveNext
            Next
        End If
        .ColComboList(m����.mint���ڿⷿ) = strTemp
        .ColComboList(m����.mint�Է��ⷿ) = strTemp
        
        .ColComboList(m����.mint����) = "1-���ڿⷿ������Է��ⷿ|2-�Է��ⷿ���������ڿⷿ|3-���ⷿ���˫����ͨ"
'        Do Until rsTemp.EOF
'            .AddItem rsTemp("����") & "-" & rsTemp("����")
'            .ItemData(.NewIndex) = rsTemp("ID")
'            rsTemp.MoveNext
'        Loop
        
        'װ�������������
        gstrSQL = "select A.���ڿⷿID,A.�Է��ⷿID,A.����" & _
                ",B.���� as ���ڱ���,B.���� as ��������,C.���� as �Է�����,C.���� as �Է����� " & _
                " from ����������� A,���ű� B,���ű� C " & _
                " where A.���ڿⷿID= B.ID and A.�Է��ⷿID=C.ID " & _
                "   and (b.����ʱ��=to_date('3000-1-1','yyyy-mm-dd') or b.����ʱ�� is null) " & _
                " order by b.����, c.���� "
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            
'            .RowData(lngRow) = rsTemp("���ڿⷿID")
'            .TextMatrix(lngRow, 0) = rsTemp("���ڱ���") & "-" & rsTemp("��������")
'            .TextMatrix(lngRow, 1) = rsTemp("�Է�����") & "-" & rsTemp("�Է�����")
'            .TextMatrix(lngRow, 2) = rsTemp("�Է��ⷿID")
'            .TextMatrix(lngRow, 3) = Switch(rsTemp("����") = 1, "1-���ڿⷿ������Է��ⷿ", _
'                                            rsTemp("����") = 2, "2-�Է��ⷿ���������ڿⷿ", _
'                                                          True, "3-���ⷿ���˫����ͨ")
                                                          
            .TextMatrix(lngRow, m����.mint���ڿⷿ) = IIf(IsNull(rsTemp!���ڿⷿid), "", rsTemp!���ڱ��� & "-" & rsTemp!��������)
            .TextMatrix(lngRow, m����.mint���ڿⷿid) = rsTemp!���ڿⷿid
            .TextMatrix(lngRow, m����.mint�Է��ⷿ) = IIf(IsNull(rsTemp!�Է��ⷿID), "", rsTemp!�Է����� & "-" & rsTemp!�Է�����)
            .TextMatrix(lngRow, m����.mint�Է��ⷿid) = rsTemp!�Է��ⷿID
            .TextMatrix(lngRow, m����.mint����) = Switch(rsTemp("����") = 1, "1-���ڿⷿ������Է��ⷿ", _
                                            rsTemp("����") = 2, "2-�Է��ⷿ���������ڿⷿ", _
                                                          True, "3-���ⷿ���˫����ͨ")
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        For i = 0 To vsf����.Rows - 1
            .RowHeight(i) = 300
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load����ⷿ����()
    '����:װ����������ⷿ���չ�ϵ
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With vsf����
        'ȡ���з��ϲ��ţ����Ŀ⣬����ⷿ
        mrs����.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.����, b.�������� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('���Ŀ�','���ϲ���','����ⷿ') " & _
                   " and  b.����ID=a.ID and " & Where����ʱ��("A") & " order by ����"
        
        zlDatabase.OpenRecordset mrs����, gstrSQL, Me.Caption
        
        'װ��Ŀǰ������ⷿ���չ�ϵ
        gstrSQL = "Select b.Id As ����id, b.���� || '-' || b.���� As ���ϲ���, c.Id As �ⷿid, c.���� || '-' || c.���� As ���Ĳֿ�," & _
                  " d.Id As ����ⷿid,d.���� || '-' || d.���� As ����ⷿ " & _
                  "From ����ⷿ���� A, ���ű� B, ���ű� C, ���ű� D " & _
                  "Where a.����id = b.Id And a.�ⷿid = c.Id And a.����ⷿid = d.Id " & _
                  "  And (b.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd') or b.����ʱ�� is null) " & _
                  "Order by b.���� "
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        lngRow = 1
        
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .TextMatrix(lngRow, m�ⷿ����.mint����id) = rsTemp!����id
            .TextMatrix(lngRow, m�ⷿ����.mint���ϲ���) = rsTemp!���ϲ���
            .TextMatrix(lngRow, m�ⷿ����.mint�ⷿid) = rsTemp!�ⷿID
            .TextMatrix(lngRow, m�ⷿ����.mint���Ĳֿ�) = rsTemp!���Ĳֿ�
            .TextMatrix(lngRow, m�ⷿ����.mint����ⷿid) = rsTemp!����ⷿid
            .TextMatrix(lngRow, m�ⷿ����.mint����ⷿ) = rsTemp!����ⷿ
            .TextMatrix(lngRow, m�ⷿ����.mint����) = "��"
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs���� = Nothing
    
    SaveFlexState vsf����, App.ProductName & "\" & Me.Name & "\��������"
    If mblnChange = False Then Exit Sub
    
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

'Private Sub lvwCheckMed_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If lvwCheckMed.SortKey = ColumnHeader.Index - 1 Then
'        lvwCheckMed.SortOrder = IIf(lvwCheckMed.SortOrder = lvwAscending, lvwDescending, lvwAscending)
'    Else
'        lvwCheckMed.SortKey = ColumnHeader.Index - 1
'        lvwCheckMed.SortOrder = lvwAscending
'    End If
'End Sub

Private Sub stbPage_Click(PreviousTab As Integer)
    Select Case stbPage.Tab
        Case 0
            vsfParameter.SetFocus
        Case 1
            vsf����.SetFocus
        Case 2
            vsf�ⷿ���.SetFocus
        Case Else
    End Select
End Sub


Private Sub Load�ⷿ���()
    '���ܣ���ʼ���ⷿ
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim objItem As ListItem
    On Error GoTo ErrHandle
    
    gstrSQL = _
        "SELECT B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0) ��鷽ʽ" & vbCrLf & _
        " FROM ��������˵�� A, ���ű� B, ���ϳ����� C" & vbCrLf & _
        " WHERE A.����ID = B.ID AND A.����ID = C.�ⷿID(+) AND" & vbCrLf & _
        "      A.�������� IN" & vbCrLf & _
        "      ('���Ŀ�','�Ƽ���','���ϲ���','����ⷿ') " & vbCrLf & _
        "     And (b.����ʱ��=to_date('3000-1-1', 'yyyy-mm-dd') or b.����ʱ�� is null) " & vbCrLf & _
        " GROUP BY B.ID,B.����, B.����, NVL(C.��鷽ʽ, 0)" & vbCrLf & _
        " ORDER BY B.���� "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    Me.vsf�ⷿ���.Rows = 1
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        vsf�ⷿ���.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
'            Set objItem = Me.lvwCheckMed.ListItems.Add(, "C_" & rsTmp!Id, "[" & zlStr.Nvl(rsTmp!����) & "]", "bm", "bm")
'            objItem.SubItems(1) = zlStr.Nvl(rsTmp!����)
'            objItem.SubItems(2) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
'            objItem.Tag = rsTmp!Id
            With vsf�ⷿ���
                .TextMatrix(i, m�ⷿ���.mintid) = rsTmp!Id
                .TextMatrix(i, m�ⷿ���.mint����) = rsTmp!����
                .TextMatrix(i, m�ⷿ���.mint����) = rsTmp!����
                .TextMatrix(i, m�ⷿ���.mint��鷽ʽ) = Switch(rsTmp!��鷽ʽ = 0, "0-�����", rsTmp!��鷽ʽ = 1, "1-��飬��������", rsTmp!��鷽ʽ = 2, "2-��飬�����ֹ")
            End With
            rsTmp.MoveNext
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'Private Sub lvwCheckMed_DblClick()
'    If Not Me.lvwCheckMed.SelectedItem Is Nothing Then
'        lvwCheckMed.SelectedItem.SubItems(2) = Switch(lvwCheckMed.SelectedItem.SubItems(2) = "0-�����", "1-��飬��������", lvwCheckMed.SelectedItem.SubItems(2) = "1-��飬��������", "2-��飬�����ֹ", lvwCheckMed.SelectedItem.SubItems(2) = "2-��飬�����ֹ", "0-�����")
'    End If
'End Sub

'Private Sub lvwCheckMed_KeyPress(KeyAscii As Integer)
'    If UCase(Chr(KeyAscii)) = "C" Then
'        Call lvwCheckMed_DblClick
'    End If
'End Sub

Private Sub optUnit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If optUnit(1).Enabled Then optUnit(1).SetFocus
    Else
        stbPage.Tab = 1
    End If
End Sub

Private Sub txt��������ǰ׺_Change()
    txt��������ǰ׺.Text = UCase(txt��������ǰ׺.Text)
End Sub

Private Sub txt��������ǰ׺_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Exit Sub
        End If
    End Select
    KeyAscii = 0
End Sub


Private Sub txt��������ǰ׺_LostFocus()
    If Len(txt��������ǰ׺) > 8 Then
        txt��������ǰ׺.Text = Mid(txt��������ǰ׺.Text, 1, 8)
    End If
End Sub

Private Sub initVsfPara()
    Dim i As Integer
    '��ʼ���������vsflexgrid�ؼ�
    With vsfParameter
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        
        .GridLines = flexGridInset
        .GridColor = &H0&
        .AllowUserResizing = flexResizeColumns
        .Rows = mPara.mintCount
        .Cols = 4
        .ExtendLastCol = True '���һ�������
        .WordWrap = True
        .AutoSize 3, 3, False, 0 = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ColDataType(1) = flexDTBoolean
        .ScrollBars = flexScrollBarVertical '�����������ȡ����
        .ColHidden(0) = True
    End With
    
    With vsf�ⷿ���
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        .ExtendLastCol = True
        .Cols = 4
        For i = 0 To .Rows - 1
            .RowHeight(i) = 300
        Next
        .ColComboList(m�ⷿ���.mint��鷽ʽ) = "0-�����|1-��飬��������|2-��飬�����ֹ"
        .ColHidden(0) = True
    End With
    
    With vsf����
        .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        For i = 0 To .Rows - 1
            .RowHeight(i) = 300
        Next
    End With
End Sub

Private Sub vsfParameter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With vsfParameter
            If .CellBackColor = mlngColor Then
                Exit Sub
            End If
            Select Case .Row
                Case mPara.mint������¿��ÿ��
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mint������������������, 0, mPara.mint������������������, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mint������������������, 1) = "1"
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mint������������������, 0, mPara.mint������������������, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mintʱ���������ȡ�ϴ��ۼ�
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mintʱ�����������ԼӼ������, 1) = ""
                        .TextMatrix(mPara.mint���İ��ֶμӳ������, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mintʱ�����������ԼӼ������
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mintʱ���������ȡ�ϴ��ۼ�, 1) = ""
                        .TextMatrix(mPara.mint���İ��ֶμӳ������, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case mPara.mint���İ��ֶμӳ������
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                        .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = mlngColor
                        .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = mlngColor
                        .TextMatrix(mPara.mintʱ���������ȡ�ϴ��ۼ�, 1) = ""
                        .TextMatrix(mPara.mintʱ�����������ԼӼ������, 1) = ""
                    Else
                        .TextMatrix(.Row, 1) = ""
                        .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = MCON_LNGCOLOR
                        .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = MCON_LNGCOLOR
                    End If
                Case Else
                    If .TextMatrix(.Row, 1) = "" Then
                        .TextMatrix(.Row, 1) = "1"
                    Else
                        .TextMatrix(.Row, 1) = ""
                    End If
            End Select
        End With
    End If
End Sub

Private Sub SetColor()
    '�����޸���������ɫΪ��ɫ
    With vsfParameter
        If .TextMatrix(mPara.mint������¿��ÿ��, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mint������������������, 0, mPara.mint������������������, .Cols - 1) = mlngColor
        End If
        If .TextMatrix(mintʱ���������ȡ�ϴ��ۼ�, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mintʱ�����������ԼӼ������, 1) = ""
            .TextMatrix(mPara.mint���İ��ֶμӳ������, 1) = ""
        End If
        If .TextMatrix(mintʱ�����������ԼӼ������, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mint���İ��ֶμӳ������, 0, mPara.mint���İ��ֶμӳ������, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mintʱ���������ȡ�ϴ��ۼ�, 1) = ""
            .TextMatrix(mPara.mint���İ��ֶμӳ������, 1) = ""
        End If
        If .TextMatrix(mint���İ��ֶμӳ������, 1) = "1" Then
            .Cell(flexcpBackColor, mPara.mintʱ�����������ԼӼ������, 0, mPara.mintʱ�����������ԼӼ������, .Cols - 1) = mlngColor
            .Cell(flexcpBackColor, mPara.mintʱ���������ȡ�ϴ��ۼ�, 0, mPara.mintʱ���������ȡ�ϴ��ۼ�, .Cols - 1) = mlngColor
            .TextMatrix(mPara.mintʱ�����������ԼӼ������, 1) = ""
            .TextMatrix(mPara.mintʱ���������ȡ�ϴ��ۼ�, 1) = ""
        End If
    End With
End Sub

Private Sub vsf����_EnterCell()
    Dim strTemp As String
    
    With vsf����
        If .Col = m�ⷿ����.mint���� Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = 1 Then
            mrs����.Filter = "��������='���ϲ���'"
        ElseIf .Col = 3 Then
            mrs����.Filter = "��������='���Ŀ�'"
        ElseIf .Col = 5 Then
            mrs����.Filter = "��������='����ⷿ'"
        End If
        
'        .Clear
        strTemp = ""
        Do While Not mrs����.EOF
            strTemp = strTemp & mrs����("����") & "-" & mrs����("����") & "|"
'            .AddItem mrs����("����") & "-" & mrs����("����")
'            .ItemData(.NewIndex) = mrs����("ID")
            mrs����.MoveNext
        Loop
        .ColComboList(.Col) = strTemp
    End With
End Sub

Private Sub vsf����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf����
            If .Col = m�ⷿ����.mint���� - 1 And .TextMatrix(.Row, m�ⷿ����.mint���ϲ���) <> "" And .TextMatrix(.Row, m�ⷿ����.mint���Ĳֿ�) <> "" And .TextMatrix(.Row, m�ⷿ����.mint����ⷿ) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m�ⷿ����.mint���� - 1 And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 2
            End If
        End With
    End If
End Sub

Private Sub vsf�ⷿ���_DblClick()
    With vsf�ⷿ���
        If .Col = m�ⷿ���.mint��鷽ʽ Then
            Select Case .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ)
                Case "0-�����"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "1-��飬��������"
                Case "1-��飬��������"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "2-��飬�����ֹ"
                Case "2-��飬�����ֹ"
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "0-�����"
                Case Else
                    .TextMatrix(.Row, m�ⷿ���.mint��鷽ʽ) = "0-�����"
            End Select
        End If
    End With
End Sub

Private Sub vsf����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strID As String
    Dim str���� As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    With vsf����
        strTemp = .TextMatrix(Row, Col)
        If strTemp <> "" Then
            If Col = m����.mint���ڿⷿ Then
                gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str���� = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID, str����)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, m����.mint���ڿⷿid) = rsTemp!Id
                End If
            ElseIf Col = m����.mint�Է��ⷿ Then
                strID = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
                str���� = Mid(strTemp, InStr(1, strTemp, "-") + 1)
                gstrSQL = "select id from ���ű� where ����=[1] and ����=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڿⷿ��ѯ", strID, str����)
                If rsTemp.RecordCount > 0 Then
                    .TextMatrix(Row, m����.mint�Է��ⷿid) = rsTemp!Id
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub vsf����_DblClick()
    With vsf����
        If .Col = m����.mint���� Then
            If .MouseRow = 0 Then Exit Sub
            .Editable = flexEDNone
            Select Case Left(.TextMatrix(.Row, .Col), 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                Case "2"
                    .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                Case Else
                    .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
            End Select
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsf����.Rows > 1 Then
        vsf����.RemoveItem vsf����.Row
    End If
End Sub

Private Sub vsf����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With vsf����
            If .Col = m����.mint���� And .TextMatrix(.Row, m����.mint���ڿⷿ) <> "" And .TextMatrix(.Row, m����.mint�Է��ⷿ) <> "" And .TextMatrix(.Row, m����.mint����) <> "" Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 1
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            ElseIf .Col < m����.mint���� And .TextMatrix(.Row, .Col) <> "" Then
                .Col = .Col + 1
            End If
        End With
    End If
End Sub





