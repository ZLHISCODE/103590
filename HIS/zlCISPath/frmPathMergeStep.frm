VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathMergeStep 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "�ϲ�·���׶�"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7700
   ScaleMode       =   0  'User
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   488
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11805
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7335
      Width           =   11805
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   10560
         TabIndex        =   17
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   9360
         TabIndex        =   16
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   11760
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   11760
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VB.Frame fraMerge 
      Caption         =   "�ϲ�·��5"
      Height          =   1400
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   11535
      Begin VSFlex8Ctl.VSFlexGrid vsPhase 
         Height          =   705
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   11295
         _cx             =   19923
         _cy             =   1244
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   15724768
         BackColorSel    =   15597549
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   32768
         FloodColor      =   192
         SheetBorder     =   15724768
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   450
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathMergeStep.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.TabStrip tabBranch 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��·��"
               Key             =   "_0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMerge 
      Caption         =   "�ϲ�·��4"
      Height          =   1400
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   11535
      Begin VSFlex8Ctl.VSFlexGrid vsPhase 
         Height          =   705
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   570
         Width           =   11295
         _cx             =   19923
         _cy             =   1244
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   15724768
         BackColorSel    =   15597549
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   32768
         FloodColor      =   192
         SheetBorder     =   15724768
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   450
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathMergeStep.frx":0095
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.TabStrip tabBranch 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��·��"
               Key             =   "_0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMerge 
      Caption         =   "�ϲ�·��3"
      Height          =   1400
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   11535
      Begin VSFlex8Ctl.VSFlexGrid vsPhase 
         Height          =   705
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   11295
         _cx             =   19923
         _cy             =   1244
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   15724768
         BackColorSel    =   15597549
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   32768
         FloodColor      =   192
         SheetBorder     =   15724768
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   450
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathMergeStep.frx":012A
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.TabStrip tabBranch 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��·��"
               Key             =   "_0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMerge 
      Caption         =   "�ϲ�·��1"
      Height          =   1400
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VSFlex8Ctl.VSFlexGrid vsPhase 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   570
         Width           =   11295
         _cx             =   19923
         _cy             =   1244
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   15724768
         BackColorSel    =   15597549
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   32768
         FloodColor      =   192
         SheetBorder     =   15724768
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   450
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathMergeStep.frx":01BF
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.TabStrip tabBranch 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��·��"
               Key             =   "_0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraMerge 
      Caption         =   "�ϲ�·��2"
      Height          =   1400
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   11535
      Begin VSFlex8Ctl.VSFlexGrid vsPhase 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   11295
         _cx             =   19923
         _cy             =   1244
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   15724768
         BackColorSel    =   15597549
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   32768
         GridColorFixed  =   32768
         TreeColor       =   32768
         FloodColor      =   192
         SheetBorder     =   15724768
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   2
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   450
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathMergeStep.frx":0254
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.TabStrip tabBranch 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "��·��"
               Key             =   "_0"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPathMergeStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsMerge As Recordset
Private mblnOK As Boolean
Private mlngMergeCount As Long   '�ϲ�·������
Private mstrMerge As String      '��������������ѡ���Ľ�����Ա��´��ٴ��룬ѡ��Ĭ�ϵĽ׶�

Public Function ShowMe(frmParent As Object, ByVal rsMerge As Recordset, ByVal lngMergeCount As Long, ByRef strMerge As String) As Boolean
'������mrsMerge=���кϲ�·���Ľ׶μ���
    
    Set mrsMerge = rsMerge
    mlngMergeCount = lngMergeCount
    mstrMerge = strMerge
    mblnOK = False
    Me.Show 1, frmParent
    strMerge = mstrMerge
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strFilter As String
    
    mstrMerge = ""
    For i = 0 To mlngMergeCount - 1
        strFilter = strFilter & " Or ID=" & vsPhase(i).ColData(vsPhase(i).Col)
        mstrMerge = mstrMerge & "," & Mid(tabBranch(i).SelectedItem.Key, 2) & ":" & vsPhase(i).ColData(vsPhase(i).Col)
    Next
    mstrMerge = Mid(mstrMerge, 2)
    mrsMerge.Filter = Mid(strFilter, 5)
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strTmp As String
    Dim intCount As Integer
    Dim lng·��ID As Long
    Dim lng�汾�� As Long
    Dim lng��ǰ�׶�ID As Long
    Dim str·������ As String
    
    mrsMerge.Filter = 0
    If mrsMerge.RecordCount > 0 Then
        mrsMerge.MoveFirst
        For i = 1 To mrsMerge.RecordCount
            If Val(mrsMerge!·��ID & "") <> lng·��ID And lng·��ID <> 0 Then
                strTmp = strTmp & ","
                Call LoadBranch(intCount, lng·��ID, lng�汾��, lng��ǰ�׶�ID, strTmp)
                Call LoadPhase(intCount, lng·��ID, lng��ǰ�׶�ID)
                lng·��ID = 0
                lng�汾�� = 0
                lng��ǰ�׶�ID = 0
                strTmp = ""
                fraMerge(intCount).Caption = str·������
                intCount = intCount + 1
            End If
            If Val(mrsMerge!��֧ID & "") <> 0 Then
                strTmp = strTmp & "," & mrsMerge!��֧ID
            End If
            lng·��ID = Val(mrsMerge!·��ID & "")
            lng�汾�� = Val(mrsMerge!�汾�� & "")
            lng��ǰ�׶�ID = Val(mrsMerge!��ǰ�׶�ID & "")
            str·������ = mrsMerge!·������ & ""
            If i = mrsMerge.RecordCount Then
                strTmp = strTmp & ","
                Call LoadBranch(intCount, lng·��ID, lng�汾��, lng��ǰ�׶�ID, strTmp)
                Call LoadPhase(intCount, lng·��ID, lng��ǰ�׶�ID)
                fraMerge(intCount).Caption = str·������
            End If
            
            mrsMerge.MoveNext
        Next
        mrsMerge.MoveFirst
    End If
    
    For i = 4 To mlngMergeCount Step -1
        fraMerge(i).Visible = False
    Next
    For i = 1 To mlngMergeCount - 1
        fraMerge(i).Top = fraMerge(i - 1).Top + fraMerge(i - 1).Height + 45
    Next
   
    Me.Height = fraMerge(mlngMergeCount - 1).Top + fraMerge(mlngMergeCount - 1).Height + picBottom.Height + 555
    
End Sub

Private Sub LoadBranch(ByVal Index As Integer, ByVal lng·��ID As Long, ByVal lng�汾�� As Long, ByVal lngǰһ�׶�ID As Long, ByVal strBranch As String)
'���ܣ����ط�֧·��
'������strBranch=���õķ�֧·��IDs
    Dim i As Long, j As Long, strSQL As String
    Dim rstmp As ADODB.Recordset
    
    strSQL = "Select ID,���� From �ٴ�·����֧ Where ǰһ�׶�ID=[3] And ·��ID=[1] And �汾��=[2]"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��֧��Ϣ", lng·��ID, lng�汾��, lngǰһ�׶�ID)
    If rstmp.RecordCount > 0 Then
        Do While Not rstmp.EOF
            '���ܷ�֧·���Ľ׶���û���ʺϵ�ǰ�����Ľ׶�
            If InStr(strBranch, "," & rstmp!ID & ",") > 0 Then
                tabBranch(Index).Tabs.Add , "_" & rstmp!ID, "��֧:" & rstmp!����
            End If
            'Ĭ��ѡ���Ѿ����صĽ׶�
            If InStr("," & mstrMerge, "," & rstmp!ID & ":") > 0 Then
                tabBranch(Index).Tabs.Item("_" & rstmp!ID).Selected = True
            End If
            rstmp.MoveNext
        Loop
        '���ڷ�֧
        fraMerge(Index).Tag = "1"
    End If
    If tabBranch(Index).Tabs.count = 1 Then
        tabBranch(Index).Visible = False
        fraMerge(Index).Tag = ""
        vsPhase(Index).Top = tabBranch(Index).Top
        fraMerge(Index).Height = vsPhase(Index).Top + vsPhase(Index).Height + 125
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPhase(ByVal Index As Integer, ByVal lng·��ID As Long, ByVal lngǰһ�׶�ID As Long)
'���ܣ����ؿ�ѡ��Ľ׶�,������˵ĵ�ǰʱ��׶���Ȼ���ã���ѡ�У�����ȱʡΪ��һ��
    Dim i As Long, j As Long, str�׶η��� As String
    Dim rsNode As ADODB.Recordset
    Dim lngRecord As Long

    With vsPhase(Index)
        .Clear
        .Redraw = flexRDNone
        .Col = -1
        lngRecord = mrsMerge.AbsolutePosition
        mrsMerge.Filter = "·��ID=" & lng·��ID & IIf(fraMerge(Index).Tag = "1", " And ��֧ID=" & Mid(tabBranch(Index).SelectedItem.Key, 2), "")
        .Cols = mrsMerge.RecordCount
        str�׶η��� = Get�׶η���(0, lngǰһ�׶�ID)

        For i = 0 To .Cols - 1
            .ColWidth(i) = 2000
            .ColAlignment(i) = flexAlignCenterCenter
            .TextMatrix(0, i) = mrsMerge!����
            .Cell(flexcpData, 0, i) = CStr(IIf(IsNull(mrsMerge!����), "", "���ࣺ" & mrsMerge!���� & " ") & mrsMerge!˵��)
            .ColData(i) = Val(mrsMerge!ID)
            If .ColData(i) = lngǰһ�׶�ID Then .Col = i
            If InStr("," & mstrMerge & ",", "," & mrsMerge!��֧ID & ":" & .ColData(i) & ",") > 0 Then
                .Col = i
            End If

            If Not rsNode Is Nothing Then
                rsNode.Filter = "��ID=" & mrsMerge!ID
                If rsNode.RecordCount = 0 Then
                     .MergeCol(i) = True
                     .TextMatrix(1, i) = mrsMerge!����
                Else
                     .TextMatrix(1, i) = "ȱʡ"
                     .ColWidth(i) = 1000
                    For j = 1 To rsNode.RecordCount
                        i = i + 1
                         .ColWidth(i) = 1000
                         .ColAlignment(i) = flexAlignCenterCenter
                        .TextMatrix(0, i) = mrsMerge!���� '��һ��������ͬ�������ںϲ�
                        .TextMatrix(1, i) = IIf(IsNull(rsNode!˵��), "��֧" & j, "" & rsNode!˵��)
                        .Cell(flexcpData, 1, i) = CStr(IIf(IsNull(rsNode!����), "", "���ࣺ" & rsNode!���� & " ") & rsNode!˵��)

                        .ColData(i) = Val(rsNode!ID)
                        If .ColData(i) = lngǰһ�׶�ID Then
                            .Col = i
                        ElseIf .Col = 0 And str�׶η��� <> "" Then
                            If str�׶η��� = "" & rsNode!���� Then .Col = i
                        End If
                        rsNode.MoveNext
                    Next
                End If
            End If

            mrsMerge.MoveNext
        Next

        If .Col < 0 Then .Col = 0
        mrsMerge.Filter = 0
        mrsMerge.AbsolutePosition = lngRecord
        .Redraw = True
    End With
End Sub

Private Sub tabBranch_Click(Index As Integer)
    If vsPhase(Index).ColData(vsPhase(Index).Col) <> 0 Then
        mrsMerge.Filter = "ID=" & vsPhase(Index).ColData(vsPhase(Index).Col)
        If mrsMerge.RecordCount > 0 Then
            Call LoadPhase(Index, Val(mrsMerge!·��ID & ""), Val(mrsMerge!��ǰ�׶�ID & ""))
        End If
    End If
End Sub
