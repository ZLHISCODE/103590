VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmBalanceTabNormal 
   BorderStyle     =   0  'None
   Caption         =   "frmBalanceTabNormal"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBalanceInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   7440
      ScaleHeight     =   675
      ScaleWidth      =   2715
      TabIndex        =   11
      Top             =   5655
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalanceInfo 
         Height          =   1845
         Left            =   0
         TabIndex        =   5
         Top             =   15
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   7455
      ScaleHeight     =   1065
      ScaleWidth      =   2550
      TabIndex        =   10
      Top             =   4335
      Width           =   2550
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
         Height          =   1845
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
   Begin VB.PictureBox picInvoice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   7515
      ScaleHeight     =   945
      ScaleWidth      =   2415
      TabIndex        =   9
      Top             =   3135
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1845
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   4395
      ScaleHeight     =   2115
      ScaleWidth      =   2745
      TabIndex        =   7
      Top             =   3420
      Width           =   2745
      Begin XtremeSuiteControls.TabControl tabInfo 
         Height          =   2010
         Left            =   -705
         TabIndex        =   8
         Top             =   -375
         Width           =   2820
         _Version        =   589884
         _ExtentX        =   4974
         _ExtentY        =   3545
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   1065
      ScaleHeight     =   2805
      ScaleWidth      =   2715
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3345
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Width           =   1800
         _cx             =   3175
         _cy             =   3254
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
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
         ExplorerBar     =   3
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3015
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   3120
      Begin VB.PictureBox picMainControl 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   540
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   12
         Top             =   300
         Width           =   210
         Begin VB.Image imgMainControl 
            Height          =   195
            Left            =   0
            Picture         =   "frmBalanceTabNormal.frx":0000
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1800
         Left            =   510
         TabIndex        =   1
         Top             =   270
         Width           =   2550
         _cx             =   4498
         _cy             =   3175
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
         BackColorSel    =   16772055
         ForeColorSel    =   8
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBalanceTabNormal.frx":054E
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
         ExplorerBar     =   3
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
      Left            =   1725
      Top             =   1710
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBalanceTabNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnPrint As Boolean
Private mfrmFilter As New frmBalanceFilter
Private mblnNOMoved As Boolean

Public Sub MakeFilter(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String)
    Call mfrmFilter.InitFilter(Me, lngModul, strPrivs)
End Sub

Public Sub ReadData(ByVal intTYPE As Integer, ByVal strPrivs As String, Optional ByVal lngPatiID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ʼ�¼
    '����:������
    '���:intType-��ȡ��¼�ķ�ʽ��0Ϊʹ�ù���������ȡ��1Ϊʹ��IDKIND������ȡ
    '����:2015-01-06
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset, strTable As String
    Dim strFilter As String, strInvoice As String, strSQLtmp As String
    Dim DatBegin As Date, DatEnd As Date, blnMoved As Boolean, strSource As String
    Dim i As Integer, str��Դ As String, strUpgrade As String
    On Error GoTo ErrHand
    If intTYPE = 0 Then
        If mfrmFilter.mblnInit = True Then
            With mfrmFilter
                DatBegin = .dtpBegin.Value
                DatEnd = .dtpEnd.Value
                blnMoved = zlDatabase.DateMoved(IIf(DatBegin < DatEnd, DatBegin, DatEnd))
                strFilter = " And A.�շ�ʱ�� Between [1] And [2] "
                strFilter = strFilter & IIf(.txt����.Text = "", "", " And C.����=[3] ")
                strFilter = strFilter & IIf(.cbo����Ա.Text = "���н�����", "", " And A.����Ա����=[4] ")
                strFilter = strFilter & IIf(.txt�����.Text = "", "", " And C.�����=[5] ")
                strFilter = strFilter & IIf(.txtסԺ��.Text = "", "", " And C.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[6]) ")
                If Not (.chkType(0).Value = 1 And .chkType(1).Value = 1) Then
                    If .chkType(0).Value = 1 Then
                        strFilter = strFilter & " And A.��¼״̬ In (1,3) "
                    Else
                        strFilter = strFilter & " And A.��¼״̬ = 2 "
                    End If
                End If
                If .txtNOBegin.Text <> "" Then
                    If .txtNoEnd.Text <> "" Then
                        strFilter = strFilter & " And A.NO Between [7] And [8] "
                    Else
                        strFilter = strFilter & " And A.NO=[7] "
                    End If
                End If
                strInvoice = ""
                If (.txtFactBegin.Text <> "" And .txtFactEnd.Text <> "") Or (.txtFactBegin.Text <> "" And .txtFactEnd.Text = "") Then
                    '�������Ʊ�ݺ��ж�,ֱ�Ӹ��ݵ��ݵĵǼ�ʱ���ж�
                    strSQLtmp = IIf(.txtFactEnd.Text = "", " =[9] ", " Between [9] And [10] ")
                    If blnMoved Then
                        strInvoice = "" & _
                         "(  Select A.NO" & _
                         "   From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
                         "   Where A.��������=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.��ӡID And B.Ʊ��=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.����=1" & _
                         "         And B.���� " & strSQLtmp & ")  Union All" & _
                         " (Select A.NO " & _
                         " From HƱ�ݴ�ӡ���� A,HƱ��ʹ����ϸ B" & _
                         " Where A.��������=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.��ӡID And B.Ʊ��=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.����=1" & _
                         " And B.���� " & strSQLtmp & ")"
                    Else
                        strInvoice = "Select A.NO" & _
                        " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
                        " Where A.��������=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And A.ID=B.��ӡID And B.Ʊ��=" & IIf(gbytInvoiceKind = 0, 3, 1) & " And B.����=1" & _
                        " And B.���� " & strSQLtmp
                    End If
                End If
                If strInvoice <> "" Then strFilter = strFilter & " And A.NO In (" & strInvoice & ") "
                
                For i = 0 To .chkFeeOrigin.Count - 1
                    strSource = strSource & IIf(.chkFeeOrigin(i).Value = 1, 1, 0) '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
                Next
                If strSource = "" Then strSource = "0100"
                str��Դ = ""
                For i = 1 To Len(strSource)
                    If Mid(strSource, i, 1) = 1 Then
                        str��Դ = str��Դ & "," & Choose(i, 1, 2, 4, 3)  '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
                    End If
                Next
                If str��Դ <> "" Then str��Դ = Mid(str��Դ, 2)
                If str��Դ = "" Then str��Դ = "-1"
                
                strTable = "" & _
                "   Select A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID," & _
                "           B.����ID as ���ò���ID,Nvl(D.�Ա�,C.�Ա�) as �Ա�,Nvl(D.����,C.����) as ����,A.��ʼ����,A.��������,Max(A.��¼״̬) As ��¼״̬,Sum(B.���ʽ��) As ���ʽ��," & _
                "           A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ�� as ��Լ��λ,A.��������,A.��ҳID,Max(Decode(a.���ʽ��,Null,1,0)) As ��Ҫ����" & _
                "   From ���˽��ʼ�¼ A,סԺ���ü�¼ B,������Ϣ C,������ҳ D " & _
                "   Where A.ID=B.����ID and  B.����ID=C.����ID And A.����ID =D.����ID(+) And A.��ҳID = D.��ҳID(+) And (A.����״̬ = 2 Or A.����״̬ Is Null) " & _
                        IIf(strSource = "1111", "", " And Instr(',' || [11] || ',',',' || Nvl(B.�����־,0) || ',') > 0 ") & strFilter & _
                "   Group By A.ID,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID,Nvl(D.�Ա�,C.�Ա�),Nvl(D.����,C.����),A.��ʼ����,A.��������,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ��,A.��������,A.��ҳID "

                Select Case strSource
                Case "1010", "1000", "0010"  '����
                    strTable = Replace(strTable, "סԺ���ü�¼", "������ü�¼")
                    strTable = Replace(strTable, "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
                Case "0101", "0001", "0100" 'סԺ
                    '�Ѿ�����
                Case Else '�����סԺ
                    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
                End Select
                
                If blnMoved Then
                    strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "���˽��ʼ�¼", "H���˽��ʼ�¼"), "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
                End If
                
                strTable = "Select ID, Decode(Max(סԺ��־), 1, Decode(Max(�����־), 1, 3, 2), 1) As ��־, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������," & vbNewLine & _
                            "              ��¼״̬, Sum(���ʽ��) As ���ʽ��, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ����" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������, ��¼״̬, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ���� "
                
                strSql = _
                " Select A.ID ����ID,��־,decode(A.��������,1,'�������',2,'סԺ����','') As �������� ,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
                "        Decode(A.����ID,Null,' ',A.����ID) ����ID,Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)) �����,Decode(A.����ID,Null,' ',Decode(A.��ҳID,Null,C.סԺ��,P.סԺ��)) סԺ��," & _
                "        Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����) ����,Decode(A.����ID,Null,' ',A.�Ա�) �Ա�," & _
                "        Decode(A.����ID,Null,' ',A.����) ����,Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�)) as �ѱ�," & _
                "        To_Char(A.��ʼ����,'YYYY-MM-DD') as ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
                "        To_Char(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��,'999999999" & gstrDec & "') as ���ʽ��," & _
                "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,A.��¼״̬ as ��¼״̬,A.��Ҫ����" & _
                " From ( " & strTable & ") A,������Ϣ C,������ҳ P,��Լ��λ Q,��Ա�� N" & _
                " Where  A.���ò���ID=C.����ID And A.����Ա����=N.���� " & _
                "        And A.���ò���ID=P.����ID(+) And Nvl(A.��ҳID,0)=P.��ҳID(+) And C.��ͬ��λID=Q.ID(+)" & _
                "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null)" & vbNewLine
                
                strSql = strSql & " Order by �շ�ʱ�� Desc,���ݺ� Desc"
                
                Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, _
                                                    .txt����.Text, zlStr.NeedName(.cbo����Ա.Text), .txt�����.Text, .txtסԺ��.Text, _
                                                    .txtNOBegin.Text, .txtNoEnd.Text, .txtFactBegin.Text, .txtFactEnd.Text, str��Դ)
                Do While Not rsMain.EOF
                    If Val(NVL(rsMain!��Ҫ����)) = 1 Then
                        strUpgrade = "Zl_���˽��ʼ�¼_Upgrade(" & rsMain!����ID & ")"
                        zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
                    End If
                    rsMain.MoveNext
                Loop
                If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
                Set vsfMain.DataSource = rsMain
                
                Call SetMain
            End With
        Else
            DatBegin = Format(zlDatabase.Currentdate, "YYYY-MM-DD 00:00:00")
            DatEnd = Format(zlDatabase.Currentdate, "YYYY-MM-DD 23:59:59")
            strFilter = " And A.�շ�ʱ�� Between [1] And [2] "
            strFilter = strFilter & " And A.����Ա����=[3] "
            strSource = "1111"
            str��Դ = ""
            For i = 1 To Len(strSource)
                If Mid(strSource, i, 1) = 1 Then
                    str��Դ = str��Դ & "," & Choose(i, 1, 2, 4, 3)  '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
                End If
            Next
            If str��Դ <> "" Then str��Դ = Mid(str��Դ, 2)
            If str��Դ = "" Then str��Դ = "-1"
            
            strTable = "" & _
            "   Select " & IIf(strSource = "1111", "", " /*+cardinality(L1,10)*/ ") & " A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID as ���ò���ID,Nvl(D.�Ա�,C.�Ա�) as �Ա�,Nvl(D.����,C.����) as ����,A.��ʼ����,A.��������,Max(A.��¼״̬) As ��¼״̬,Sum(B.���ʽ��) As ���ʽ��,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ�� as ��Լ��λ,A.��������,A.��ҳID,Max(Decode(a.���ʽ��,Null,1,0)) As ��Ҫ���� " & _
            "   From ���˽��ʼ�¼ A,סԺ���ü�¼ B,������Ϣ C,������ҳ D " & _
                    IIf(strSource = "1111", "", ",Table(Cast(f_Num2list('" & str��Դ & "') As Zltools.t_Numlist)) L1") & _
            "   Where A.ID=B.����ID and  B.����ID=C.����ID And A.����ID =D.����ID(+) And A.��ҳID = D.��ҳID(+) And (A.����״̬ = 2 Or A.����״̬ Is Null) " & _
                    IIf(strSource = "1111", "", " And nvl(B.�����־,0)=L1.Column_Value ") & strFilter & _
            "   Group By A.ID,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID,Nvl(D.�Ա�,C.�Ա�),Nvl(D.����,C.����),A.��ʼ����,A.��������,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ��,A.��������,A.��ҳID "
                    
            Select Case strSource
            Case "1010", "1000", "0010"  '����
                strTable = Replace(strTable, "סԺ���ü�¼", "������ü�¼")
                strTable = Replace(strTable, "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
            Case "0101", "0001", "0100" 'סԺ
                '�Ѿ�����
            Case Else '�����סԺ
                strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
            End Select
            
            strTable = "Select ID, Decode(Max(סԺ��־), 1, Decode(Max(�����־), 1, 3, 2), 1) As ��־, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������," & vbNewLine & _
                            "              ��¼״̬, Sum(���ʽ��) As ���ʽ��, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ����" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������, ��¼״̬, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ���� "
                
            strSql = _
                " Select A.ID ����ID,��־,decode(A.��������,1,'�������',2,'סԺ����','') As �������� ,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
                "        Decode(A.����ID,Null,' ',A.����ID) ����ID,Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)) �����,Decode(A.����ID,Null,' ',Decode(A.��ҳID,Null,C.סԺ��,P.סԺ��)) סԺ��," & _
                "        Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����) ����,Decode(A.����ID,Null,' ',A.�Ա�) �Ա�," & _
                "        Decode(A.����ID,Null,' ',A.����) ����,Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�)) as �ѱ�," & _
                "        To_Char(A.��ʼ����,'YYYY-MM-DD') as ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
                "        To_Char(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��,'999999999" & gstrDec & "') as ���ʽ��," & _
                "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,A.��¼״̬ as ��¼״̬,A.��Ҫ����" & _
                " From ( " & strTable & ") A,������Ϣ C,������ҳ P,��Լ��λ Q,��Ա�� N" & _
                " Where  A.���ò���ID=C.����ID And A.����Ա����=N.���� " & _
                "        And A.���ò���ID=P.����ID(+) And Nvl(A.��ҳID,0)=P.��ҳID(+) And C.��ͬ��λID=Q.ID(+)" & _
                "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null)" & vbNewLine
            strSql = strSql & " Order by �շ�ʱ�� Desc,���ݺ� Desc"
            
            Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, DatBegin, DatEnd, UserInfo.����)
            Do While Not rsMain.EOF
                If Val(NVL(rsMain!��Ҫ����)) = 1 Then
                    strUpgrade = "Zl_���˽��ʼ�¼_Upgrade(" & rsMain!����ID & ")"
                    zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
                End If
                rsMain.MoveNext
            Loop
            If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
            Set vsfMain.DataSource = rsMain
            Call SetMain
        End If
    End If
    
    If intTYPE = 1 Then
        strFilter = " And C.����ID = [1]  "
        strSource = "1111"
        str��Դ = ""
        For i = 1 To Len(strSource)
            If Mid(strSource, i, 1) = 1 Then
                str��Դ = str��Դ & "," & Choose(i, 1, 2, 4, 3)  '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
            End If
        Next
        If str��Դ <> "" Then str��Դ = Mid(str��Դ, 2)
        If str��Դ = "" Then str��Դ = "-1"
        
        strTable = "" & _
        "   Select " & IIf(strSource = "1111", "", " /*+cardinality(L1,10)*/ ") & " A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID as ���ò���ID,Nvl(D.�Ա�,C.�Ա�) as �Ա�,Nvl(D.����,C.����) as ����,A.��ʼ����,A.��������,A.��¼״̬,B.���ʽ��,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ�� as ��Լ��λ,A.��������,A.��ҳID,Decode(a.���ʽ��,Null,1,0) As ��Ҫ���� " & _
        "   From ���˽��ʼ�¼ A,סԺ���ü�¼ B,������Ϣ C,������ҳ D " & _
                IIf(strSource = "1111", "", ",Table(Cast(f_Num2list([11]) As Zltools.t_Numlist)) L1") & _
        "   Where A.ID=B.����ID and  B.����ID=C.����ID And A.����ID =D.����ID(+) And A.��ҳID = D.��ҳID(+) And Nvl(A.����״̬,2) = 2 " & _
                IIf(strSource = "1111", "", " And nvl(B.�����־,0)=L1.Column_Value ") & strFilter
                
        Select Case strSource
        Case "1010", "1000", "0010"  '����
            strTable = Replace(strTable, "סԺ���ü�¼", "������ü�¼")
            strTable = Replace(strTable, "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
        Case "0101", "0001", "0100" 'סԺ
            '�Ѿ�����
        Case Else '�����סԺ
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
        End Select
        strTable = "Select ID, Decode(Max(סԺ��־), 1, Decode(Max(�����־), 1, 3, 2), 1) As ��־, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������," & vbNewLine & _
                            "              ��¼״̬, Sum(���ʽ��) As ���ʽ��, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ����" & vbNewLine & _
                            "       From (" & strTable & ") " & _
                            "       Group By ID, NO, ʵ��Ʊ��, ����id, ���ò���id, �Ա�, ����, ��ʼ����, ��������, ��¼״̬, ����Ա����, �շ�ʱ��, ��;����, ��Լ��λ, ��������, ��ҳid, ��Ҫ���� "
                
        'ʹ��IDKIND������ȡ
        strSql = _
                " Select A.ID ����ID,��־,decode(A.��������,1,'�������',2,'סԺ����','') As �������� ,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
                "        Decode(A.����ID,Null,' ',A.����ID) ����ID,Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)) �����,Decode(A.����ID,Null,' ',Decode(A.��ҳID,Null,C.סԺ��,P.סԺ��)) סԺ��," & _
                "        Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����) ����,Decode(A.����ID,Null,' ',A.�Ա�) �Ա�," & _
                "        Decode(A.����ID,Null,' ',A.����) ����,Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�)) as �ѱ�," & _
                "        To_Char(A.��ʼ����,'YYYY-MM-DD') as ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
                "        To_Char(Sum(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��),'999999999" & gstrDec & "') as ���ʽ��," & _
                "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,Max(A.��¼״̬) as ��¼״̬,A.��Ҫ����" & _
                " From ( " & strTable & ") A,������Ϣ C,������ҳ P,��Լ��λ Q,��Ա�� N" & _
                " Where  A.���ò���ID=C.����ID And A.����Ա����=N.���� " & _
                "        And A.���ò���ID=P.����ID(+) And Nvl(A.��ҳID,0)=P.��ҳID(+) And C.��ͬ��λID=Q.ID(+)" & _
                "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null)" & vbNewLine & _
                " Group by A.ID,��־,Decode(a.��������, 1, '�������', 2, 'סԺ����', ''),Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��'),A.NO,A.ʵ��Ʊ��,Decode(A.����ID,Null,' ',A.����ID),Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)),Decode(A.����ID,Null,' ',Decode(A.��ҳID,Null,C.סԺ��,P.סԺ��))," & _
                "           Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����),Decode(A.����ID,Null,' ',A.�Ա�),Decode(A.����ID,Null,' ',A.����),Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�))," & _
                "           To_Char(A.��ʼ����,'YYYY-MM-DD'),To_Char(A.��������,'YYYY-MM-DD')," & _
                "           A.��Ҫ����,A.����Ա����,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),Decode(Nvl(A.��;����,0),1,'��',' ')"
        strSql = strSql & " Order by �շ�ʱ�� Desc,���ݺ� Desc"
        
        Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID)
        Do While Not rsMain.EOF
            If Val(NVL(rsMain!��Ҫ����)) = 1 Then
                strUpgrade = "Zl_���˽��ʼ�¼_Upgrade(" & rsMain!����ID & ")"
                zlDatabase.ExecuteProcedure strUpgrade, Me.Caption
            End If
            rsMain.MoveNext
        Loop
        If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
        Set vsfMain.DataSource = rsMain
        Call SetMain
    End If
    
    If Not rsMain Is Nothing Then
        If rsMain.RecordCount <> 0 Then
            frmManageBalance.stbThis.Panels(2).Text = "��ǰ����" & rsMain.RecordCount & "�����˼�¼,�ϼ�:" & Format(GetTotal, gstrDec) & "Ԫ"
        Else
            frmManageBalance.stbThis.Panels(2).Text = ""
        End If
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '����:����DOCKINGPANEL�ؼ�
    '����:������
    '����:2013-09-04
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(1, 2000, 4000, DockTopOf)
        objPanel.Handle = picMain.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption

        Set objPanel = .CreatePane(2, 1700, 2000, DockBottomOf, objPanel)
        objPanel.Handle = picDetail.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption

        Set objPanel = .CreatePane(3, 1000, 2000, DockRightOf, objPanel)
        objPanel.Handle = picInfo.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetInvoiceList()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "ID,1,0|Ʊ�ݺ�,4,1000|ʹ��ԭ��,4,1000|ʹ��ʱ��,4,1200|ʹ����,1,1000"
    
    With vsfInvoice
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfInvoice, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        .Row = 1: .Col = 0: .ColSel = .Cols - 1

        .Redraw = True
    End With
End Sub

Private Function GetTotal() As Double
    Dim dblTotal As Double
    Dim i As Integer
    With vsfMain
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) <> 2 Then
                dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("���ʽ��")))
            Else
                dblTotal = dblTotal - Val(.TextMatrix(i, .ColIndex("���ʽ��")))
            End If
        Next i
    End With
    GetTotal = dblTotal
End Function

Private Sub SetMain()
    Dim i As Long, strHead As String
    Dim dblTotal As Double
    
    strHead = "����ID,1,0|��־,1,0|  ��������,4,1100|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,850|����ID,1,750|�����,1,750|סԺ��,1,750|����,4,800|�Ա�,4,500|����,4,500|�ѱ�,4,750|��ʼ����,4,1000|��������,4,1000|���ʽ��,7,850|����Ա,4,800|�շ�ʱ��,4,1850|��;����,4,800|��¼״̬,1,0"
    vsfBalance.Clear 1
    vsfInvoice.Clear 1
    vsfBalanceInfo.Clear 1
    vsfDetail.Clear 1
    With vsfMain
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            If .TextMatrix(0, i) = "����ID" Then .ColHidden(i) = True
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "����ID" Or .ColKey(i) = "��־" Or .ColKey(i) = "��¼״̬" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "���ݺ�" Or .ColKey(i) = "�շ�ʱ��" Or .ColKey(i) = "  ��������" Then .ColData(i) = "-1|1"
        Next
        
        zl_vsGrid_Para_Restore 1137, vsfMain, Me.Name, "������Ϣ�б�", False
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            ElseIf Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
            End If
            .RowHeight(i) = 300
        Next i
        
        .Select 1, 1

        .Redraw = True
    End With
    Call vsfMain_GotFocus
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "����,4,750|���ݺ�,4,850|��������,1,850|��Ŀ,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|��Ŀ,1,850|Ӥ����,4,650|���ʽ��,7,850|����ʱ��,1,1850"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsfMain.Cell(flexcpForeColor, vsfMain.Row, 1, vsfMain.Row, 1)
        Next i
        
        .Redraw = True
    End With
End Sub

Private Sub Form_Load()
    Call SetDockingPanel
    Call SetTab
    Call SetMain
    Call SetInvoiceList
    Call SetDetail
    Call SetExtendInfo
    Call SetBalanceList
End Sub

Private Sub SetBalanceList()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "����,4,800|���ݺ�,4,1000|���,7,1000|���㷽ʽ,1,1200|�������,1,1000"
    
    With vsfBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vsfMain.Cell(flexcpForeColor, vsfMain.Row, 1, vsfMain.Row, 1)
            .TextMatrix(i, .ColIndex("���")) = Formatex(Val(.TextMatrix(i, .ColIndex("���"))), 6, , , 2)
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        
        .Redraw = True
    End With
End Sub

Private Sub SetExtendInfo()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|���㷽ʽ,1,0|����,1,0|���,1,0|��Ŀ,1,1200|����,1,2000|������ˮ��,1,0"
    
    With vsfBalanceInfo
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "ID" Or .ColKey(i) = "������ˮ��" Or .ColKey(i) = "���㷽ʽ" Or .ColKey(i) = "����" Or .ColKey(i) = "���" Or .ColKey(i) = "λ��" Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        
        .RowHeight(0) = 350
        '.Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
        
        If .TextMatrix(1, 0) = "" Then Exit Sub

        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        .Subtotal flexSTNone, .ColIndex("ID"), .ColIndex("��Ŀ"), gstrDec, &H8000000F
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("��Ŀ")
        .OutlineCol = .ColIndex("��Ŀ")
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("��Ŀ")) = strTemp

                strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���㷽ʽ"))
                strTemp = strTemp & "(" & Format(.Cell(flexcpTextDisplay, i + 1, .ColIndex("���")), gstrDec) & ")"
                If .Cell(flexcpTextDisplay, i + 1, .ColIndex("������ˮ��")) <> "" Then
                   strTemp = strTemp & Space(1) & "������ˮ��:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������ˮ��"))
                End If
                
                .MergeRow(i) = True
                .MergeCells = flexMergeRestrictRows
                .Cell(flexcpAlignment, i, .ColIndex("��Ŀ"), i, .ColIndex("��Ŀ")) = 1
                
                For j = 0 To .Cols - 1
                   If j <= .ColIndex("����") Then
                       If j >= .ColIndex("��Ŀ") Then
                           .Cell(flexcpText, i, j) = strTemp
                           .Cell(flexcpFontBold, i, j) = False
                       End If
                   End If
                Next
            End If
        Next
        Call .AutoSize(.ColIndex("��Ŀ"))
        For j = 0 To .Cols - 1
            .MergeCol(j) = True
        Next
    End With
End Sub

Private Sub SetTab()
    On Error GoTo errHandle
    With tabInfo
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.HotTracking = True
        .PaintManager.Color = xtpTabColorOffice2003
        'Set .PaintManager.Font = txtSendFeeNO.Font
        .InsertItem 1, "Ʊ����Ϣ", picInvoice.hWnd, 0
        .InsertItem 2, "������Ϣ", picBalance.hWnd, 0
        .InsertItem 3, "���������Ϣ", picBalanceInfo.hWnd, 0
        .Item(0).Selected = True
        .Item(2).Visible = False
        .PaintManager.BoldSelected = True
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnPrint = False
    zl_vsGrid_Para_Save 1137, vsfMain, Me.Name, "������Ϣ�б�", False
End Sub

Private Sub picDetail_Resize()
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = picDetail.Height
        .Width = picDetail.Width
    End With
End Sub


Private Sub picInfo_Resize()
    With tabInfo
        .Top = 0
        .Left = 0
        .Width = picInfo.Width
        .Height = picInfo.Height
    End With
End Sub

Private Sub picInvoice_Resize()
    With vsfInvoice
        .Top = 0
        .Left = 0
        .Height = picInvoice.Height
        .Width = picInvoice.Width
    End With
End Sub

Private Sub picBalance_Resize()
    With vsfBalance
        .Top = 0
        .Left = 0
        .Height = picBalance.Height
        .Width = picBalance.Width
    End With
End Sub

Private Sub picBalanceInfo_Resize()
    With vsfBalanceInfo
        .Top = 0
        .Left = 0
        .Height = picBalanceInfo.Height
        .Width = picBalanceInfo.Width
    End With
End Sub

Private Sub picMain_Resize()
    With vsfMain
        .Top = 0
        .Left = 0
        .Height = picMain.Height
        .Width = picMain.Width
    End With
    With picMainControl
        .Top = 90
        .Left = 75
    End With
End Sub

Private Sub imgMainControl_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picMainControl.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picMainControl.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsfMain, lngLeft, lngTop, picMainControl.Height)
    zl_vsGrid_Para_Save 1137, vsfMain, Me.Name, "������Ϣ�б�", False
End Sub

Private Sub picMainControl_Click()
    Call imgMainControl_Click
End Sub

Private Sub vsfBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalance_GotFocus()
    zl_VsGridGotFocus vsfBalance, &HFFC0C0
End Sub

Private Sub vsfBalance_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalance, , vsfBalance.Cell(flexcpForeColor, vsfBalance.Row, vsfBalance.Col)
End Sub

Private Sub vsfBalanceInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfBalanceInfo, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfBalanceInfo_GotFocus()
    zl_VsGridGotFocus vsfBalanceInfo, &HFFC0C0
End Sub

Private Sub vsfBalanceInfo_LostFocus()
    zl_VsGridLOSTFOCUS vsfBalanceInfo
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfDetail, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfDetail_GotFocus()
    zl_VsGridGotFocus vsfDetail, &HFFC0C0
End Sub

Private Sub vsfDetail_LostFocus()
    zl_VsGridLOSTFOCUS vsfDetail, , vsfDetail.Cell(flexcpForeColor, vsfDetail.Row, vsfDetail.Col)
End Sub

Private Sub vsfInvoice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsfInvoice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsfInvoice_GotFocus()
    zl_VsGridGotFocus vsfInvoice, &HFFC0C0
End Sub

Private Sub vsfInvoice_LostFocus()
    zl_VsGridLOSTFOCUS vsfInvoice, , vsfInvoice.Cell(flexcpForeColor, vsfInvoice.Row, vsfInvoice.Col)
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    If NewRow <> 0 And OldRow <> 0 Then zl_VsGridRowChange vsfMain, OldRow, NewRow, OldCol, NewCol
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID"))))
    Call ReadInVoice(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("���ݺ�")))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID"))))
    Call ReadBalanceInfo(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����ID"))))
End Sub

Private Sub ReadBalance(ByVal lngBalanceID As Long)
    Dim strSql As String, rsBalance As ADODB.Recordset, blnDel As Boolean
    
    If mblnPrint Then Exit Sub
    
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��¼״̬"))) = 2
    strSql = _
            "Select Decode(Substr(��¼����,Length(��¼����),1),1,'��Ԥ��',2,'����') as ����," & _
            " NO as ���ݺ�," & IIf(blnDel, "-1*", "") & "��Ԥ�� as ���," & _
            " ���㷽ʽ,������� From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ Where ����ID=[1] And ��Ԥ�� <> 0 " & _
            " Order by ���� Desc,NO Desc,���㷽ʽ"

    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfBalance.Redraw = False
    vsfBalance.Clear
    vsfBalance.Rows = 2
    If Not rsBalance.EOF Then
        Set vsfBalance.DataSource = rsBalance
    End If
    Call SetBalanceList
    vsfBalance.Redraw = True
End Sub

Private Sub ReadBalanceInfo(ByVal lngBalanceID As Long)
    Dim strSql As String, rsInfo As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
        "Select b.����id || '_' || b.ԭԤ��id As ID, a.���㷽ʽ, Max(c.����) As ����, Sum(Nvl(-1 * f.���, a.��Ԥ��)) As ���," & vbNewLine & _
        "       b.������Ŀ, b.��������, Max(Nvl(f.������ˮ��, a.������ˮ��)) As ������ˮ��" & vbNewLine & _
        "From ����Ԥ����¼ A, �������㽻�� B, ҽ�ƿ���� C, ����Ԥ����¼ E, �����˿���Ϣ F" & vbNewLine & _
        "Where a.Id = b.����id And a.�����id = c.Id(+) And a.����id = [1] And a.��¼���� <> 1" & vbNewLine & _
        "      And b.ԭԤ��id = e.Id(+) And e.id = f.��¼id(+) And f.����id(+) =  [1]" & vbNewLine & _
        "Group By b.����id, b.ԭԤ��id, a.���㷽ʽ, b.������Ŀ, b.��������" & vbNewLine & _
        "Order By ID"
    Set rsInfo = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    Set vsfBalanceInfo.DataSource = rsInfo
    If rsInfo.RecordCount = 0 Then
        'û�е��������׼�¼ʱ�����ط�ҳ
        tabInfo.Item(2).Visible = False
        If tabInfo.Selected.Index = 2 Then tabInfo.Item(0).Selected = True
    Else
        tabInfo.Item(2).Visible = True
    End If
    Call SetExtendInfo
End Sub

Private Sub ReadDetail(ByVal lngBalanceID As Long)
    Dim strSql As String, rsDetail As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnDel As Boolean, strDec As String, int��Դ As Integer
    
    If mblnPrint Then Exit Sub
    
    mblnNOMoved = zlDatabase.NOMoved("���˽��ʼ�¼", vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("���ݺ�")))
    int��Դ = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��־")))
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��¼״̬"))) = 2
    strDec = gstrDec
    If lngBalanceID <> 0 Then
        Select Case int��Դ
        Case 1 '����
            strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1]"
        Case 2 'סԺ
            strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
        Case Else
            
            strSql = "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))  as  declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1] Union ALL " & _
                     "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))   as  declen  From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
            strSql = "Select Max(declen)-1 as declen  From ( " & strSql & ")"
        End Select
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        If rsTmp.RecordCount > 0 Then
            If Len(strDec) < Len("0." & String(rsTmp!declen, "0")) Then
                strDec = "0." & String(rsTmp!declen, "0")
            End If
        End If
    End If
    
    Select Case int��Դ
    Case 1  '����
        strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,Sum(���ʽ��) As ���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] Group By ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,�վݷ�Ŀ,Ӥ����,����ʱ��) A "
        'strSQL = IIf(mblnNOMoved, "H", "") & "������ü�¼ A "
    Case 2  'סԺ
        strSql = IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A"
    Case Else '�����סԺ
        strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] Union ALL " & _
                   " Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A where A.����ID=[1] )  A"
    End Select
    
    strSql = _
    "   Select Decode(�����־,1,'����',4,'����',Decode(Nvl(A.��ҳID,0),0,'','��'||Nvl(A.��ҳID,0)||'��')) As ����," & _
    "         A.NO as ���ݺ�,Nvl(B.����,'δ֪') as ��������,Nvl(E.����,D.����) as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & _
    "       A.�վݷ�Ŀ as ��Ŀ,Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����," & _
    "       To_Char(" & IIf(blnDel, "-1*", "") & "A.���ʽ��,'999999999" & strDec & "') as ���ʽ��," & _
    "       To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
    " From " & strSql & ",���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.��������ID=B.ID And A.�շ�ϸĿID=D.ID" & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
    "       And A.����ID=[1]" & _
    " Order by ���� Desc,����ʱ�� Desc,���ݺ� Desc,A.���"
    Set rsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    
End Sub

Private Sub ReadInVoice(ByVal strNO As String)
    Dim strSql As String, rsInvoice As ADODB.Recordset
    
    If mblnPrint Then Exit Sub
    
    strSql = _
    " Select b.Id, b.���� As Ʊ�ݺ�," & vbNewLine & _
    " Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��," & vbNewLine & _
    "    To_Char(b.ʹ��ʱ��, 'MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����" & vbNewLine & _
    " From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B" & vbNewLine & _
    " Where a.�������� = 3 And a.Id = b.��ӡid And a.No = [1]" & vbNewLine & _
    " Order By ID"

    Set rsInvoice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNO)

    vsfInvoice.Redraw = False
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    If Not rsInvoice.EOF Then
        Set vsfInvoice.DataSource = rsInvoice
    End If
    Call SetInvoiceList
    vsfInvoice.Redraw = True
End Sub

Private Sub vsfMain_DblClick()
    Call frmManageBalance.ViewBalance(0)
End Sub

Private Sub vsfMain_GotFocus()
    If vsfMain.Row = 0 Then Exit Sub
    zl_VsGridGotFocus vsfMain, &HFFC0C0
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is vsfMain Then
        vsfMain.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalance Then
        vsfBalance.BackColorSel = &HC0C0C0
        vsfMain.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfBalanceInfo Then
        vsfBalanceInfo.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfDetail Then
        vsfDetail.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
        vsfInvoice.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfInvoice Then
        vsfInvoice.BackColorSel = &HC0C0C0
        vsfBalance.BackColorSel = &HE0E0E0
        vsfBalanceInfo.BackColorSel = &HE0E0E0
        vsfDetail.BackColorSel = &HE0E0E0
        vsfMain.BackColorSel = &HE0E0E0
    End If
End Sub

Private Sub vsfMain_LostFocus()
    zl_VsGridLOSTFOCUS vsfMain, , vsfMain.Cell(flexcpForeColor, vsfMain.Row, vsfMain.Col)
End Sub

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
            Call frmManageBalance.ShowPopup
        End If
    End With
End Sub

Public Property Get zlGetFeeState() As Integer
    '------------------------------------------------------------
    '���ܣ���ȡ��ǰѡ�м�¼�е��˷ѱ�־
    '���ƣ�Ƚ����
    'ʱ�䣺2014-12-11
    '���أ�0-�޼�¼,1-�շѼ�¼,2-�˷Ѽ�¼,3-�ѱ��˷ѵ��շѼ�¼
    '------------------------------------------------------------
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("����ID")) = "" Then Exit Property
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��¼״̬")) = "" Then
        zlGetFeeState = 0
    Else
        zlGetFeeState = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("��¼״̬")))
    End If
End Property

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:������
    '����:2013-09-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long, intActive As Integer
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsBill As Object, strTittle As String
    
    lngRow = vsfMain.Row
    Set vsBill = vsfMain: strTittle = GetUnitName & "���˽��ʼ�¼��Ϣ"
    mblnPrint = True
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = strTittle
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If vsBill Is Nothing Then Exit Sub
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsBill
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsBill
    If bytFunc = 1 Then
        Select Case zlPrintAsk(objPrint)
            Case 1
                zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '�ָ�
    With vsBill
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    
    mblnPrint = False
    vsfMain.Select lngRow, 1
    Exit Sub
ErrHand:
    mblnPrint = False
    vsfMain.Select lngRow, 1
    If ErrCenter = 1 Then Resume
End Sub
