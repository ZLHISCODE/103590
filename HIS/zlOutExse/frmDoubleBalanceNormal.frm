VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDoubleBalanceNormal 
   BorderStyle     =   0  'None
   Caption         =   "frmDoubleBalanceNormal"
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
      TabIndex        =   8
      Top             =   5655
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfBalanceInfo 
         Height          =   1845
         Left            =   0
         TabIndex        =   11
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
         ForeColorSel    =   -2147483634
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
      TabIndex        =   7
      Top             =   4335
      Width           =   2550
      Begin VSFlex8Ctl.VSFlexGrid vsfBalance 
         Height          =   1845
         Left            =   0
         TabIndex        =   10
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
         ForeColorSel    =   -2147483634
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
      TabIndex        =   6
      Top             =   3135
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1845
         Left            =   0
         TabIndex        =   9
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
         ForeColorSel    =   -2147483634
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
      TabIndex        =   2
      Top             =   3420
      Width           =   2745
      Begin XtremeSuiteControls.TabControl tabInfo 
         Height          =   2010
         Left            =   -705
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   3345
      Width           =   2715
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1845
         Left            =   300
         TabIndex        =   5
         Top             =   330
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
         ForeColorSel    =   -2147483634
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
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3015
      ScaleHeight     =   2640
      ScaleWidth      =   3120
      TabIndex        =   0
      Top             =   540
      Width           =   3120
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1800
         Left            =   510
         TabIndex        =   4
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
         BackColorSel    =   12632256
         ForeColorSel    =   -2147483634
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDoubleBalanceNormal.frx":0000
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
Attribute VB_Name = "frmDoubleBalanceNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmFilter As New frmDoubleBalanceFilter
Public mblnNOMoved As Boolean
Private mblnPrinting As Boolean

Public Sub MakeFilter(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String)
    Call mfrmFilter.InitFilter(Me, lngModul, strPrivs)
End Sub

Public Sub ReadData(ByVal intType As Integer, ByVal strPrivs As String, Optional ByVal lngPatiID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ղ�������¼
    '����:������
    '���:intType-��ȡ��¼�ķ�ʽ��0Ϊʹ�ù���������ȡ��1Ϊʹ��IDKIND������ȡ
    '����:2014-9-11
    '��ע:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMain As ADODB.Recordset, strTable As String
    Dim strFilter As String, strInvoice As String, strSQLtmp As String
    Dim DatBegin As Date, DatEnd As Date, blnMoved As Boolean
    Dim str��Ʊ�Ѵ�ӡ As String
    
    On Error GoTo ErrHand
    str��Ʊ�Ѵ�ӡ = ",Nvl((Select 1" & vbNewLine & _
            "            From Ʊ�ݴ�ӡ���� M, Ʊ��ʹ����ϸ N" & vbNewLine & _
            "            Where m.Id = n.��ӡid And n.���� = 1 And n.ԭ�� = 6 And m.No = a.no And Rownum < 2), 0) As ��Ʊ�Ѵ�ӡ" & vbNewLine
    If intType = 0 Then
        If mfrmFilter.mblnInit = True Then
            With mfrmFilter
                DatBegin = .dtpBegin.Value
                DatEnd = .dtpEnd.Value
                blnMoved = zlDatabase.DateMoved(IIf(DatBegin < DatEnd, DatBegin, DatEnd))
                strFilter = ""
                strFilter = strFilter & IIf(.txt����.Text = "", "", " And C.����=[3] ")
                strFilter = strFilter & IIf(.cbo����Ա.Text = "�����շ�Ա", "", " And A.����Ա����=[4] ")
                strFilter = strFilter & IIf(.txt�����.Text = "", "", " And C.�����=[5] ")
                strFilter = strFilter & IIf(.txtסԺ��.Text = "", "", " And C.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [6]) ")
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
                    strInvoice = "Select A.NO" & _
                    " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
                    " Where A.��������=1 And A.ID=B.��ӡID And B.Ʊ��=1 And B.����=1" & _
                    " And B.���� " & strSQLtmp
                End If
                If strInvoice <> "" Then strFilter = strFilter & " And A.NO In (" & strInvoice & ") "
                'strFilter = strFilter & IIf(.chkDelRecord, " And A.��¼״̬ <> 0 ", " ")
                '������Դ
                If .opt����(0).Value Then '����
                    strFilter = strFilter & " And  b.�����־ in (1,4)"
                ElseIf .opt����(1).Value Then 'סԺ
                    strFilter = strFilter & " And  b.�����־ =2"
                Else '�����סԺ
                End If
                If blnMoved Then
                    strTable = zlGetFullFieldsTable("���ò����¼", 2, "", True, "A", True)
                Else
                    strTable = "���ò����¼ A"
                End If
                strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������, Max(A.��¼״̬) As �˷ѱ�־,A.����ID,A.ʵ��Ʊ��" & str��Ʊ�Ѵ�ӡ & _
                         " From " & strTable & ", ������ü�¼ B, ������Ϣ C " & _
                         " Where A.�Ǽ�ʱ�� Between [1] And [2] And A.����ID=C.����ID And A.�շѽ���ID=B.����ID And Nvl(A.����״̬,0)=0 " & _
                         "      And A.��¼״̬ In (1,3) And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2) " & strFilter & _
                         " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������,A.����ID,A.ʵ��Ʊ��"
                         
                If .chkDelRecord.Value = 1 Then
                    strSQL = strSQL & " Union " & _
                        "   Select NO, ���ӱ�־, ����, �Ա�, ����, -1 * Sum(ʵ�ս��) As ������, ����Ա����, �Ǽ�ʱ��, �������, 2 As �˷ѱ�־, ����ID, ʵ��Ʊ��" & str��Ʊ�Ѵ�ӡ & _
                        "   From (Select Distinct a.No, Decode(Nvl(a.���ӱ�־, 0), 1, '�Һ�', '�շ�') As ���ӱ�־, b.����, b.�Ա�, b.����, b.ʵ�ս��," & _
                        "                a.����Ա����, a.�Ǽ�ʱ��, a.�������, b.No As ���ݺ�, b.����id, a.����id, a.ʵ��Ʊ��" & _
                        "          From  " & strTable & ", ������ü�¼ B, ����Ԥ����¼ D, ������Ϣ C" & _
                        "          Where a.�Ǽ�ʱ�� Between [1] And [2] And a.������� = d.������� And b.����id = d.����id And a.����id = c.����id" & _
                        "                And Nvl(a.����״̬, 0) = 0 And a.��¼״̬ = 2 " & strFilter & ") A" & _
                        "   Group By NO, ���ӱ�־, ����, �Ա�, ����, ����Ա����, �Ǽ�ʱ��, �������,����ID,ʵ��Ʊ��"
                End If
                strSQL = strSQL & " Order By �Ǽ�ʱ�� Desc"
                Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, DatBegin, DatEnd, _
                                                    .txt����.Text, zlStr.NeedName(.cbo����Ա.Text), .txt�����.Text, .txtסԺ��.Text, _
                                                    .txtNOBegin.Text, .txtNoEnd.Text, .txtFactBegin.Text, .txtFactEnd.Text)
                Set vsfMain.DataSource = rsMain
                Call SetMain
            End With
        Else
            strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������, Max(A.��¼״̬) As �˷ѱ�־, A.����ID, A.ʵ��Ʊ��" & str��Ʊ�Ѵ�ӡ & _
                     " From ���ò����¼ A, ������ü�¼ B " & _
                     " Where Trunc(A.�Ǽ�ʱ��)=Trunc(Sysdate) And A.�շѽ���ID=B.����ID And A.��¼״̬ In (1,3) And Nvl(A.����״̬,0)=0 " & _
                     "      And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2) " & _
                     " And A.����Ա����=[1] " & _
                     " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������,A.����ID,A.ʵ��Ʊ��" & _
                     " Order By �Ǽ�ʱ�� Desc"
            Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
            Set vsfMain.DataSource = rsMain
            Call SetMain
        End If
    End If
    If intType = 1 Then
        'ʹ��IDKIND������ȡ
        strSQL = " Select A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, Sum(B.���ʽ��), A.����Ա����, A.�Ǽ�ʱ��, A.�������, Max(A.��¼״̬) As �˷ѱ�־,A.����ID,A.ʵ��Ʊ��" & str��Ʊ�Ѵ�ӡ & _
                 " From ���ò����¼ A, ������ü�¼ B " & _
                 " Where A.����ID= [1] And A.�շѽ���ID=B.����ID And A.��¼״̬ In (1,3) And Nvl(A.����״̬,0)=0 " & _
                 "      And Not Exists (Select 1 From ���ò����¼ Where �������=A.������� And ��¼״̬=2) " & _
                 IIf(InStr(strPrivs, "���в���Ա") > 0, "", " And A.����Ա����=[2] ") & _
                 " Group By A.No, Decode(Nvl(A.���ӱ�־,0),1,'�Һ�','�շ�'), B.����, B.�Ա�, B.����, A.����Ա����, A.�Ǽ�ʱ��, A.�������,A.����ID,A.ʵ��Ʊ��"
        strSQL = strSQL & " Union " & _
                        "   Select NO, ���ӱ�־, ����, �Ա�, ����, -1 * Sum(ʵ�ս��) As ������, ����Ա����, �Ǽ�ʱ��, �������, 2 As �˷ѱ�־, ����ID, ʵ��Ʊ��" & str��Ʊ�Ѵ�ӡ & _
                        "   From (Select Distinct a.No, Decode(Nvl(a.���ӱ�־, 0), 1, '�Һ�', '�շ�') As ���ӱ�־, b.����, b.�Ա�, b.����, b.ʵ�ս��," & _
                        "                a.����Ա����, a.�Ǽ�ʱ��, a.�������, b.No As ���ݺ�, b.����id, a.����id,a.ʵ��Ʊ��" & _
                        "          From ���ò����¼ A, ������ü�¼ B, ����Ԥ����¼ D, ������Ϣ C" & _
                        "          Where A.����ID= [1] And a.������� = d.������� And b.����id = d.����id And a.����id = c.����id" & _
                        "                And Nvl(a.����״̬, 0) = 0 And a.��¼״̬ = 2" & _
                        IIf(InStr(strPrivs, "���в���Ա") > 0, "", " And A.����Ա����=[2]") & ") A" & _
                        "   Group By NO, ���ӱ�־, ����, �Ա�, ����, ����Ա����, �Ǽ�ʱ��, �������, ����ID, ʵ��Ʊ��" & _
                        " Order By �Ǽ�ʱ�� Desc"
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, UserInfo.����)
        Set vsfMain.DataSource = rsMain
        Call SetMain
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
        .Redraw = flexRDNone
        
        If .Rows = 1 Then .Rows = 2
        
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            If Not Visible Then
                .ColWidth(i) = Split(varData(i), ",")(2)
                If .ColWidth(i) = 0 Then .ColHidden(i) = True
            End If
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsfInvoice, App.ProductName & "\" & Me.Name)
        
        .HighLight = flexHighlightWithFocus
        .RowHeight(-1) = 300: .RowHeight(0) = 350

        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetMain()
    Dim i As Long
    With vsfMain
        zl_vsGrid_Para_Restore 1124, vsfMain, Me.Caption, "������Ϣ�б�", True
        .RowHeight(0) = 350
        If .Rows = 1 Then .Rows = 2
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("�˷ѱ�־"))) = 2 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
            ElseIf Val(.TextMatrix(i, .ColIndex("�˷ѱ�־"))) = 3 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
            Else
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
            End If
            .TextMatrix(i, .ColIndex("������")) = Format(.TextMatrix(i, .ColIndex("������")), gstrDec)
            .TextMatrix(i, .ColIndex("����ʱ��")) = Format(.TextMatrix(i, .ColIndex("����ʱ��")), "yyyy-mm-dd hh:mm:ss")
            .RowHeight(i) = 300
        Next i
        If .Rows >= 2 Then .Select 1, 1
    End With
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "���ݺ�,1,0|���,1,0|��������,1,0|������,1,0|�ѱ�,1,0|���,4,800|����,1,2000|��Ʒ��,1,2000|" & _
            "���,1,1200|��λ,4,500|����,7,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1000|" & _
            "����,4,1000|˵��,1,1800|��¼״̬,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        'Call RestoreFlexState(vsfDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
        If .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then Call DetailSplitGroup
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                .RowHeight(i) = 300
            End If
        Next i
        
        If gTy_System_Para.bytҩƷ������ʾ = 0 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = True
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 1 Then
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 2 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
    End With
End Sub

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Է����б���Ϣ���з�����ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsfDetail
        For i = 0 To .COLS - 1
            If i < .ColIndex("���") And i > .ColIndex("˵��") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("���")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���ݺ�"))
                 strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                 strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                 For j = 0 To .COLS - 1
                    If j < .ColIndex("Ӧ�ս��") Then
                        If j >= .ColIndex("���") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    ElseIf .ColIndex("ʵ�ս��") = j Then
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("Ӧ�ս��") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrDec)
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrDec)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), gstrDec)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("���"))
        Call .AutoSize(.ColIndex("����"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("Ӧ�ս��") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    
    strHead = "���㷽ʽ,4,1000|���,7,1000|�������,4,1000|ժҪ,1,1200|����,1,1000|������ˮ��,1,1000|����˵��,1,1200|����,1,0"
    
    With vsfBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) Like "*���*" Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                strTemp = Val(.TextMatrix(i, .ColIndex("���")))
                If InStr(strTemp, ".") = 0 Then
                    strAcc = "0.00"
                Else
                    strTemp = Split(strTemp, ".")(1)
                    strAcc = "0."
                    If Len(strTemp) < 2 Then
                        strAcc = "0.00"
                    Else
                        For j = 1 To Len(strTemp)
                            strAcc = strAcc & "0"
                        Next j
                    End If
                End If
                .TextMatrix(i, .ColIndex("���")) = Format(.TextMatrix(i, .ColIndex("���")), strAcc)
            Else
                If .TextMatrix(i, .ColIndex("���")) <> "" Then .TextMatrix(i, .ColIndex("���")) = Format(.TextMatrix(i, .ColIndex("���")), "0.00")
            End If
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsfBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        
        .Redraw = True
    End With
End Sub

Private Sub SetExtendInfo()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|���㷽ʽ,1,0|����,1,0|���,1,0|��Ŀ,1,1200|����,1,2000|������ˮ��,1,0"
    
    With vsfBalanceInfo
        .Redraw = flexRDNone
        .HighLight = flexHighlightWithFocus
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        .RowHeight(0) = 350
        
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
                 
                 For j = 0 To .COLS - 1
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
        For j = 0 To .COLS - 1
            .MergeCol(j) = True
        Next
        .Redraw = flexRDBuffered
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
    zl_vsGrid_Para_Save 1124, vsfMain, Me.Caption, "������Ϣ�б�", True
End Sub

Private Sub PicDetail_Resize()
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
End Sub

Private Sub vsfBalance_GotFocus()
    SetActiveList vsfBalance
End Sub

Private Sub vsfBalanceInfo_GotFocus()
    SetActiveList vsfBalanceInfo
End Sub

Private Sub vsfDetail_GotFocus()
    SetActiveList vsfDetail
End Sub

Private Sub vsfInvoice_GotFocus()
    SetActiveList vsfInvoice
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������")) = "" Then Exit Sub
    Call ReadDetail(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))), _
                    vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("����")) = "�Һ�")
    Call ReadInVoice(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))))
    Call ReadBalance(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))))
    Call ReadBalanceInfo(Val(vsfMain.TextMatrix(NewRow, vsfMain.ColIndex("�������"))))
End Sub

Private Sub ReadBalance(ByVal lngBalanceID As Long)
    Dim strSQL As String, rsBalance As ADODB.Recordset
    
    strSQL = _
    "Select Decode(Mod(a.��¼����, 10), 1, '��Ԥ���', Nvl(a.���㷽ʽ, 'δ����')) As ���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��," & vbNewLine & _
    "       Decode(Mod(Max(a.��¼����), 10), 1, '', Max(a.�������)) As �������," & vbNewLine & _
    "       Decode(Mod(Max(a.��¼����), 10), 1, '', Max(a.ժҪ)) As ժҪ," & vbNewLine & _
    "       Decode(Mod(Max(a.��¼����), 10), 1, '', Max(a.����)) As ����," & vbNewLine & _
    "       Decode(Mod(Max(a.��¼����), 10), 1, '', Max(a.������ˮ��)) As ������ˮ��," & vbNewLine & _
    "       Decode(Mod(Max(a.��¼����), 10), 1, '', Max(a.����˵��)) As ����˵��" & vbNewLine & _
    "From ����Ԥ����¼ A" & vbNewLine & _
    "Where a.������� = [1]" & vbNewLine & _
    "Group By Decode(Mod(a.��¼����, 10), 1, '��Ԥ���', Nvl(a.���㷽ʽ, 'δ����'))"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    
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
    Dim strSQL As String, rsInfo As ADODB.Recordset
    
    strSQL = _
        "Select b.����id || '_' || b.ԭԤ��id As ID, a.���㷽ʽ, Max(c.����) As ����, Sum(Nvl(-1 * f.���, a.��Ԥ��)) As ���, b.������Ŀ, b.��������," & vbNewLine & _
        "       Max(Nvl(f.������ˮ��, a.������ˮ��)) As ������ˮ��" & vbNewLine & _
        "From ����Ԥ����¼ A, �������㽻�� B, ҽ�ƿ���� C, ����Ԥ����¼ E, �����˿���Ϣ F" & vbNewLine & _
        "Where a.Id = b.����id And a.�����id = c.Id(+) And a.������� = [1]" & vbNewLine & _
        "      And b.ԭԤ��id = e.Id(+) And e.����id = f.��¼id(+) And f.����id(+) = [1]" & vbNewLine & _
        "Group By b.����id, b.ԭԤ��id, a.���㷽ʽ, b.������Ŀ, b.��������" & vbNewLine & _
        "Order By ID"
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = Replace(strSQL, "�������㽻��", "H�������㽻��")
        strSQL = Replace(strSQL, "�����˿���Ϣ", "H�����˿���Ϣ")
    End If
    Set rsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    
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

Private Sub ReadDetail(ByVal lngBalanceID As Long, ByVal bln�ҺŲ��� As Boolean)
    Dim strSQL As String, rsDetail As ADODB.Recordset
    Dim blnDel As Boolean
    mblnNOMoved = zlDatabase.NOMoved("���ò����¼", vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("���㵥��")))
    blnDel = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("�˷ѱ�־"))) = 2
    If blnDel Then
        strSQL = _
                " Select NO As ���ݺ�, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, " & _
                "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, Max(����) As ����, Max(˵��),Max(״̬), Min(�˷�״̬)" & vbNewLine & _
                " From (Select a.����ID,D1.���� as ��������,A.������,a.No,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                        IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
                "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                        IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
                "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                        IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
                "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����,Max(Decode(A.��¼״̬,2,'��'||ABS(A.ִ��״̬)||'���˷�',Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��'))) As ˵��," & _
                "       Max(A.��¼״̬) As ״̬,Min(A.��¼״̬) As �˷�״̬, Nvl(a.�۸񸸺�, a.���) As ���" & _
                " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҩƷ��� X," & _
                "       (Select Distinct ����ID From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ Where �������= [1]) F" & _
                " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
                "       And Mod(A.��¼����,10)=[2] And A.����ID = F.����ID " & _
                "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And A.��������ID=D1.ID(+) And E1.����(+)=1 And E1.����(+)=3" & _
                " Group by a.����id, D1.����, a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
                "       Nvl(A.��������,B.��������),X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1) )" & _
                " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, ����, ִ�п��� Having Sum(����) <> 0" & _
                " Order By ���ݺ�, ���"
    Else
        strSQL = _
                " Select NO As ���ݺ�, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, " & _
                "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, Max(����) As ����, Max(˵��),Max(״̬), Min(�˷�״̬)" & vbNewLine & _
                " From (Select a.����ID,D1.���� as ��������,A.������,a.No,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                        IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
                "       To_Char(Avg(Nvl(A.����,1)*A.����)" & _
                        IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
                "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                        IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
                "       To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
                "       To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
                "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����,Max(Decode(A.��¼״̬,2,'��'||ABS(A.ִ��״̬)||'���˷�',Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��'))) As ˵��," & _
                "       Max(A.��¼״̬) As ״̬,Min(A.��¼״̬) As �˷�״̬, Nvl(a.�۸񸸺�, a.���) As ���" & _
                " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҩƷ��� X," & _
                "       (Select Distinct �շѽ���ID As ����ID From " & IIf(mblnNOMoved, "H", "") & "���ò����¼ Where �������= [1]) F" & _
                " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
                "       And Mod(A.��¼����,10)=[2] And A.����ID = F.����ID " & _
                "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And A.��������ID=D1.ID(+) And E1.����(+)=1 And E1.����(+)=3" & _
                " Group by a.����id, D1.����, a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
                "       Nvl(A.��������,B.��������),X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1) )" & _
                " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, ����, ִ�п��� Having Sum(����) <> 0" & _
                " Order By ���ݺ�, ���"
    End If
    Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID, IIf(bln�ҺŲ���, 4, 1))
    vsfDetail.Redraw = False
    vsfDetail.Clear 1
    vsfDetail.Rows = 2
    If Not rsDetail.EOF Then
        Set vsfDetail.DataSource = rsDetail
    End If
    Call SetDetail
    vsfDetail.Redraw = True
End Sub

Private Sub ReadInVoice(ByVal lngBalanceID As Long)
    Dim strSQL As String, rsInvoice As ADODB.Recordset
    
    strSQL = _
    " Select Distinct B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ����') as ʹ��ԭ��," & _
    " To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
    " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B," & _
            "(Select Distinct NO From ���ò����¼ Where �������= [1]) C" & _
    " Where A.��������=1 And A.ID=B.��ӡID" & _
    " And B.Ʊ��=1 And A.NO=C.NO" & _
    " Order by ID"
    
    Set rsInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    Set vsfInvoice.DataSource = rsInvoice
    Call SetInvoiceList
End Sub

Private Sub vsfMain_DblClick()
    Call frmReplenishTheBalanceManage.ViewBalance(0)
End Sub

Private Sub vsfMain_GotFocus()
    SetActiveList vsfMain
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

Private Sub vsfMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRow As Integer
    With vsfMain
        'If .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
        If Button = 2 Then
            If Y <= 300 Then
                Exit Sub
            End If
'            intRow = Y \ 300
'            If intRow <= .Rows - 1 Then
'                If .Enabled And .Visible Then .SetFocus
'                .Select intRow, 0
'            End If
            Call frmReplenishTheBalanceManage.ShowPopup
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
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("�������")) = "" Then Exit Property
    If vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("�˷ѱ�־")) = "" Then
        zlGetFeeState = 0
    Else
        zlGetFeeState = Val(vsfMain.TextMatrix(vsfMain.Row, vsfMain.ColIndex("�˷ѱ�־")))
    End If
End Property

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim i As Long, lngCurrentRow As Long
    Dim objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    
    With vsfMain
        If .Rows = 1 Then Exit Sub
        If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("�������"))) = 0 Then Exit Sub
    End With
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = "���ղ���������������¼�嵥"
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsfMain
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .COLS - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Then
                .ColWidth(i) = 0
            End If
        Next
    End With

    Err = 0: On Error GoTo ErrHand:
    mblnPrinting = True
    lngCurrentRow = vsfMain.Row
    Set objPrint.Body = vsfMain
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
    With vsfMain
        For i = 0 To .COLS - 1
            If .ColHidden(i) = True Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    vsfMain.Row = lngCurrentRow
    mblnPrinting = False
    Exit Sub
ErrHand:
    mblnPrinting = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
