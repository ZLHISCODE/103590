VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMultiDocView 
   BorderStyle     =   0  'None
   Caption         =   "���ĵ��鿴"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmMultiDocView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2175
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   3900
      _cx             =   6879
      _cy             =   3836
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMultiDocView.frx":6852
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   3960
      Picture         =   "frmMultiDocView.frx":6927
      Top             =   2025
      Width           =   240
   End
End
Attribute VB_Name = "frmMultiDocView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    ID = 0
    ͼ��
    ����
    ʱ��
End Enum
Public Event SelFileChanged(ByVal lngFileID As Long, ByVal strFileName As String)   '��ǰѡ����ļ��ı�
Public Event RequestModifyDoc(ByVal lngFileID As Long)  '����༭ĳ���ļ�

'################################################################################################################
'## ���ܣ�  ��ʾ��ʼ�����ĵ���Ϣ
'################################################################################################################
Public Sub InitData(ByRef ofrmParent As Object, ByVal Document As cEPRDocument, Optional ByVal lngID As Long = 0)
    Dim rs As New ADODB.Recordset, i As Long, j As Long
    Dim strIDs As String, strTime As String, varPar() As String
        
    On Error GoTo errHand
    strTime = Format(Document.EPRPatiRecInfo.����ʱ��, "yyyy-MM-dd HH:mm:ss")
            
    strIDs = GetFileRange(Document.EPRFileInfo.ID, Document.EPRPatiRecInfo.ID, strTime, Document.EPRPatiRecInfo.��������, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, False, Document.EPRPatiRecInfo.ҽ��id)
    
    gstrSQL = "Select /*+ rule*/ a.Id, a.�ļ�id, a.��������, a.���汾, a.������, a.���ʱ��, a.����ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ A," & LongIDsTable(strIDs, varPar) & vbNewLine & _
                "Where a.Id = b.Id" & vbNewLine & _
                "Order By a.���, a.����ʱ��"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������б�", varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
    With vfgThis
        .Clear
        .Rows = rs.RecordCount + 1
        .Cols = 4
        .FixedCols = 1
        .ColWidth(mCol.ID) = 0
        .ColWidth(mCol.ͼ��) = 270
        .ColWidth(mCol.����) = 1800

        i = 0
        .Cell(flexcpText, i, mCol.ID) = "ID"
        .Cell(flexcpText, i, mCol.����) = "��������"
        .Cell(flexcpText, i, mCol.ʱ��) = "���ʱ��"
        
        For i = 1 To .Rows - 1
            .Cell(flexcpText, i, mCol.ID) = NVL(rs("ID"), 0)
            .Cell(flexcpPicture, i, mCol.ͼ��) = imgIcon.Picture
            .Cell(flexcpText, i, mCol.����) = NVL(rs("��������")) & "(��" & NVL(rs("���汾")) & "��)"
            .Cell(flexcpText, i, mCol.ʱ��) = Format(NVL(rs("���ʱ��")), "YY-MM-DD")
            If lngID = rs("ID") Then j = i
            rs.MoveNext
        Next
        .ROW = j
    End With

    Exit Sub
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vfgThis.Move 20, 20, Me.ScaleWidth - 40, Me.ScaleHeight - 40
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set imgIcon.Picture = Nothing
End Sub

Private Sub vfgThis_DblClick()
    RaiseEvent RequestModifyDoc(vfgThis.TextMatrix(vfgThis.ROW, mCol.ID))
End Sub
