VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRSort 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ļ�����"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   Icon            =   "frmEPRSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkShare 
      Caption         =   "��ʾ�ǹ���ҳ���ļ�"
      Height          =   270
      Left            =   3480
      TabIndex        =   4
      Top             =   3795
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   6870
      TabIndex        =   2
      Top             =   3735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   5610
      TabIndex        =   1
      Top             =   3735
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      DragIcon        =   "frmEPRSort.frx":6852
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   7950
      _cx             =   14023
      _cy             =   6403
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
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
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   45
      Picture         =   "frmEPRSort.frx":D0A4
      Top             =   3750
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "�������Ϸ����ı�����˳��"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   3795
      Width           =   2610
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'���峣��
'-----------------------------------------------------
Private Enum mCol
    ��� = 0: ID = 1
End Enum
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng���� As Long
Private mstrҳ���� As String
Private mblnOk As Boolean
Private Sub FillList()
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

    If chkShare.Value = vbUnchecked Then
        gstrSQL = "Select r.���, r.Id, r.��������, r.������ As ������, To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��," & vbNewLine & _
                    "       To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.���汾 As �汾" & vbNewLine & _
                    "From ���Ӳ�����¼ R, (Select ID From �����ļ��б� Where ҳ�� = [4]) F" & vbNewLine & _
                    "Where r.�ļ�id = f.Id And r.������Դ = 2 And r.�������� = [3] And r.����id = [1] And r.��ҳid = [2]" & vbNewLine & _
                    "Order By r.���, r.����ʱ��"
    Else
        gstrSQL = "Select r.���, r.Id, r.��������, r.������ As ������, To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��," & vbNewLine & _
                    "       To_Char(r.���ʱ��, 'yyyy-mm-dd hh24:mi') As ���ʱ��, r.���汾 As �汾" & vbNewLine & _
                    "From ���Ӳ�����¼ R" & vbNewLine & _
                    "Where r.������Դ = 2 And r.�������� = [3] And r.����id = [1] And r.��ҳid = [2]" & vbNewLine & _
                    "Order By r.���, r.����ʱ��"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID, mlng����, mstrҳ����)

    
    With Me.vfgThis
        .Clear
        .FixedCols = 0
        Set .DataSource = rsTemp
        .FixedCols = 1: .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        .ColAlignment(mCol.���) = flexAlignCenterCenter
        For lngCount = .FixedRows To .Rows - 1
            .TextMatrix(lngCount, mCol.���) = lngCount
        Next
    End With
End Sub
Public Function ShowMe(ByRef frmParent As Object, _
    Optional ByVal lng����ID As Long, _
    Optional ByVal lng��ҳID As Long, _
    Optional ByVal lng���� As Long, _
    Optional ByVal strҳ���� As String) As Boolean
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng���� = lng����
    mstrҳ���� = strҳ����
    
    Call FillList
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkShare_Click()
    Call FillList
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Err = 0: On Error GoTo LL
    Dim i As Long

    For i = 1 To vfgThis.Rows - 1
        gstrSQL = "Zl_���Ӳ�����¼_�������(" & Val(vfgThis.Cell(flexcpText, i, mCol.ID)) & "," & Val(vfgThis.Cell(flexcpText, i, mCol.���)) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    mblnOk = True
    Unload Me
    Exit Sub
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    mblnOk = False
End Sub
Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If vfgThis.ROW < vfgThis.Rows - 1 Then vfgThis.ROW = vfgThis.ROW + 1
    End If
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '����϶�
    With vfgThis
        If Button = vbLeftButton Then
'            Cancel = True
            Dim r%
            r = .ROW
            .Cell(flexcpBackColor, r, 0, r, .Cols - 1) = vbRed
            r = .DragRow(r)
            .Cell(flexcpCustomFormat, r, 0, r, .Cols - 1) = False
            Dim i As Long
            For i = 1 To vfgThis.Rows - 1
                vfgThis.Cell(flexcpText, i, mCol.���) = i
            Next
        End If
    End With
End Sub




