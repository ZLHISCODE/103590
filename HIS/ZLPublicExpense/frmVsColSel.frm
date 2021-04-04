VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmVsColSel 
   BorderStyle     =   0  'None
   Caption         =   "������"
   ClientHeight    =   3252
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   2772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   3210
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2700
      _cx             =   4762
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmVsColSel.frx":0000
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.CommandButton cmdClose 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2415
         TabIndex        =   1
         Top             =   30
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVsColSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type WinLocate
    Left As Double
    Top As Double
    lngTxtH As Long
End Type
Private mWindowPosition As WinLocate           '����λ��
Private mVsGrid As VSFlexGrid
Private Const MFRM_MIN_WIDTH = 2775
Private Const MFRM_MIN_HEIGHT = 3255

Public Function ShowColSet(ByVal frmMain As Form, ByVal strTittle As String, vsGrid As VSFlexGrid, _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngTxtHeight As Long = 0) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ýӿ�
    '����:
    '����:�����óɹ�,����true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Err = 0: On Error Resume Next
    Set mVsGrid = vsGrid
    With mWindowPosition
        .Left = WinLeft
        .Top = WinTop
        .lngTxtH = lngTxtHeight
    End With
    Call LoadFulltoColSel
    Call ReSetWindowsFormLocal
    With Me
        .Show 1, frmMain
    End With
End Function

Public Sub ReSetWindowsFormLocal()
    '����:�������ô��ڵĴ�С��λ��
    Dim dblColsWidth As Double, dblMinRowheight As Double, lngScrW As Long
    Dim lngTaskHeight As Long
    Dim dblRowsHeight As Double
    Dim dblTemp As Double
    Dim i As Long
    '��λ
    With mWindowPosition
        Me.Left = .Left + 15
        Me.Top = .Top
    End With
    
    dblColsWidth = 0
    For i = 0 To vsColSet.Cols - 1
        If Not vsColSet.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsColSet.ColWidth(i) + 15
        End If
    Next
    dblMinRowheight = vsColSet.RowHeightMin
    lngTaskHeight = GetTaskbarHeight
    dblColsWidth = dblColsWidth + 300
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
    dblRowsHeight = dblMinRowheight * vsColSet.Rows + 30
    
    dblColsWidth = IIf(dblColsWidth < MFRM_MIN_WIDTH, MFRM_MIN_WIDTH, dblColsWidth)
    
    If Me.Top + dblRowsHeight <= Screen.Height Then
        '���嶥��+���и߶�+С�ڵ�����Ļ�߶ȡ�
        '���Ƿ����С�߶Ȼ�С,�����С,������С�ȸ�Ϊ׼
        If dblRowsHeight < MFRM_MIN_HEIGHT Then
            Me.Height = MFRM_MIN_HEIGHT
        Else
            Me.Height = dblRowsHeight
        End If
    Else
        '���嶥��+�������߶�+���ŵ��ܸ߶ȴ�����Ļ�߶�,��Ҫ��һ�¼��
        '1.���ϰ���Ļ�߶��Ƿ���°����߶�Ҫ�ߣ���������ϰ����ĸ߶�Ϊ׼���������°���Ϊ׼.
        If Screen.Height - Me.Top > Me.Top - mWindowPosition.lngTxtH - 15 Then
            '�°���Ҫ��
            Me.Height = Screen.Height - Me.Top - lngTaskHeight
            '������ȫװ��,ֻ�ܸ���������������б�������б�ĸ߶�
         Else
            dblTemp = Me.Top - mWindowPosition.lngTxtH - 15
            Me.Top = Me.Top - mWindowPosition.lngTxtH - 15
            '�ϰ���Ҫ��
            If dblTemp - dblRowsHeight > 0 Then
                '�ϰ�������ȫ��װ��
                Me.Height = dblRowsHeight
                If Me.Height < MFRM_MIN_HEIGHT Then Me.Height = MFRM_MIN_HEIGHT
            Else
                Me.Height = dblTemp
            End If
            Me.Top = Me.Top - Me.Height
        End If
    End If
    
    '�����ȶ�λ
    '����п�����С�ڵ��ڵ�ǰ����Ŀ��,����������Ϊ׼
    If dblColsWidth + Me.Left < Screen.Width Then
        '���еĿ����ȫ����ʾ
        Me.Width = dblColsWidth
    Else
        '����Ƿ������Ļ�����ұ���Ļ��
        If Screen.Width - Me.Left >= Me.Left Then
            '�ұ���Ļ��
            Me.Width = Screen.Width - Me.Left
        Else
            Me.Left = Me.Left
            '�����Ļ��
            If dblColsWidth < Me.Left Then
                Me.Width = dblColsWidth
            Else
                Me.Width = Me.Left
            End If
            Me.Left = Me.Left - Me.Width
        End If
    End If
 
End Sub

Private Function LoadFulltoColSel() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long, arrSplit As Variant
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    

    vsColSet.Clear 1
    vsColSet.Rows = 2
    With mVsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            arrSplit = Split(.ColData(i) & "||", "||")
            
            If Trim(.ColKey(i)) <> "" And (Val(arrSplit(0)) = 1 Or Val(arrSplit(0)) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("����")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("ѡ��")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = Val(arrSplit(0))
                If Val(arrSplit(0)) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    sngFrmHeight = Me.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = mVsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        .Top = mVsGrid.Top + mVsGrid.RowHeight(0) + 15
        sngFrmHeight = sngFrmHeight - .Top
        If sngFrmHeight > sngSelSumHeight Then
            .Height = sngSelSumHeight
        Else
            .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
        End If
        .SetFocus
    End With
End Function
Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, lngRow As Long
    With mVsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
End Function

Private Sub cmdClose_Click()
    Form_KeyDown vbKeyEscape, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
 
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsColSet
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight
        .Width = ScaleWidth
        cmdClose.Left = .Left + .Width - cmdClose.Width - 10
    End With
    
End Sub

Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�޸ĺ�
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("ѡ��"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("����")), blnShow, IIf(.Tag = "Head", False, True))
        Case Else
        End Select
    End With
End Sub

Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            'rowdata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If Val(.RowData(Row)) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
'Private Sub vsColSet_LostFocus()
'    vsColSet.Visible = False
'End Sub


