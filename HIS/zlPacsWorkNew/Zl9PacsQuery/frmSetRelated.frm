VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetRelated 
   Caption         =   "�й�����ʾ����"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmSetRelated.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdDetlete 
      Caption         =   "ɾ������"
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "��������"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   360
      Left            =   9480
      TabIndex        =   3
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   360
      Left            =   10680
      TabIndex        =   4
      Top             =   3720
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   2280
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgConnSet 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _cx             =   20558
      _cy             =   6165
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
      BackColorBkg    =   -2147483638
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   360
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
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   4920
      Picture         =   "frmSetRelated.frx":6852
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   5160
      Picture         =   "frmSetRelated.frx":6BC4
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmSetRelated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjScShowCfg  As New clsScShowCfg
Private mstrContrasCol As String
Private Const M_STR_COLNAME = "��������|��ʾͼ��|ͼ�������|�Ƿ�״̬ͼ��ʾ|��ǰ��ɫ|�б���ɫ|��ǰǰ��ɫ|��ǰ����ɫ|��ɫ������|��˸��ʱ(��)|��ʱʱ��ο���"
Private mstrCurDataChange As String
Private mblnIsEnabled As Boolean
Private mblnFontColor As Boolean    '��ǰ��ɫ�Ƿ�������
Private mblnBackColor As Boolean    '�б���ɫ�Ƿ�������
Private mblnFlickerTimeOut As Boolean   '��˸��ʱ�Ƿ�������
Private mblnEdit As Boolean

Private Enum ColTitle
    ct�������� = 0
    ct��ʾͼ�� = 1
    ctͼ������� = 2
    ct״̬ͼ = 3
    ct��ǰ��ɫ = 4
    ct�б���ɫ = 5
    ct��ǰǰ��ɫ = 6
    ct��ǰ����ɫ = 7
    ct��ɫ������ = 8
    ct��˸��ʱ = 9
    ctʱ��ο��� = 10
End Enum

Public Event IsItemSetted(ByVal lngItem As Long, ByRef lngRow As Long, ByRef strRowName As String)   '�жϸ������Ƿ����������ù�
Public Event ClearItemSet(ByVal lngItem As Long, ByVal lngRow As Long)   '��������жԸ����õ�����

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mblnEdit = False
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdDetlete_Click()
    On Error GoTo errHandle
    
    If vsgConnSet.Row < 1 Then Exit Sub
    vsgConnSet.RemoveItem vsgConnSet.Row
    If vsgConnSet.Rows < 2 Then cmdDetlete.Enabled = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNew_Click()
    On Error GoTo errHandle
    
    If vsgConnSet.Rows = 1 Then cmdDetlete.Enabled = True
    Call NewRow(vsgConnSet)
    Call DataDefault(vsgConnSet.Row)
    vsgConnSet.EditCell
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub CmdOK_Click()
    On Error GoTo errHandle
    
    mblnEdit = True
    Call InWriteValue
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle

    Call GridInit(M_STR_COLNAME, vsgConnSet)
    Call GridShow
    
    Call RefreshItem
    Call ShowRelatedSet
    Call RefreshWindowState(mblnIsEnabled)
    Call SetFontSize(gbytFontSize)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub GridShow()
    With vsgConnSet
        .SelectionMode = flexSelectionFree
'        .ColWidth(ColTitle.ct��ǰ����ɫ) = 1000
'        .ColWidth(ColTitle.ct��ǰǰ��ɫ) = 1000
'        .ColWidth(ColTitle.ctͼ�������) = 1000
'        .ColWidth(ColTitle.ct��ɫ������) = 1000
'        .ColWidth(ColTitle.ct״̬ͼ) = 1400
'        .ColWidth(ColTitle.ct��˸��ʱ) = 1200
        .ColComboList(ColTitle.ct��ǰ����ɫ) = "..."
        .ColComboList(ColTitle.ct��ǰǰ��ɫ) = "..."
        .ColComboList(ColTitle.ct�б���ɫ) = "..."
        .ColComboList(ColTitle.ct��ǰ��ɫ) = "..."
        .ColComboList(ColTitle.ct��ʾͼ��) = "..."
    End With
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)
    With vsgConnSet
        Select Case bytSize
            Case 0
                .ColWidth(ColTitle.ct��ǰ����ɫ) = 1000
                .ColWidth(ColTitle.ct��ǰǰ��ɫ) = 1000
                .ColWidth(ColTitle.ctͼ�������) = 1000
                .ColWidth(ColTitle.ct��ɫ������) = 1000
                .ColWidth(ColTitle.ct״̬ͼ) = 1400
                .ColWidth(ColTitle.ct��˸��ʱ) = 1200
            Case 1
                .ColWidth(ColTitle.ct��ǰ����ɫ) = 1350
                .ColWidth(ColTitle.ct��ǰǰ��ɫ) = 1350
                .ColWidth(ColTitle.ctͼ�������) = 1350
                .ColWidth(ColTitle.ct��ɫ������) = 1350
                .ColWidth(ColTitle.ct״̬ͼ) = 1600
                .ColWidth(ColTitle.ct��˸��ʱ) = 1450
            Case 2
                .ColWidth(ColTitle.ct��ǰ����ɫ) = 1700
                .ColWidth(ColTitle.ct��ǰǰ��ɫ) = 1700
                .ColWidth(ColTitle.ctͼ�������) = 1700
                .ColWidth(ColTitle.ct��ɫ������) = 1700
                .ColWidth(ColTitle.ct״̬ͼ) = 1800
                .ColWidth(ColTitle.ct��˸��ʱ) = 1700
        End Select
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsgConnSet.Height = Me.ScaleHeight - vsgConnSet.Top - cmdNew.Height - 120
    vsgConnSet.Width = Me.ScaleWidth - vsgConnSet.Left * 2
    
    cmdNew.Top = vsgConnSet.Top + vsgConnSet.Height + 60
    cmdDetlete.Left = cmdNew.Left + cmdNew.Width + 60
    cmdDetlete.Top = cmdNew.Top
    
    cmdCancel.Top = cmdNew.Top
    cmdCancel.Left = vsgConnSet.Width + vsgConnSet.Left - cmdCancel.Width
    
    cmdOK.Top = cmdNew.Top
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub

Private Sub vsgConnSet_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim objIconManage As frmIconManage
    Dim strValue As String
    Dim strIconName As String
    Dim strRowName As String
    Dim objIcon As Object
    Dim lngRow As Long
    
    If Row > 0 Then
        dlgMain.Color = 0
        Select Case Col
            Case ColTitle.ct�б���ɫ
                If IsSetted(0, lngRow, strRowName) Then
                    If MsgBox("��" & strRowName & "�������������б���ɫ��������������е��б���ɫ" & vbCrLf & "���ý���գ��Ƿ������", vbYesNo, Me.Caption) = vbNo Then
                        Exit Sub
                    Else
                        RaiseEvent ClearItemSet(0, lngRow)
                    End If
                End If
                
                dlgMain.Flags = cdlCCFullOpen
                dlgMain.ShowColor
                If dlgMain.Color > 0 Then
                    vsgConnSet.Cell(flexcpBackColor, Row, Col) = dlgMain.Color
                End If
            Case ColTitle.ct��ǰ��ɫ
                If IsSetted(2, lngRow, strRowName) Then
                    If MsgBox("��" & strRowName & "��������������ǰ��ɫ��������������е���ǰ��ɫ" & vbCrLf & "���ý���գ��Ƿ������", vbYesNo, Me.Caption) = vbNo Then
                        Exit Sub
                    Else
                        RaiseEvent ClearItemSet(2, lngRow)
                    End If
                End If
                
                dlgMain.Flags = cdlCCFullOpen
                dlgMain.ShowColor
                If dlgMain.Color > 0 Then
                    vsgConnSet.Cell(flexcpBackColor, Row, Col) = dlgMain.Color
                End If
            Case ColTitle.ct��ǰ����ɫ, ColTitle.ct��ǰǰ��ɫ
                dlgMain.Flags = cdlCCFullOpen
                dlgMain.ShowColor
                If dlgMain.Color > 0 Then
                    vsgConnSet.Cell(flexcpBackColor, Row, Col) = dlgMain.Color
                End If
            Case ColTitle.ct��ʾͼ��
                strIconName = vsgConnSet.Cell(flexcpData, Row, Col)
                Set objIconManage = New frmIconManage
                Set objIcon = objIconManage.ShowIconWindow(strIconName, Me, 1)
                If Not objIcon Is Nothing Then
                    vsgConnSet.Cell(flexcpPicture, Row, Col) = objIcon
                    vsgConnSet.Cell(flexcpPictureAlignment, Row, Col) = flexPicAlignCenterCenter
                End If
                
                vsgConnSet.Cell(flexcpData, Row, Col) = strIconName
                Set objIconManage = Nothing
        End Select
    End If
End Sub

Private Sub vsgConnSet_Click()
    Dim lngRow As Long
    Dim lngCol As Long

    On Error GoTo errHandle
    
    lngRow = vsgConnSet.Row
    lngCol = vsgConnSet.Col
    If mblnIsEnabled Then
        If lngRow <= 0 Then Exit Sub
        
        If lngCol = ColTitle.ct״̬ͼ Then
            If vsgConnSet.Cell(flexcpData, lngRow, lngCol) = 1 Then
                vsgConnSet.Cell(flexcpData, lngRow, lngCol) = 0
                vsgConnSet.Cell(flexcpPicture, lngRow, lngCol) = imgNoCheck.Picture
            Else
                vsgConnSet.Cell(flexcpPicture, lngRow, lngCol) = imgCheck.Picture
                vsgConnSet.Cell(flexcpData, lngRow, lngCol) = 1
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsgConnSet_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo errHandle
    
    If Not mblnIsEnabled Then Exit Sub
    If KeyAscii <> 8 Then Exit Sub
    lngRow = vsgConnSet.Row
    lngCol = vsgConnSet.Col
    
    If lngRow <= 0 Then Exit Sub
    
    Select Case lngCol
        Case ct��ǰ����ɫ, ct��ǰǰ��ɫ, ct�б���ɫ, ct��ǰ��ɫ
            vsgConnSet.Cell(flexcpBackColor, lngRow, lngCol) = 0
        Case ct��ʾͼ��
            vsgConnSet.Cell(flexcpPicture, lngRow, lngCol) = Nothing
            vsgConnSet.Cell(flexcpData, lngRow, lngCol) = ""
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsgConnSet_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strRowName As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If Row <= 0 Then Exit Sub
                    
    If Col = ColTitle.ct��˸��ʱ Then
        If InStr("0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsgConnSet_RowColChange()
    Dim lngRow As Long
    Dim lngCol As Long
    
    
    On Error GoTo errHandle
    
    lngRow = vsgConnSet.Row
    lngCol = vsgConnSet.Col
    
    vsgConnSet.Editable = flexEDKbdMouse
    If mblnIsEnabled Then
        If lngRow <= 0 Then Exit Sub
        
        If lngCol = ColTitle.ctͼ������� Or lngCol = ColTitle.ct��ɫ������ Or lngCol = ColTitle.ctʱ��ο��� Then
            vsgConnSet.ColComboList(lngCol) = mstrContrasCol
        End If
        
        If lngCol = ColTitle.ct״̬ͼ Then
            vsgConnSet.Editable = flexEDNone
        End If
    Else
        vsgConnSet.Editable = flexEDNone
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub RefreshItem()
'ˢ�´�����
    Dim arrItem() As String
    Dim arrValue() As String
    Dim i As Long
    
    If Len(mstrCurDataChange) > 0 Then
        arrValue = Split(mstrCurDataChange, ";")
    
        For i = 0 To UBound(arrValue)
            arrItem = Split(arrValue(i), "-")
            If UBound(arrItem) > 0 Then
                vsgConnSet.Rows = vsgConnSet.Rows + 1
                vsgConnSet.TextMatrix(i + 1, ColTitle.ct��������) = arrItem(1)
                Call DataDefault(i + 1)
            End If
        Next
    End If
    
    For i = 1 To mobjScShowCfg.RowRelationCount
        If InStr(1, UCase(mstrCurDataChange), "-" & UCase(mobjScShowCfg.RowRelation(i).TiggerData)) = 0 Then
            vsgConnSet.Rows = vsgConnSet.Rows + 1
            vsgConnSet.TextMatrix(vsgConnSet.Rows - 1, ColTitle.ct��������) = mobjScShowCfg.RowRelation(i).TiggerData
        End If
    Next
End Sub

Private Sub InWriteValue()
'д���������
    Dim objScRowRelation As clsScRowRelation
    Dim i As Long
    
    Set mobjScShowCfg = Nothing
    
    For i = 1 To vsgConnSet.Rows - 1
        If Len(Trim(vsgConnSet.TextMatrix(i, ColTitle.ct��������))) > 0 Then
            Set objScRowRelation = New clsScRowRelation
            With vsgConnSet
                objScRowRelation.TiggerData = .TextMatrix(i, ColTitle.ct��������)
                objScRowRelation.Icon = .Cell(flexcpData, i, ColTitle.ct��ʾͼ��)
                objScRowRelation.IconPerformCol = Trim(.TextMatrix(i, ColTitle.ctͼ�������))
                objScRowRelation.IsStateIcon = IIf(Val(.Cell(flexcpData, i, ColTitle.ct״̬ͼ)) = 1, True, False)
                objScRowRelation.RowFontColor = .Cell(flexcpBackColor, i, ColTitle.ct��ǰ��ɫ)
                objScRowRelation.RowBackColor = .Cell(flexcpBackColor, i, ColTitle.ct�б���ɫ)
                objScRowRelation.CellFontColor = .Cell(flexcpBackColor, i, ColTitle.ct��ǰǰ��ɫ)
                objScRowRelation.CellBackColor = .Cell(flexcpBackColor, i, ColTitle.ct��ǰ����ɫ)
                objScRowRelation.ColorPerformCol = Trim(.TextMatrix(i, ColTitle.ct��ɫ������))
                objScRowRelation.FlickerTimeOut = Val(.TextMatrix(i, ColTitle.ct��˸��ʱ))
                objScRowRelation.TimeOutReferCol = Trim(.TextMatrix(i, ColTitle.ctʱ��ο���))
            End With
            mobjScShowCfg.AddRowRelation objScRowRelation
        End If
    Next
End Sub

Public Function ShowScRowRelation(ObjScShowCfg As clsScShowCfg, strCurDataChange As String, strPerformCol As String, blnIsEnabled As Boolean, ByRef blnEdit As Boolean, ower As Object) As clsScShowCfg
    Set mobjScShowCfg = ObjScShowCfg
    mstrCurDataChange = strCurDataChange
    mblnIsEnabled = blnIsEnabled
    mstrContrasCol = strPerformCol
    mblnEdit = False
    Me.Show 1, ower
    
    blnEdit = mblnEdit
    Set ShowScRowRelation = mobjScShowCfg
End Function


Private Sub ShowRelatedSet()
'��ʾ��������
    Dim i As Long
    Dim j As Long
    Dim strFile As String
    
    For i = 1 To mobjScShowCfg.RowRelationCount
        With vsgConnSet
            For j = 1 To .Rows - 1
                If UCase(Trim(.TextMatrix(j, ColTitle.ct��������))) = UCase(Trim(mobjScShowCfg.RowRelation(i).TiggerData)) Then
                    .Cell(flexcpData, j, ColTitle.ct��ʾͼ��) = mobjScShowCfg.RowRelation(i).Icon
                    If Len(mobjScShowCfg.RowRelation(i).Icon) > 0 Then
                        strFile = zlBlobRead(mobjScShowCfg.RowRelation(i).Icon)
                        If Len(strFile) > 0 Then
                            If Len(Dir(strFile)) > 0 Then
                                .Cell(flexcpPicture, j, ColTitle.ct��ʾͼ��) = LoadPicture(strFile)
                                .Cell(flexcpPictureAlignment, j, ColTitle.ct��ʾͼ��) = flexPicAlignCenterCenter
                                Kill strFile
                            End If
                        End If
                    End If
                    .TextMatrix(j, ColTitle.ctͼ�������) = mobjScShowCfg.RowRelation(i).IconPerformCol
                    If NVL(mobjScShowCfg.RowRelation(i).IsStateIcon, False) Then
                        .Cell(flexcpData, j, ColTitle.ct״̬ͼ) = 1
                        .Cell(flexcpPicture, j, ColTitle.ct״̬ͼ) = imgCheck.Picture
                    Else
                        .Cell(flexcpData, j, ColTitle.ct״̬ͼ) = 0
                        .Cell(flexcpPicture, j, ColTitle.ct״̬ͼ) = imgNoCheck.Picture
                    End If
                    .Cell(flexcpPictureAlignment, j, ColTitle.ct״̬ͼ) = flexPicAlignCenterCenter
                    .Cell(flexcpBackColor, j, ColTitle.ct��ǰ��ɫ) = mobjScShowCfg.RowRelation(i).RowFontColor
                    .Cell(flexcpBackColor, j, ColTitle.ct�б���ɫ) = mobjScShowCfg.RowRelation(i).RowBackColor
                    .Cell(flexcpBackColor, j, ColTitle.ct��ǰǰ��ɫ) = mobjScShowCfg.RowRelation(i).CellFontColor
                    .Cell(flexcpBackColor, j, ColTitle.ct��ǰ����ɫ) = mobjScShowCfg.RowRelation(i).CellBackColor
                    .TextMatrix(j, ColTitle.ct��ɫ������) = mobjScShowCfg.RowRelation(i).ColorPerformCol
                    .TextMatrix(j, ColTitle.ct��˸��ʱ) = mobjScShowCfg.RowRelation(i).FlickerTimeOut
                    .TextMatrix(j, ColTitle.ctʱ��ο���) = mobjScShowCfg.RowRelation(i).TimeOutReferCol
                    Exit For
                End If
            Next
        End With
    Next
End Sub

Private Sub RefreshWindowState(blnState As Boolean)
    cmdNew.Enabled = blnState
    cmdDetlete.Enabled = False
    cmdOK.Enabled = blnState
    
    If blnState Then
        vsgConnSet.BackColor = &H80000005
        If vsgConnSet.Rows > 1 Then
            cmdDetlete.Enabled = blnState
        End If
    Else
        vsgConnSet.BackColor = &H8000000F
    End If
End Sub

Public Sub UnloadMe()
    Set mobjScShowCfg = Nothing
    Unload Me
End Sub

Private Sub DataDefault(lngRow As Long)
    Dim i As Long
    
    With vsgConnSet
        For i = 1 To 9
            .TextMatrix(lngRow, i) = ""
            If i = ColTitle.ct״̬ͼ Then
                .Cell(flexcpData, lngRow, i) = 0
                .Cell(flexcpPicture, lngRow, i) = imgNoCheck.Picture
                .Cell(flexcpPictureAlignment, lngRow, i) = flexPicAlignCenterCenter
            ElseIf i = ColTitle.ct��ʾͼ�� Then
                .Cell(flexcpData, lngRow, i) = ""
                .Cell(flexcpPicture, lngRow, i) = Nothing
            ElseIf i = ColTitle.ct��˸��ʱ Then
                .TextMatrix(lngRow, i) = 0
            End If
        Next
    End With
End Sub

'�������Ƿ����������ù�
Private Function IsSetted(ByVal lngItem As Long, ByRef lngRow As Long, ByRef strRowName As String) As Boolean
    RaiseEvent IsItemSetted(lngItem, lngRow, strRowName)
    
    If lngRow > 0 Then
        IsSetted = True
    Else
        IsSetted = False
    End If
End Function

Private Sub vsgConnSet_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strRowName As String
    Dim lngRow As Long
    
    If Col = ColTitle.ct��˸��ʱ Then
        If IsSetted(1, lngRow, strRowName) Then
            If MsgBox("��" & strRowName & "��������������˸��ʱ��������������е���˸��ʱ" & vbCrLf & "���ý���գ��Ƿ������", vbYesNo, Me.Caption) = vbNo Then
                Cancel = True
                Exit Sub
            Else
                RaiseEvent ClearItemSet(1, lngRow)
            End If
        End If
    End If
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim lngCmdHeight As Long
    Dim lngCmdWithd As Long
    
    If bytFontSize = 9 Then
        lngCmdHeight = 350
        lngCmdWithd = 1100
        Call SetColWithd(0)
    ElseIf bytFontSize = 12 Then
        lngCmdHeight = 385
        lngCmdWithd = 1300
        Call SetColWithd(1)
    ElseIf bytFontSize = 15 Then
        lngCmdHeight = 420
        lngCmdWithd = 1500
        Call SetColWithd(2)
    End If
    
    
    vsgConnSet.FontSize = bytFontSize
    
    cmdNew.FontSize = bytFontSize
    cmdNew.Height = lngCmdHeight
    cmdNew.Width = lngCmdWithd
    
    cmdOK.FontSize = bytFontSize
    cmdOK.Height = lngCmdHeight
    cmdOK.Width = lngCmdWithd
    
    cmdDetlete.FontSize = bytFontSize
    cmdDetlete.Height = lngCmdHeight
    cmdDetlete.Width = lngCmdWithd
    
    cmdCancel.FontSize = bytFontSize
    cmdCancel.Height = lngCmdHeight
    cmdCancel.Width = lngCmdWithd
    
    Call Form_Resize
End Sub
