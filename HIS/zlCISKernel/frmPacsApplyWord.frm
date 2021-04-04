VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPacsApplyWord 
   Caption         =   "Ӱ�����볣�ôʾ�"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frmPacsApplyWord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6255
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�༭(&E)"
      Height          =   360
      Left            =   3000
      TabIndex        =   6
      Top             =   5400
      Width           =   900
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfWord 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _cx             =   10610
      _cy             =   9128
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
      BackColorFixed  =   14811105
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
      ExplorerBar     =   1
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5235
      TabIndex        =   5
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   360
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   900
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   900
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   4200
      Picture         =   "frmPacsApplyWord.frx":6852
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   3960
      Picture         =   "frmPacsApplyWord.frx":6BC4
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmPacsApplyWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDept As Long    '����ID
Private mlngNo As Long      '��ԱID
Private mstrSort As String  '��Ŀ����
Private mstrWord As String  '���شʾ�
Private mblnEdit As Boolean '�Ƿ����༭
Private mblnIsEdit As Boolean
Private mstrCurWord As String '���ڴʾ��Ƿ����Ķ��ļ�¼

Private Const M_STR_TITLE = "Ӱ�����뵥���ôʾ�"

Private Enum TColName
    colID = 0
    col��� = 1
    colͼ�� = 2
    col�Ƿ�ͨ�� = 3
    col������ = 4
    col���� = 5
    col�ʾ����� = 6
End Enum

Public Function ShowPacsApplyWord(lngDept As Long, lngNo As Long, strSort As String, ower As Object) As String
    mlngDept = lngDept
    mlngNo = lngNo
    mstrSort = strSort
    mstrWord = ""
    mblnEdit = False
    mblnIsEdit = False
    
    Me.Show 1, ower
    
    ShowPacsApplyWord = mstrWord
End Function

Private Sub cmdAdd_Click()
    On Error GoTo errHandle
        
    cmdSave.Enabled = True
    
    Call AddRow(0, 0, "", "")
    
    vsfWord.Select vsfWord.Row, TColName.col�ʾ�����
    vsfWord.EditCell
    mblnEdit = True
    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    Unload Me
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdDelete_Click()
'ɾ���ʾ�ʱ�����������һ��ͨ����ѡ���ж���ɾ��������ֻɾ����ǰѡ�дʾ�
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim blnSelect As Boolean
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To vsfWord.Rows - 1
        If vsfWord.Cell(flexcpData, i, TColName.colͼ��) = 1 Then
            If Val(vsfWord.RowData(i)) < 0 Then
                MsgBox "�����������ʾ�Ĵ����ߣ��޷�ִ�иò�����", vbInformation, M_STR_TITLE
                vsfWord.Select i, 1
                vsfWord.ShowCell i, 1
                Exit Sub
            End If
            blnSelect = True
        End If
    Next
    
    If blnSelect Then
    'ɾ�������ʾ�
        If MsgBox("�Ƿ�ɾ����ѡ�ʾ䣿", vbYesNo, M_STR_TITLE) = vbYes Then
            i = 1
            While i <= vsfWord.Rows - 1
                If vsfWord.Cell(flexcpData, i, TColName.colͼ��) = 1 Then
                    strSQL = "Zl_Ӱ�����볣�ôʾ�_Delete(" & Val(GetValue(i, TColName.colID)) & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, "ɾ���ʾ�")
                    vsfWord.RemoveItem i
                    vsfWord.Refresh
                Else
                    i = i + 1
                End If
                
            Wend
        Else
            Exit Sub
        End If
    Else
    'ɾ����ǰһ���ʾ�
        If vsfWord.Row < 1 Then
            MsgBox "����ѡ����Ҫ�����Ĵʾ䡣", vbInformation, M_STR_TITLE
            Exit Sub
        End If
        
        If Val(vsfWord.RowData(vsfWord.Row)) < 0 Then
            MsgBox "�����������ʾ�Ĵ����ߣ��޷�ִ�иò�����", vbInformation, M_STR_TITLE
            Exit Sub
        End If
        
        If MsgBox("�Ƿ�ɾ���ʾ䡾" & Trim(GetValue(vsfWord.Row, TColName.col�ʾ�����)) & "����", vbYesNo, M_STR_TITLE) = vbYes Then
            strSQL = "Zl_Ӱ�����볣�ôʾ�_Delete(" & Val(GetValue(vsfWord.Row, TColName.colID)) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "ɾ���ʾ�")
            vsfWord.RemoveItem vsfWord.Row
            RefreshNum vsfWord.Row
        Else
            Exit Sub
        End If
    End If

    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    
    If vsfWord.Rows <= 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub RefreshNum(lngRow As Long)
    Dim i As Long
    Dim lngCount As Long
    
    If lngRow <= 0 Then Exit Sub
    lngCount = 0
    With vsfWord
        For i = lngRow To .Rows - 1
            .TextMatrix(i, TColName.col���) = .Row + lngCount
            lngCount = lngCount + 1
        Next
    End With
End Sub

Private Function IsCreator(lngID As Long) As Boolean
'�жϵ�ǰ������Ա�Ƿ�Ϊ�ʾ�Ĵ�����
    
    If lngID <> mlngNo Then
        IsCreator = False
    Else
        IsCreator = True
    End If
End Function

Private Sub cmdEdit_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Val(cmdEdit.Tag) = 0 Then
        cmdEdit.Tag = 1
        cmdEdit.Caption = "�˳�(&E)"
        cmdAdd.Visible = True
        cmdDelete.Visible = True
        cmdSave.Visible = True
        cmdSave.Enabled = False
        cmdSure.Visible = False
        cmdCancel.Visible = False
        vsfWord.ColHidden(TColName.col�Ƿ�ͨ��) = False
        vsfWord.ColHidden(TColName.col������) = False
                  
        If vsfWord.Rows > 1 Then
            vsfWord.Cell(flexcpSort, 1, TColName.col����, vsfWord.Rows - 1, TColName.col�ʾ�����) = flexSortStringNoCaseAscending
        End If
        
        For i = 1 To vsfWord.Rows - 1
            If vsfWord.RowHidden(i) = True Then
                vsfWord.RowHidden(i) = False
            End If
            If vsfWord.RowData(i) < 0 Then
                vsfWord.Cell(flexcpBackColor, i, TColName.col�Ƿ�ͨ��, i, TColName.col�ʾ�����) = &HC0FFFF
            End If
        Next
    Else
        If mblnEdit Then
            If MsgBox("�༭�����Ƿ񱣴棿", vbYesNo, M_STR_TITLE) = vbYes Then
                If Not SaveData Then
                    Exit Sub
                End If
            Else
                Call InitData
            End If
        End If
        mblnEdit = False
        
        cmdEdit.Tag = 0
        cmdEdit.Caption = "�༭(&E)"
        cmdAdd.Visible = False
        cmdDelete.Visible = False
        cmdSave.Visible = False
        cmdSure.Visible = True
        cmdCancel.Visible = True
        
        cmdSave.Enabled = False
        If vsfWord.Row > 0 Then
            cmdDelete.Enabled = True
        Else
            cmdDelete.Enabled = False
        End If
        
        vsfWord.ColHidden(TColName.col�Ƿ�ͨ��) = True
        vsfWord.ColHidden(TColName.col������) = True

        If vsfWord.Rows > 1 Then
            vsfWord.Cell(flexcpSort, 1, TColName.col�ʾ�����, vsfWord.Rows - 1, TColName.col�ʾ�����) = flexSortStringNoCaseAscending
        End If

        For i = 1 To vsfWord.Rows - 1
            If IsRepeted(i - 1, Trim(GetValue(i, TColName.col�ʾ�����))) Then
                vsfWord.RowHidden(i) = True
            End If
            
            If vsfWord.RowData(i) < 0 Then
                vsfWord.Cell(flexcpBackColor, i, TColName.col�Ƿ�ͨ��, i, TColName.col������) = &H80000005
            End If
        Next
    End If
    
    Call Form_Resize
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If Not SaveData Then Exit Sub
    
    If vsfWord.Row > 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
    
    cmdSave.Enabled = False
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim rsResult As ADODB.Recordset
    Dim blnTag As Boolean
    Dim lngID As Long
    Dim i As Long
    
    For i = 1 To vsfWord.Rows - 1
        '�������޸Ĵʾ�ʱ��Ҫ�жϴʾ��Ƿ��Ѵ���
        If Val(GetValue(i, TColName.colID)) = 0 Or Val(vsfWord.RowData(i)) = 1 Then
            If CheckRepeted(i) Then
                MsgBox IIF(vsfWord.Cell(flexcpData, i, TColName.col�Ƿ�ͨ��) = 0, "��������˸ôʾ䡣", "�ôʾ��ڿ���ͨ�ôʾ����Ѵ��ڡ�"), vbInformation, M_STR_TITLE
                vsfWord.Select i, TColName.col�ʾ�����
                vsfWord.EditCell
                Exit Function
            End If
        End If
        
        If Len(GetValue(i, TColName.col�ʾ�����)) > 200 Then
            MsgBox "�ʾ����ݹ��������ܳ���200���֡�", vbInformation, M_STR_TITLE
            vsfWord.Select i, TColName.col�ʾ�����
            vsfWord.EditCell
            Exit Function
        End If
    Next
    
    i = 1
    While i <= vsfWord.Rows - 1
        blnTag = False
        '����
        If Val(GetValue(i, TColName.colID)) = 0 Then
            If Len(Trim(GetValue(i, TColName.col�ʾ�����))) > 0 Then
                
                strSQL = "select Zl_Ӱ�����볣�ôʾ�_Insert([1],[2],[3],[4],[5]) as ����ֵ from dual"
                Set rsResult = zlDatabase.OpenSQLRecord(strSQL, "��������", mstrSort, Replace(Trim(GetValue(i, TColName.col�ʾ�����)), "'", "''"), vsfWord.Cell(flexcpData, i, TColName.col�Ƿ�ͨ��), mlngDept, mlngNo)
                
                If rsResult.RecordCount > 0 Then
                    vsfWord.TextMatrix(i, TColName.colID) = Val(Nvl(rsResult.Fields!����ֵ))
                    vsfWord.TextMatrix(i, TColName.col������) = UserInfo.����
                    lngID = Val(Nvl(rsResult.Fields!����ֵ))
                End If
            Else
                vsfWord.RemoveItem i
                blnTag = True
            End If
        End If
        
        '�޸�
        If Not blnTag Then
            If Val(vsfWord.RowData(i)) = 1 Then
                strSQL = "Zl_Ӱ�����볣�ôʾ�_Update(" & Val(GetValue(i, TColName.colID)) & ",'" & Replace(Trim(GetValue(i, TColName.col�ʾ�����)), "'", "''") & "'," & vsfWord.Cell(flexcpData, i, TColName.col�Ƿ�ͨ��) & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, "�޸�����")
                
                vsfWord.RowData(i) = 0
                lngID = Val(GetValue(i, TColName.colID))
            End If
            
            i = i + 1
        End If
    Wend
    
    If vsfWord.Rows > 1 Then
        vsfWord.Cell(flexcpSort, 1, TColName.col����, vsfWord.Rows - 1, TColName.col�ʾ�����) = flexSortStringNoCaseAscending
    End If
    
    If vsfWord.Rows > 1 And lngID > 0 Then
        For i = 1 To vsfWord.Rows - 1
            If lngID = Val(GetValue(i, TColName.colID)) Then
                vsfWord.Select i, 1
                vsfWord.ShowCell i, 1
            End If
        Next
    End If
    mblnEdit = False
    SaveData = True
End Function

Private Function GetValue(lngRow As Long, lngCol As Long) As String
    GetValue = vsfWord.TextMatrix(lngRow, lngCol)
End Function

Private Function CheckRepeted(lngRow As Long) As Boolean
'�༭�ʾ��ж��Ƿ�ͬ�����ظ�
    Dim i As Long
    
    CheckRepeted = False
    
    For i = 1 To vsfWord.Rows - 1
        If Trim(GetValue(lngRow, TColName.col�ʾ�����)) = Trim(GetValue(i, TColName.col�ʾ�����)) And vsfWord.Cell(flexcpData, i, TColName.col�Ƿ�ͨ��) = vsfWord.Cell(flexcpData, lngRow, TColName.col�Ƿ�ͨ��) And Len(Trim(GetValue(lngRow, TColName.col�ʾ�����))) > 0 And i <> lngRow Then
            CheckRepeted = True
            Exit Function
        End If
    Next
End Function

Private Sub cmdSure_Click()
    Dim i As Long
    
    On Error GoTo errHandle
    
'    For i = 1 To vsfWord.Rows - 1
'        If vsfWord.Cell(flexcpData, i, TColName.colͼ��) = 1 Then
'             mstrWord = mstrWord & IIF(Len(mstrWord) = 0, "", "��") & Trim(vsfWord.TextMatrix(i, TColName.col�ʾ�����))
'        End If
'    Next
    mstrWord = Trim(GetValue(vsfWord.Row, TColName.col�ʾ�����))
    
    Unload Me
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Call InitGrid
    Call InitFace
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Val(cmdEdit.Tag) = 1 And mblnIsEdit Then
        Call vsfWord_AfterEdit(vsfWord.Row, vsfWord.Col)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.ScaleHeight > 3000 And Me.ScaleWidth > 6000 Then
        vsfWord.Left = 120
        vsfWord.Top = 120
        
        vsfWord.Height = Me.ScaleHeight - cmdAdd.Height - 360
        vsfWord.Width = Me.ScaleWidth - 240
        
        cmdAdd.Left = vsfWord.Left
        cmdAdd.Top = vsfWord.Top + vsfWord.Height + 120
        
        cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 60
        cmdDelete.Top = cmdAdd.Top
        
        cmdSave.Left = cmdDelete.Left + cmdDelete.Width + 60
        cmdSave.Top = cmdAdd.Top
        
        cmdEdit.Left = IIF(Val(cmdEdit.Tag) = 1, Me.ScaleWidth - cmdCancel.Width - 120, vsfWord.Left)
        cmdEdit.Top = cmdAdd.Top
        
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
        cmdCancel.Top = cmdAdd.Top
        
        cmdSure.Left = cmdCancel.Left - cmdSure.Width - 60
        cmdSure.Top = cmdAdd.Top
        
'        vsfWord.ColWidth(TColName.col�ʾ�����) = vsfWord.Width - 340 - IIF(vsfWord.ColHidden(TColName.col������), 0, vsfWord.ColWidth(TColName.col������)) - IIF(vsfWord.ColHidden(TColName.col�Ƿ�ͨ��), 0, vsfWord.ColWidth(TColName.col�Ƿ�ͨ��))
    End If
End Sub

Private Sub InitFace()
'��ʼ��������ʾ

    Me.Caption = mstrSort & IIF(InStr(mstrSort, "��Ŀ") = 0, "��Ŀ", "") & "ѡ��"
    
    Call InitData
    
    cmdEdit.Tag = 0
    
    cmdAdd.Visible = False
    cmdDelete.Visible = False
    cmdSave.Visible = False
End Sub

Private Sub InitGrid()
    
    With vsfWord
        
        .Cols = 7
        .ColHidden(TColName.colID) = True
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(TColName.colID) = 0
        .ColWidth(TColName.col���) = 480
        .ColWidth(TColName.colͼ��) = 600
        .ColWidth(TColName.col�Ƿ�ͨ��) = 480
        .ColWidth(TColName.col����) = 0
        .ColWidth(TColName.col������) = 1000
        .RowHeightMin = 350
        .RowHeightMax = 350
        .ExtendLastCol = True
        .ScrollTrack = True
        .ColHidden(TColName.colͼ��) = True
        .ColHidden(TColName.col�Ƿ�ͨ��) = True
        .ColHidden(TColName.col���) = True
        .ColHidden(TColName.col����) = True
        .ColHidden(TColName.col������) = True
        
        .TextMatrix(0, TColName.col���) = "���"
        .TextMatrix(0, TColName.colID) = "ID"
        .TextMatrix(0, TColName.colͼ��) = "ѡ��"
        .TextMatrix(0, TColName.col�Ƿ�ͨ��) = "ͨ��"
        .TextMatrix(0, TColName.col����) = "����"
        .TextMatrix(0, TColName.col�ʾ�����) = "�ʾ�����"
        .TextMatrix(0, TColName.col������) = "������"
        
    End With
End Sub

Private Sub InitData()
'��ʼ����������
    Dim strSQL As String
    Dim rsResult As New ADODB.Recordset
    Dim blnOwer As Boolean
    Dim blnHidden As Boolean
    
    vsfWord.Rows = 1
    strSQL = "Select a.Id, a.�ʾ�����, a.�Ƿ�ͨ��, a.������Աid,b.���� as ������" & vbNewLine & _
                "From (Select Id, ��Ŀ����, �ʾ�����, �Ƿ�ͨ��, ������Աid" & vbNewLine & _
                "       From Ӱ�����볣�ôʾ�" & vbNewLine & _
                "       Where ������Աid = [1] And �Ƿ�ͨ�� = [2]" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Id, ��Ŀ����, �ʾ�����, �Ƿ�ͨ��, ������Աid From Ӱ�����볣�ôʾ� Where ����id = [3] And �Ƿ�ͨ�� = [4]) a,��Ա�� b" & vbNewLine & _
                "Where a.������Աid = b.id and a.��Ŀ���� = [5]" & vbNewLine & _
                "Order By a.�ʾ�����"

    Set rsResult = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ʾ�", mlngNo, 0, mlngDept, 1, mstrSort)
    
    While Not rsResult.EOF
        blnOwer = IsCreator(Val(Nvl(rsResult.Fields!������ԱID)))
        blnHidden = IsRepeted(vsfWord.Rows - 1, Nvl(rsResult.Fields!�ʾ�����))
        AddRow Val(Nvl(rsResult.Fields!ID)), Val(Nvl(rsResult.Fields!�Ƿ�ͨ��)), Nvl(rsResult.Fields!�ʾ�����), Nvl(rsResult.Fields!������), blnOwer, blnHidden
        rsResult.MoveNext
    Wend
    
    If vsfWord.Rows > 1 Then
        vsfWord.Select 1, 1
        vsfWord.ShowCell 1, 1
    End If
End Sub

Private Function IsRepeted(lngRow As Long, strValue As String) As Boolean
'�Ǳ༭�����ظ��ʾ��ж�
    Dim i As Long
    
    If lngRow < 1 Then Exit Function
    
    IsRepeted = False
    For i = 1 To lngRow
        If Trim(GetValue(i, TColName.col�ʾ�����)) = Trim(strValue) Then
            IsRepeted = True
            Exit Function
        End If
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandle
    
    If mblnEdit Then
        If MsgBox("�༭�����Ƿ񱣴棿", vbYesNo, M_STR_TITLE) = vbYes Then
            If Not SaveData Then
                Cancel = 1
                Exit Sub
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clea
End Sub

Private Sub vsfWord_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    mblnIsEdit = False
    
    If Col = TColName.col�ʾ����� And Row > 0 And Val(cmdEdit.Tag) = 1 Then
        '�޸Ĵʾ�ʱ���ܿ�
        If Len(Trim(GetValue(Row, TColName.col�ʾ�����))) = 0 And Val(GetValue(Row, TColName.colID)) > 0 Then
            MsgBox "�ʾ����ݲ���Ϊ�ա�", vbInformation, M_STR_TITLE
            vsfWord.TextMatrix(Row, TColName.col�ʾ�����) = mstrCurWord
            Exit Sub
        End If
        
        '�ж���Щ�ʾ���й��޸�
        If Val(GetValue(Row, TColName.colID)) > 0 Then
            vsfWord.RowData(Row) = 1
            mblnEdit = True
            cmdSave.Enabled = True
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_Click()
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If vsfWord.Row < 1 Then Exit Sub
    If Val(cmdEdit.Tag) = 1 Then
        lngRow = vsfWord.Row
        If vsfWord.Col = TColName.colͼ�� And lngRow > 0 Then
            If vsfWord.Cell(flexcpData, lngRow, TColName.colͼ��) = 0 Then
                vsfWord.Cell(flexcpData, lngRow, TColName.colͼ��) = 1
                vsfWord.Cell(flexcpPicture, lngRow, TColName.colͼ��) = imgCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.colͼ��) = flexPicAlignCenterCenter
            Else
                vsfWord.Cell(flexcpData, lngRow, TColName.colͼ��) = 0
                vsfWord.Cell(flexcpPicture, lngRow, TColName.colͼ��) = imgNoCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.colͼ��) = flexPicAlignCenterCenter
            End If
        End If
    
        If vsfWord.Col = TColName.col�Ƿ�ͨ�� And lngRow > 0 And Val(vsfWord.RowData(lngRow)) >= 0 Then
            mblnEdit = True
            vsfWord.RowData(lngRow) = 1
            cmdSave.Enabled = True
            If vsfWord.Cell(flexcpData, lngRow, TColName.col�Ƿ�ͨ��) = 0 Then
                vsfWord.Cell(flexcpData, lngRow, TColName.col�Ƿ�ͨ��) = 1
                vsfWord.TextMatrix(lngRow, TColName.col����) = 1
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col�Ƿ�ͨ��) = imgCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col�Ƿ�ͨ��) = flexPicAlignCenterCenter
            Else
                vsfWord.Cell(flexcpData, lngRow, TColName.col�Ƿ�ͨ��) = 0
                vsfWord.TextMatrix(lngRow, TColName.col����) = 0
                vsfWord.Cell(flexcpPicture, lngRow, TColName.col�Ƿ�ͨ��) = imgNoCheck.Picture
                vsfWord.Cell(flexcpPictureAlignment, lngRow, TColName.col�Ƿ�ͨ��) = flexPicAlignCenterCenter
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_DblClick()
    On Error GoTo errHandle
    
    If vsfWord.Row > 0 Then
        If Val(cmdEdit.Tag) = 0 Then
            If vsfWord.Row <= 0 Then Exit Sub
            
            If Val(cmdEdit.Tag) = 0 Then
                mstrWord = Trim(GetValue(vsfWord.Row, TColName.col�ʾ�����))
                Unload Me
            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub vsfWord_RowColChange()
    On Error GoTo errHandle
    
    If vsfWord.Row <= 0 Then Exit Sub
    If Val(cmdEdit.Tag) = 1 Then
        If Val(vsfWord.RowData(vsfWord.Row)) >= 0 And (vsfWord.Col = TColName.col�ʾ�����) Then
            vsfWord.Editable = flexEDKbdMouse
        Else
            vsfWord.Editable = flexEDNone
        End If
    Else
        vsfWord.Editable = flexEDNone
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub

Private Sub AddRow(lngID As Long, lngGeneral As Long, strWord As String, strOwer As String, Optional blnOwer As Boolean = True, Optional blnHidden As Boolean = False)
    With vsfWord
        .Rows = vsfWord.Rows + 1
        .ShowCell .Rows - 1, TColName.col�ʾ�����
        .Select .Rows - 1, TColName.col�ʾ�����
        .TextMatrix(vsfWord.Rows - 1, TColName.colID) = lngID
        .TextMatrix(vsfWord.Rows - 1, TColName.col���) = vsfWord.Rows - 1
        .Cell(flexcpPicture, vsfWord.Rows - 1, TColName.colͼ��) = imgNoCheck.Picture
        .Cell(flexcpData, vsfWord.Rows - 1, TColName.colͼ��) = 0
        .Cell(flexcpPictureAlignment, vsfWord.Rows - 1, TColName.colͼ��) = flexPicAlignCenterCenter
        
        .Cell(flexcpPicture, vsfWord.Rows - 1, TColName.col�Ƿ�ͨ��) = IIF(lngGeneral = 1, imgCheck.Picture, imgNoCheck.Picture)
        .Cell(flexcpData, vsfWord.Rows - 1, TColName.col�Ƿ�ͨ��) = lngGeneral
        
        '�Ƿ�ͨ���е�������
        .TextMatrix(vsfWord.Rows - 1, TColName.col����) = lngGeneral
        .Cell(flexcpPictureAlignment, vsfWord.Rows - 1, TColName.col�Ƿ�ͨ��) = flexPicAlignCenterCenter
        
        .TextMatrix(vsfWord.Rows - 1, TColName.col�ʾ�����) = strWord
        .Cell(flexcpAlignment, vsfWord.Rows - 1, TColName.col�ʾ�����) = flexAlignLeftCenter
        
        .TextMatrix(vsfWord.Rows - 1, TColName.col������) = strOwer
        .Cell(flexcpAlignment, vsfWord.Rows - 1, TColName.col������) = flexAlignLeftCenter
        
        If Not blnOwer Then
            .RowData(vsfWord.Rows - 1) = -1
        Else
            .RowData(vsfWord.Rows - 1) = 0
        End If
        
        If blnHidden Then
            .RowHidden(vsfWord.Rows - 1) = True
        End If
    End With
End Sub

Private Sub vsfWord_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭ǰ�Ĵʾ�
    On Error GoTo errHandle
    
    mstrCurWord = GetValue(Row, Col)

    cmdSave.Enabled = True
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, M_STR_TITLE
    err.Clear
End Sub


