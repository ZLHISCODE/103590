VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ѡ��"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10470
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfPatient 
      Height          =   4185
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   7875
      _cx             =   13891
      _cy             =   7382
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
      BackColorBkg    =   -2147483643
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
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.ComboBox cboȱʡ���� 
      Height          =   300
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9245
      TabIndex        =   2
      Top             =   4875
      Width           =   1150
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7665
      TabIndex        =   1
      Top             =   4875
      Width           =   1150
   End
   Begin VB.ComboBox cboSect 
      Height          =   4140
      Left            =   45
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "cboSect"
      Top             =   480
      Width           =   2400
   End
   Begin VB.Label lblȱʡ���� 
      Caption         =   "ȱʡ��������"
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Label lblSect 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mfrmParent As Form
Private mrsPati As New ADODB.Recordset
Private mintȱʡ���� As Integer
Private mstrSort As String          '����|סԺ��|����ID|����|��Ժ
Private mblnSort As Boolean
'81407,����ѡ��������ؼ������С
Public mbytSize As Byte '���壺0-С����,1-������;С����Ϊ9����,������Ϊ12����

Private Sub cboSect_Click()
    Dim strSQL As String, i As Integer, lngColor As Long, l As Integer
    
    vsfPatient.Clear 1
    If cboSect.ListIndex = -1 Then Exit Sub
    If mrsPati.State = adStateOpen Then mrsPati.Close
    
    On Error GoTo errHandle
'    If Not gblnAllowOut Then
'        '��ǰ��Ժ����
'        strSQL = " Select A.����id, A.סԺ��, A.����,A.�Ա�,A.��ͥ��ַ, A.��ǰ���� As ��λ,'��' As ��Ժ,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������" & _
'                 " From ������Ϣ A, ������ҳ B" & _
'                 " Where A.��Ժ = 1 And A.����id = B.����id And A.��ҳID = B.��ҳid And A.ͣ��ʱ�� Is Null And B.��Ժ���� Is Null And " & _
'                 " B.��Ժ����id+0 =[1]" & _
'                 " Order by " & Split(mstrSort, "|")(mintȱʡ����) & " Desc"
'    Else
        'ס(��)Ժ����
        '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
        strSQL = "Select A.����ID,A.סԺ��,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,A.��ͥ��ַ,B.��Ժ���� as ��λ,Decode(B.��Ժ����,NULL,'��','') as ��Ժ,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) ��������" & _
                " From ������Ϣ A,������ҳ B" & _
                " Where A.ͣ��ʱ�� is NULL And Nvl(B.��ҳID,0)<>0" & _
                " And A.����ID=B.����ID And A.��ҳID=B.��ҳID And (Nvl(A.��Ժ,0) = 1 Or Exists (Select 1 From ������ҳ Where ����ID=A.����ID And Nvl(��ҳID,0)=0 And Nvl(��������,0)=0)) " & _
                " And A.��ǰ����ID=[1]" & _
                " Order by " & Split(mstrSort, "|")(mintȱʡ����) & " Desc"

    'End If
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboSect.ItemData(cboSect.ListIndex)))
    With vsfPatient
        '86344:���ϴ�,2015/7/8,msh����vsfFlexGrid
        .Rows = 2
        .Cols = 8
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "סԺ��"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "�Ա�"
        .TextMatrix(0, 4) = "��ͥ��ַ"
        .TextMatrix(0, 5) = "��λ"
        .TextMatrix(0, 6) = "��Ժ"
        .TextMatrix(0, 7) = "��������"
        For i = 0 To 7
            .FixedAlignment(i) = flexAlignCenterCenter
        Next i
        Do While Not mrsPati.EOF
            .TextMatrix(.Rows - 1, 0) = NVL(mrsPati!����ID)
            .TextMatrix(.Rows - 1, 1) = NVL(mrsPati!סԺ��)
            .TextMatrix(.Rows - 1, 2) = NVL(mrsPati!����)
            .TextMatrix(.Rows - 1, 3) = NVL(mrsPati!�Ա�)
            .TextMatrix(.Rows - 1, 4) = NVL(mrsPati!��ͥ��ַ)
            .TextMatrix(.Rows - 1, 5) = NVL(mrsPati!��λ)
            .TextMatrix(.Rows - 1, 6) = NVL(mrsPati!��Ժ)
            .TextMatrix(.Rows - 1, 7) = NVL(mrsPati!��������)
            .Rows = .Rows + 1
            mrsPati.MoveNext
        Loop
        
'        .Redraw = flexRDBuffered
        
        If mrsPati.RecordCount > 0 Then
            '�Զ�����MSHFlexGrid���ĸ��п��
            Call zlControl.MshSetColWidth(vsfPatient, Me)
            For i = 1 To .Rows - 1
                lngColor = GetPatiColor(.TextMatrix(i, 7))
'                .Row = i
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
'                For l = 0 To .Cols - 1
'                    .Col = l
'                    .CellForeColor = lngColor
'                Next
            Next
            .Rows = .Rows - 1
        Else
            .Rows = 2
            .Cols = 2
        End If

'        .RowHeight(0) = 320
'        .Row = 1: .TopRow = 1
'        .Redraw = flexRDDirect
        If .Visible And .Enabled = True Then .SetFocus
        If .Cols > 2 Then
        Select Case mintȱʡ����
            Case 0
                .Col = 5
            Case 1
                .Col = 1
            Case 2
                .Col = 0
            Case 3
                .Col = 2
            Case 4
                .Col = 6
            Case 5
                .Col = 7
        End Select
        .Sort = flexSortGenericDescending
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetColor()
    Dim i As Integer, lngColor As Long
    With vsfPatient
        For i = 1 To .Rows - 1
            lngColor = GetPatiColor(.TextMatrix(i, 7))
'                .Row = i
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
'                For l = 0 To .Cols - 1
'                    .Col = l
'                    .CellForeColor = lngColor
'                Next
        Next
    End With
End Sub

Private Sub cboSect_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = 13 Then
        For i = 1 To cboSect.ListCount
            If cboSect.Text <> "" Then
                If cboSect.List(i) Like "*" & cboSect.Text & "*" Then
                    cboSect.ListIndex = i
                    Exit For
                End If
            End If
        Next
    End If
End Sub

Private Sub cboȱʡ����_Click()
    If cboȱʡ����.Visible And cboȱʡ����.ListIndex <> -1 Then
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ѡ��������", cboȱʡ����.ListIndex
        mintȱʡ���� = cboȱʡ����.ListIndex
        Call cboSect_Click
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If vsfPatient.Rows > 1 Then
        If vsfPatient.TextMatrix(1, 0) <> "" Then
            mfrmParent.txtPatient.Text = "-" & vsfPatient.TextMatrix(vsfPatient.Row, 0)
            Unload Me
        End If
    End If
End Sub

Private Sub vsfPatient_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfPatient_DblClick()
    If vsfPatient.MouseRow > 0 Then cmdOK_Click
End Sub

Private Sub vsfPatient_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub vsfPatient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsfPatient.MouseRow = 0 Then
        vsfPatient.MousePointer = 99
    Else
        vsfPatient.MousePointer = 0
    End If
End Sub

Private Sub vsfPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long

    lngCol = vsfPatient.MouseCol

    If Button = 1 And vsfPatient.MousePointer = 99 Then
        If vsfPatient.TextMatrix(0, lngCol) = "" Then Exit Sub
        vsfPatient.Col = lngCol
        If mblnSort Then
            vsfPatient.Sort = flexSortGenericAscending
            mblnSort = False
        Else
            vsfPatient.Sort = flexSortGenericDescending
            mblnSort = True
        End If
'        mrsPati.Sort = vsfPatient.TextMatrix(0, lngCol) & IIf(vsfPatient.ColData(lngCol) = 0, "", " DESC")
'        Set vsfPatient.DataSource = mrsPati

        vsfPatient.ColData(lngCol) = (vsfPatient.ColData(lngCol) + 1) Mod 2
    End If
End Sub

Private Sub Form_Activate()
    vsfPatient.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    '���ý���ؼ������С��λ��
    Call SetFontSize(Me, mbytSize)
    If mbytSize = 1 Then
        cboSect.Height = cboSect.Height + 240
        vsfPatient.Height = vsfPatient.Height + 120
        lblȱʡ����.Width = lblȱʡ����.Width + 320
        cmdOK.Move cmdOK.Left - 2 * (1500 - cmdOK.Width), cmdOK.Top - 50, 1500, 420
        cmdCanc.Move cmdCanc.Left - (1800 - cmdCanc.Width), cmdCanc.Top - 50, 1500, 420
    End If
    cboȱʡ����.Left = lblȱʡ����.Left + lblȱʡ����.Width + 50
    
    Call InitPatiType
    
    mstrSort = "��λ|סԺ��|����ID|����|��Ժ|��������"
    For i = 0 To UBound(Split(mstrSort, "|"))
        cboȱʡ����.AddItem Split(mstrSort, "|")(i)
    Next
    mintȱʡ���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ѡ��������", 0))
    mintȱʡ���� = IIf(mintȱʡ���� < cboȱʡ����.ListCount, mintȱʡ����, 0)
    cboȱʡ����.ListIndex = mintȱʡ����
    
    cboSect.Clear
    
    On Error GoTo errHandle
    'by lesfeng 2010-03-08 �����Ż�
    strSQL = "Select B.ID,B.����,B.����" & _
        " From (Select Distinct ����ID From ��λ״����¼ " & _
        " ) A,���ű� B Where A.����ID=B.ID And (B.վ��=[1] Or B.վ�� is Null)" & _
        " Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)

    With rsTmp
        Do While Not .EOF
            cboSect.AddItem !���� & "-" & !����
            cboSect.ItemData(cboSect.NewIndex) = !ID
            If !ID = UserInfo.����ID Then cboSect.ListIndex = cboSect.NewIndex
            .MoveNext
        Loop
    End With
    vsfPatient.AllowUserResizing = flexResizeColumns
    If cboSect.ListCount > 0 And cboSect.ListIndex = -1 Then cboSect.ListIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblSect_Click()
    cboSect.SetFocus
End Sub

Private Sub vsfPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex - 1 >= 0 Then
                cboSect.ListIndex = cboSect.ListIndex - 1
                vsfPatient.Row = 1: vsfPatient.Col = 0: vsfPatient.SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyRight Then
        If cboSect.ListIndex <> -1 Then
            If cboSect.ListIndex + 1 <= cboSect.ListCount - 1 Then
                cboSect.ListIndex = cboSect.ListIndex + 1
                vsfPatient.Row = 1: vsfPatient.Col = 0: vsfPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub SetFontSize(ByVal objForm As Object, ByVal bytSize As Byte)
    '���ý���ؼ������С
    '���:
    '   objForm-�������
    '   bytSize-�����С: 0-С����,1-������;С����Ϊ9����,������Ϊ12����
    Dim objCtl As Control
    
    On Error Resume Next
    objForm.Font.Size = IIf(bytSize = 1, 12, 9)
    For Each objCtl In objForm.Controls
        '0-С����,1-������;С����Ϊ9����,������Ϊ12����
        objCtl.Font.Size = IIf(bytSize = 1, 12, 9)
    Next
End Sub
