VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCooperateRegArrangePlan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicUnit 
      BorderStyle     =   0  'None
      Height          =   6765
      Left            =   0
      ScaleHeight     =   6765
      ScaleWidth      =   2580
      TabIndex        =   3
      Top             =   240
      Width           =   2580
      Begin VB.ListBox lstUnits 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         ItemData        =   "frmCooperateRegArrangePlan.frx":0000
         Left            =   0
         List            =   "frmCooperateRegArrangePlan.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblUnitTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox picUnitReg 
      BorderStyle     =   0  'None
      Height          =   7485
      Left            =   3720
      ScaleHeight     =   7485
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   0
      Width           =   5580
      Begin VB.CheckBox chkDisable 
         Caption         =   "��������λ���øúű�"
         Height          =   330
         Left            =   2460
         TabIndex        =   6
         Top             =   -15
         Width           =   2265
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUnits 
         Height          =   4455
         Left            =   -240
         TabIndex        =   1
         Top             =   1920
         Width           =   7335
         _cx             =   12938
         _cy             =   7858
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCooperateRegArrangePlan.frx":0004
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   110
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
         Editable        =   2
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblUnitRegTitle 
         Caption         =   "***:��ŷ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmCooperateRegArrangePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlngPriItem As Long
Private mlng�ƻ�Id              As Long
Private mrs�޺�                 As ADODB.Recordset
Private mrs�ƻ�                 As ADODB.Recordset
Private mstr�Ű�                As String '����|ȫ��||��һ|����||��������
Private mblnUnload As Boolean
Private mblnʱ��                As Boolean '�������������ʱ�����ϸ���ʱ��������
Private mrsʱ���               As ADODB.Recordset
Private mstrKey      As String
Private mrsSource    As ADODB.Recordset
Private mrsUnitsReg  As ADODB.Recordset
Private mrsUnitsInfo As ADODB.Recordset
Private mrsUnits As ADODB.Recordset
Private mrsDisable As ADODB.Recordset
Private mblnChange   As Boolean
Private mbln��ſ��� As Boolean
Public Event frmUnload(ByVal blnCancel As Boolean)
Private Sub cmdCancel_Click()
    RaiseEvent frmUnload(True)
End Sub

Private Sub Form_Resize()
   Err.Number = 0
     On Error Resume Next
     If mblnʱ�� Then
        With Me.PicUnit
            .Left = Me.ScaleLeft
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
        End With
        
        With Me.picUnitReg
            .Left = PicUnit.Left + PicUnit.Width + 1 * Screen.TwipsPerPixelX
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
            .Width = Me.ScaleWidth - .Left
        End With
     Else
        PicUnit.Visible = False
        With Me.picUnitReg
            .Left = Me.ScaleLeft
            .Top = Me.ScaleTop
            .Height = Me.ScaleHeight
            .Width = Me.ScaleWidth
        End With
     End If
End Sub
 
Public Function frmInit(ByVal lng�ƻ�ID As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���óɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-29 14:16:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mlng�ƻ�Id = lng�ƻ�ID
    If InitData() = False Then Exit Function
    mblnʱ�� = chkExistsʱ��(lng�ƻ�ID)
    Call InitRs
    Call InitPage
    If InitUntils() = False Then Exit Function
    If Not mblnʱ�� Then LoadUnitsReg
   ' Call InitPlan
    frmInit = True
End Function

Private Sub lstUnits_Click()
    Static strUnits As String
    If lstUnits.Text = strUnits Then Exit Sub
    If mblnChange Then
        MoveUnitReg strUnits
    End If
    strUnits = lstUnits.Text
    lblUnitRegTitle.Caption = strUnits & ":������λԤԼ����"
    LoadUnitsReg
    
End Sub

Private Sub LoadUnitsReg()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ѿ��Ѿ������������Ϣ
    '����:2013-10-29 18:14:15
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varCol As Variant, i As Long, j        As Long
    Dim lng�������� As Long, lngRow As Long
    Dim strUnit As String, str���� As String
    
    If mblnʱ�� Then
        If Not mrsDisable Is Nothing Then
            With mrsDisable
                .Filter = "������λ='" & lstUnits.Text & "'"
                If .RecordCount = 0 Then
                    chkDisable.Value = 0
                Else
                    chkDisable.Value = 1
                End If
            End With
        End If
          With vsUnits
            .Clear 1
            .Rows = 3
            varCol = Split(mstr�Ű�, "||")
            For i = 0 To UBound(varCol)
                mrsSource.Filter = "������Ŀ='" & Split(varCol(i), "|")(0) & "'"
                 If mrsSource.RecordCount > 0 Then
                    lngRow = 2
                    Do While Not mrsSource.EOF
                        If lngRow + 1 >= .Rows Then .Rows = .Rows + 1
                        .TextMatrix(lngRow, i * 3 + 0) = Nvl(mrsSource!ʱ���)
                        .TextMatrix(lngRow, i * 3 + 1) = Val(Nvl(mrsSource!����))
                        mrsUnitsReg.Filter = "������Ŀ='" & Split(varCol(i), "|")(0) & "' and ������λ='" & lstUnits.Text & "' and ���=" & Val(Nvl(mrsSource!���))
                        If mrsUnitsReg.RecordCount > 0 Then
                            lng�������� = Val(Nvl(mrsUnitsReg!����))
                        Else
                            lng�������� = 0
                        End If
                        .TextMatrix(lngRow, i * 3 + 2) = lng��������
                        lngRow = lngRow + 1
                        mrsSource.MoveNext
                    Loop
                    
                End If
            Next
            mrsUnitsReg.Filter = 0
            mrsSource.Filter = 0
         End With
         Exit Sub
    End If
    varCol = Split(mstr�Ű�, "||")
    With vsUnits
        For i = 2 To .Rows - 1
            strUnit = .TextMatrix(i, 0)
             If strUnit <> "" Then
                For j = 0 To UBound(varCol)
                   str���� = Split(varCol(j), "|")(0)
                   mrsUnitsReg.Filter = "������λ='" & strUnit & "' and ������Ŀ='" & str���� & "'"
                   If mrsUnitsReg.RecordCount > 0 Then
                        .TextMatrix(i, j + 1) = Val(Nvl(mrsUnitsReg!����))
                   End If
                Next
             End If
        Next
    End With
    mrsUnitsReg.Filter = 0
End Sub

Private Sub PicUnit_Resize()
    On Error Resume Next
    lblUnitRegTitle.Move 0, 0, PicUnit.ScaleWidth, lblUnitRegTitle.Height
    Me.lstUnits.Move 0, lblUnitRegTitle.Height, PicUnit.ScaleWidth, PicUnit.ScaleHeight - lblUnitRegTitle.Height
End Sub
Private Sub MoveUnitReg(Optional ByVal str������λ As String)
    '�Ժ�����λ�ҺŽ������·���
    Dim str������Ŀ  As String, str���� As String
    Dim lngԭ�������� As Long, lng�������� As Long
    Dim lngԭ�������� As Long, lng��������  As Long
    Dim blnʱ�� As Boolean, varCol  As Variant
    Dim strʱ��� As String, j As Long, i As Long
    Dim str�Һź�����λ     As String
    
    If Not mblnChange Then Exit Sub
    If mblnʱ�� = False Then Exit Sub
    
    mblnChange = False
    
    On Error GoTo errHandle
    
    varCol = Split(mstr�Ű�, "||")
    If str������λ = "" Then
        str�Һź�����λ = lstUnits.Text
    Else
        str�Һź�����λ = str������λ
    End If
     
    For j = 2 To vsUnits.Rows - 1
        For i = 0 To UBound(varCol)
         
            strʱ��� = vsUnits.TextMatrix(j, i * 3 + 0)
            If Trim(vsUnits.TextMatrix(j, i * 3 + 1)) <> "" Then
            lng�������� = Val(vsUnits.TextMatrix(j, i * 3 + 1))
            lng�������� = Val(vsUnits.TextMatrix(j, i * 3 + 2))
            If Trim(strʱ���) = "" Then
                blnʱ�� = False
            Else
                blnʱ�� = True
            End If
            str���� = Split(varCol(i), "|")(0)
            'If str���� = "��һ" Then Stop
            mrsSource.Filter = "������Ŀ='" & str���� & "'" & IIf(blnʱ��, " And ʱ���='" & Trim(strʱ���) & "'", "")
            mrsUnitsReg.Filter = "������Ŀ='" & str���� & "' And ������λ='" & str�Һź�����λ & "'" & IIf(blnʱ��, " And ʱ���='" & Trim(strʱ���) & "'", "")
            
            If mrsSource.RecordCount > 0 Then
                lngԭ�������� = Val(Nvl(mrsSource!����))
            Else
                lngԭ�������� = 0
            End If
            If mrsUnitsReg.RecordCount > 0 Then
                lngԭ�������� = Val(Nvl(mrsUnitsReg!����))
            Else
                lngԭ�������� = 0
            End If
            
            lng�������� = lngԭ�������� + lngԭ�������� - lng��������
            
            If mrsSource.RecordCount > 0 Then
                mrsSource!���� = lng��������
                mrsSource.Update
            Else
                With mrsSource
                    .AddNew
                    !�ƻ�Id = mlng�ƻ�Id
                    !������Ŀ = str����
                    !��� = 0
                    !���� = lng��������
                    !ʱ��� = strʱ���
                    .Update
                End With
            End If
            
            If mrsUnitsReg.RecordCount > 0 Then
                With mrsUnitsReg
                    !���� = lng��������
                End With
            Else
                
            With mrsUnitsReg
                .AddNew
                !������λ = str�Һź�����λ
                !�ƻ�Id = mlng�ƻ�Id
                !������Ŀ = str����
                !��� = IIf(mrsSource.RecordCount > 0, mrsSource!���, 0)
                !���� = lng��������
                !ʱ��� = strʱ���
                .Update
            End With
 
            End If
            End If
            mrsUnitsReg.Filter = 0
            mrsSource.Filter = 0
        Next
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'------------------------------------------------------------------------
'ҳ����ù����뷽��
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL As String
    Dim lng�ƻ�ID       As Long
    Dim i       As Long
    Dim strTemp As String

    If mlng�ƻ�Id = -1 Then Exit Function
    lng�ƻ�ID = mlng�ƻ�Id

    On Error GoTo Hd

   strSQL = " " & _
        "   Select a.Id as �ƻ�ID,a.�ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,NVL(A.Ĭ��ʱ�μ��,5) as Ĭ��ʱ�μ��, " & _
        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
        "   From ( " & vbNewLine & _
        "       Select B.ID,a.id As �ƻ�id, B.����, A.����, B.����id, A.��Ŀid, B.ҽ������, B.ҽ��id, A.����, A.��һ, A.�ܶ�, A.����," & _
        "              A.����, A.����, A.����, B.��������, A.���﷽ʽ, A.��ſ���, A.��Чʱ�� As ��ʼʱ��, A.ʧЧʱ�� As ��ֹʱ��,A.Ĭ��ʱ�μ��  As Ĭ��ʱ�μ�� " & _
        "        From �ҺŰ��� B, �ҺŰ��żƻ� A " & _
        "       Where A.����ID = B.ID And A.Id=[1] " & _
        ") A,�շ���ĿĿ¼ B,���ű� D " & _
        "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
        "        "
    Set mrs�ƻ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
         
    If mrs�ƻ�.EOF Then
        ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
        Exit Function
    End If
    
    mbln��ſ��� = IIf(Val(Nvl(mrs�ƻ�!��ſ���)) = 1, True, False)
    mstr�Ű� = ""

    For i = 0 To 6
        strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")

        If Nvl(mrs�ƻ�("��" & strTemp)) <> "" Then
            If mstr�Ű� <> "" Then mstr�Ű� = mstr�Ű� & "||"
            mstr�Ű� = mstr�Ű� & "��" & strTemp & "|" & Nvl(mrs�ƻ�("��" & strTemp))
        End If

    Next
        
    strSQL = "" & _
    "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
    "               ��������,�Ƿ�ԤԼ" & _
    "   From  �Һżƻ�ʱ�� " & "   Where �ƻ�ID=[1]" & "   Order by ����,ʱ��,���"
    Set mrsʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID)

    If Not mrsʱ���.EOF Then mblnʱ�� = True
    '�ҺŰ�������
    strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �Һżƻ����� where �ƻ�ID=[1]  Order BY ������Ŀ      "
    Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
    
    InitData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function


Private Function InitPage() As Boolean
    Dim i As Long, j As Long
    Dim varCol As Variant, lng�������� As Long
    Dim varData As Variant
    On Error GoTo errHandle
    
    If mstr�Ű� = "" Then Exit Function
    chkDisable.Visible = False
    varCol = Split(mstr�Ű�, "||")
    If mblnʱ�� Then
        chkDisable.Visible = True
        With vsUnits
                .Clear 1
                .ColWidthMin = 1000
                .Cols = (UBound(varCol) + 1) * 3
                For i = 0 To UBound(varCol)
                    For j = 0 To 2
                          .TextMatrix(0, i * 3 + j) = Split(varCol(i), "|")(0) & "(" & Split(varCol(i), "|")(1) & ")"
                    Next
                    .TextMatrix(1, i * 3 + 0) = "ʱ���"
                    .TextMatrix(1, i * 3 + 1) = "ʣ������"
                    .TextMatrix(1, i * 3 + 2) = "��������"
                Next
            mrs�޺�.Filter = 0
            For i = 0 To .Cols - 1
                .FixedAlignment(i) = flexAlignCenterCenter
            Next
            .MergeRow(0) = True
            .MergeRow(1) = True
            .AllowUserResizing = flexResizeColumns
            .Editable = flexEDKbdMouse
        End With
        InitPage = True
        Exit Function
    End If
    With vsUnits
        .Cols = UBound(varCol) + 2
        .TextMatrix(0, 0) = "������λ"
        .TextMatrix(1, 0) = "������λ"
        .ColWidth(0) = 2000
        For i = 0 To UBound(varCol)
            varData = Split(varCol(i) & "|", "|")
            lng�������� = 0
             mrs�޺�.Filter = "������Ŀ='" & Split(varCol(i), "|")(0) & "'"
             If mrs�޺�.RecordCount > 0 Then lng�������� = IIf(Val(Nvl(mrs�޺�!��Լ��)) = 0, Val(Nvl(mrs�޺�!�޺���)), Val(Nvl(mrs�޺�!��Լ��)))
            .TextMatrix(0, i + 1) = varData(0)
            .TextMatrix(1, i + 1) = IIf(varData(1) = "", "��", varData(1)) & "(" & lng�������� & ")"
            .Cell(flexcpData, 1, i + 1) = lng��������
        Next
        mrs�޺�.Filter = 0
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "����"
        .ColDataType(.Cols - 1) = flexDTBoolean
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .MergeCol(i) = False
        Next
        .ExtendLastCol = False
        .AllowUserResizing = flexResizeColumns
        .MergeRow(0) = True
        '.MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCellsFixed = flexMergeFixedOnly
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Editable = flexEDKbdMouse
        zl_vsGrid_Para_Restore mlngModule, vsUnits, Me.Caption, "��������_��ʱ���"
    End With
    InitPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chkDisable_Click()
    With chkDisable
        If .Value = 1 Then
            vsUnits.Enabled = False
        Else
            vsUnits.Enabled = True
        End If
        With mrsDisable
            .Filter = "������λ='" & lstUnits.Text & "'"
            If .RecordCount <> 0 Then
                .MoveFirst
                .Delete adAffectCurrent
                .Update
            End If
            If chkDisable.Value = 1 Then
                .AddNew
                !������λ = lstUnits.Text
                .Update
            End If
        End With
    End With
End Sub

Private Function InitRs()
    Dim i         As Long
    Dim j         As Long
    Dim strList() As String
    Dim lng�޺���   As Long
    Dim lng��Լ��   As Long
    Dim rsTmp  As ADODB.Recordset
    Dim strSQL As String
    Dim str������Ŀ As String
    Dim strʱ��� As String
    Dim lng�������� As Long
    Dim lng�������� As Long
    
    On Error GoTo errHandle
    
    '��ʼ�� ���ݼ�
    With mrsUnitsReg
        Set mrsUnitsReg = New ADODB.Recordset
        mrsUnitsReg.Fields.Append "������λ", adVarChar, 40
        mrsUnitsReg.Fields.Append "�ƻ�ID", adBigInt
        mrsUnitsReg.Fields.Append "������Ŀ", adVarChar, 10
        mrsUnitsReg.Fields.Append "���", adBigInt, 18
        mrsUnitsReg.Fields.Append "����", adBigInt, 18
        mrsUnitsReg.Fields.Append "ʱ���", adVarChar, 60
        mrsUnitsReg.CursorLocation = adUseClient
        mrsUnitsReg.LockType = adLockOptimistic
        mrsUnitsReg.CursorType = adOpenStatic
        mrsUnitsReg.Open
    End With

    With mrsSource
        Set mrsSource = New ADODB.Recordset
        mrsSource.Fields.Append "�ƻ�ID", adBigInt
        mrsSource.Fields.Append "������Ŀ", adVarChar, 10
        mrsSource.Fields.Append "���", adBigInt, 18
        mrsSource.Fields.Append "����", adBigInt, 18
        mrsSource.Fields.Append "ʱ���", adVarChar, 60
        mrsSource.CursorLocation = adUseClient
        mrsSource.LockType = adLockOptimistic
        mrsSource.CursorType = adOpenStatic
        mrsSource.Open
    End With
    
    With mrsDisable
        Set mrsDisable = New ADODB.Recordset
        mrsDisable.Fields.Append "������λ", adVarChar, 50
        mrsDisable.CursorLocation = adUseClient
        mrsDisable.LockType = adLockOptimistic
        mrsDisable.CursorType = adOpenStatic
        mrsDisable.Open
    End With
    
    If mstr�Ű� = "" Then Exit Function
    strList = Split(mstr�Ű�, "||")
    If mblnʱ�� Then
         '����Ƿ�ʱ��
         
        For i = 0 To UBound(strList)
            mrsʱ���.Filter = "����='" & Split(strList(i), "|")(0) & "' and �Ƿ�ԤԼ=1"
            If mrsʱ���.RecordCount = 0 Then mrsʱ���.Filter = "����='" & Split(strList(i), "|")(0) & "'"
            
            If mrsʱ���.RecordCount = 0 Then
               '���û������ʱ��� ����дʱ���
               mrs�޺�.Filter = "������Ŀ='" & Split(strList(i), "|")(0) & "'"

               If mrs�޺�.RecordCount = 0 Then
                   mrs�޺�.Filter = 0
               Else
                   lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
                   lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
                   
                   If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
                    With mrsSource
                        .AddNew
                        !�ƻ�Id = mlng�ƻ�Id
                        !������Ŀ = Split(strList(i), "|")(0)
                        !��� = 0
                        !���� = lng��Լ��
                        .Update
                    End With
               End If 'mrs�޺�.recourdcount
               
            Else    'mrsʱ���.recordCount=0
                Do While Not mrsʱ���.EOF
                    With mrsSource
                        .AddNew
                        !�ƻ�Id = mlng�ƻ�Id
                        !������Ŀ = Split(strList(i), "|")(0)
                        !��� = Val(Nvl(mrsʱ���!���))
                        !���� = Val(Nvl(mrsʱ���!��������))
                        !ʱ��� = mrsʱ���!ʱ�䷶Χ
                        .Update
                    End With
                    mrsʱ���.MoveNext
                Loop
            End If
        Next
        mrsʱ���.Filter = 0
    Else
    
        For i = 0 To UBound(strList)
           '���û������ʱ��� ����дʱ���
            mrs�޺�.Filter = "������Ŀ='" & Split(strList(i), "|")(0) & "'"
    
            If mrs�޺�.RecordCount = 0 Then
                mrs�޺�.Filter = 0
            Else
                lng�޺��� = Val(Nvl(mrs�޺�!�޺���))
                lng��Լ�� = Val(Nvl(mrs�޺�!��Լ��))
                
                If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
                '���س�ʼ������
                With mrsSource
                    .AddNew
                    !�ƻ�Id = mlng�ƻ�Id
                    !������Ŀ = Split(strList(i), "|")(0)
                    !��� = 0
                    !���� = lng��Լ��
                    .Update
                End With
                
            End If 'mrs�޺�.recourdcount
        Next
    End If
    
    '�Ѿ��������
    strSQL = "Select ������λ, �ƻ�ID, ������Ŀ, ���, ���� From ������λ�ƻ�����  Where �ƻ�ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)

    If rsTmp.RecordCount > 0 Then

        Do While Not rsTmp.EOF
            mrsSource.Filter = "������Ŀ='" & rsTmp!������Ŀ & "' and ���=" & rsTmp!���

            With mrsUnitsReg
                .AddNew
                !������λ = Nvl(rsTmp!������λ)
                !�ƻ�Id = mlng�ƻ�Id
                !������Ŀ = Nvl(rsTmp!������Ŀ)
                !��� = Val(Nvl(rsTmp!���))
                !���� = Val(Nvl(rsTmp!����))

                If mrsSource.RecordCount > 0 Then
                    !ʱ��� = mrsSource!ʱ���
                End If
                
                .Update
            End With

            mrsSource.Filter = 0
            rsTmp.MoveNext
        Loop
    End If
     
    strSQL = "Select Distinct ������λ From ������λ�ƻ�����  Where �ƻ�ID=[1] And ���� = 0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
    
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            With mrsDisable
                .AddNew
                !������λ = Nvl(rsTmp!������λ)
                .Update
            End With
            rsTmp.MoveNext
        Loop
    End If
    
     Do While Not mrsSource.EOF
        str������Ŀ = mrsSource!������Ŀ
        strʱ��� = mrsSource!ʱ���
        lng�������� = Val(Nvl(mrsSource!����))
        lng�������� = 0
        mrsUnitsReg.Filter = "������Ŀ='" & str������Ŀ & "'" & IIf(Trim(strʱ���) <> "", " And ʱ���='" & strʱ��� & "'", "")
        Do While Not mrsUnitsReg.EOF
            lng�������� = Val(Nvl(mrsUnitsReg!����)) + lng��������
            mrsUnitsReg.MoveNext
        Loop
        If lng�������� <> 0 Then
           mrsSource!���� = lng�������� - lng��������
           mrsSource.Update
        End If
        mrsSource.MoveNext
     Loop
     mrsUnitsReg.Filter = 0
     If mrsUnitsReg.RecordCount > 0 Then mrsUnitsReg.MoveFirst
     mrsSource.Filter = 0
     If mrsSource.RecordCount <> 0 Then mrsSource.MoveFirst
    InitRs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function chkExistsʱ��(ByVal lng�ƻ�ID As Long) As Boolean

    '���ð����Ƿ����ʱ��
    Dim strSQL    As String
    Dim rsTmp     As ADODB.Recordset
    Dim blnExists As Boolean
    On Error GoTo Hd
    strSQL = "Select �ƻ�ID From �Һżƻ�ʱ�� A Where �ƻ�ID=[1] And RowNum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ƻ�ID)
    chkExistsʱ�� = rsTmp.RecordCount > 0
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function InitUntils() As Boolean
    Dim strSQL As String
    Dim i As Long, j        As Long
    Dim lngRow  As Long
    Dim varCol  As Variant
    vsUnits.Clear 1
    vsUnits.ColWidthMin = 1000
    If mstr�Ű� = "" Then Exit Function
    
    On Error GoTo Hd
    lstUnits.Clear
    
    strSQL = "Select ����, ����, ����, ȱʡ��־ From �Һź�����λ Order By ȱʡ��־ Desc"
    Set mrsUnits = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnits.EOF Then Exit Function
    
    If mblnʱ�� Then
        Do While Not mrsUnits.EOF
            lstUnits.AddItem Nvl(mrsUnits!����)
            mrsUnits.MoveNext
        Loop
        If lstUnits.ListCount > 0 Then lstUnits.Selected(0) = True
        InitUntils = True
        Exit Function
    End If
    With vsUnits
        .Clear 1
        .ColWidthMin = 1000
        If mstr�Ű� = "" Then Exit Function
        .Rows = 2 + mrsUnits.RecordCount
        varCol = Split(mstr�Ű�, "||")
        lngRow = 2
        Do While Not mrsUnits.EOF
            .TextMatrix(lngRow, 0) = mrsUnits!����
            lngRow = lngRow + 1
            mrsUnits.MoveNext
        Loop
        For i = 2 To .Rows - 1
            mrsDisable.Filter = "������λ='" & .TextMatrix(i, 0) & "'"
            If mrsDisable.RecordCount <> 0 Then
                .Cell(flexcpChecked, i, .Cols - 1, i, .Cols - 1) = 1
            End If
        Next i
    End With
    InitUntils = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub picUnitReg_Resize()
    On Error Resume Next
    If mblnʱ�� Then
        lblUnitRegTitle.Move picUnitReg.ScaleLeft, picUnitReg.ScaleTop, picUnitReg.ScaleWidth, lblUnitRegTitle.Height
        chkDisable.Left = picUnitReg.ScaleWidth - chkDisable.Width
        With vsUnits
            .Left = Screen.TwipsPerPixelX * 2
            .Top = lblUnitRegTitle.Top + lblUnitRegTitle.Height + Screen.TwipsPerPixelY * 4
            .Width = picUnitReg.ScaleWidth
            .Height = Me.picUnitReg.ScaleHeight - lblUnitRegTitle.Height - lblUnitRegTitle.Top - Screen.TwipsPerPixelY * 2 - 40 * Screen.TwipsPerPixelY
        End With
   Else
        lblUnitRegTitle.Visible = False
        With vsUnits
            .Left = picUnitReg.ScaleLeft
            .Top = picUnitReg.ScaleTop
            .Width = picUnitReg.ScaleWidth
            .Height = Me.picUnitReg.ScaleHeight
        End With
   End If
  
End Sub

Private Sub vsUnits_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vsUnits.ColIndex("����") <> Col Then vsUnits.TextMatrix(Row, Col) = Val(vsUnits.TextMatrix(Row, Col))
     mblnChange = True
End Sub

Public Function SaveData() As Boolean
    Dim i As Long, strSQL As String
    Dim strTmp  As String, cllPro As New Collection
    Dim str������λ As String, strPre������λ As String, j As Long
    Dim varCol As Variant
    Dim strDisable As String
    If mblnChange Then
        Call MoveUnitReg
    End If
    
    If mblnʱ�� Then
        For i = 0 To lstUnits.ListCount - 1
            mrsUnitsReg.Filter = "������λ='" & lstUnits.List(i) & "'"
            strSQL = "Zl_������λ�ƻ�����_Delete(" & mlng�ƻ�Id & ",'" & lstUnits.List(i) & "')"
            zlAddArray cllPro, strSQL
            strDisable = ""
            mrsDisable.Filter = "������λ='" & lstUnits.List(i) & "'"
            If mrsDisable.RecordCount <> 0 Then
                For j = 1 To vsUnits.Cols - 1
                    If InStr(strDisable, Mid(vsUnits.TextMatrix(0, j), 1, InStr(vsUnits.TextMatrix(0, j), "(") - 1)) = 0 Then
                        If strDisable <> "" Then strDisable = strDisable & "|"
                        strDisable = strDisable & Mid(vsUnits.TextMatrix(0, j), 1, InStr(vsUnits.TextMatrix(0, j), "(") - 1)
                    End If
                Next j
            End If
            If mrsUnitsReg.RecordCount > 0 Then
                With mrsUnitsReg
                    strTmp = ""
                    mrsUnitsReg.Filter = "������λ='" & lstUnits.List(i) & "' And ����>0"
                    Do While Not mrsUnitsReg.EOF
                        If strTmp <> "" Then strTmp = strTmp & "|"
                        strTmp = strTmp & !������Ŀ & "," & !��� & "," & !����
                        mrsUnitsReg.MoveNext
                    Loop
                    If strTmp <> "" And strDisable = "" Then
                        strSQL = "Zl_������λ�ƻ�����_Insert(" & mlng�ƻ�Id & ",'" & lstUnits.List(i) & "','" & strTmp & "')"
                        zlAddArray cllPro, strSQL
                    End If
                End With
                mrsUnitsReg.Filter = 0
            End If
            If strDisable <> "" Then
                strSQL = "Zl_������λ�ƻ�����_Insert(" & mlng�ƻ�Id & ",'" & lstUnits.List(i) & "',Null,Null,'" & strDisable & "')"
                zlAddArray cllPro, strSQL
            End If
        Next
    End If
    If Not mblnʱ�� Then
        With vsUnits
            For i = 2 To .Rows - 1
               str������λ = Trim(.TextMatrix(i, .ColIndex("������λ")))
               If str������λ <> "" Then
                    strSQL = "Zl_������λ�ƻ�����_Delete(" & mlng�ƻ�Id & ",'" & str������λ & "')"
                    zlAddArray cllPro, strSQL
                    'Zl_������λ���ſ���_Insert
                    '    ����id_In   ������λ���ſ���.����id%Type,
                    '    ������λ_In ������λ���ſ���.������λ%Type,
                    '    ���ſ���_In Varchar2
                    '    --���ſ���_in ������Ŀ,���1,����|������Ŀ,���2,����|������Ŀ,���3,����|��������
                    If .Cell(flexcpChecked, i, .Cols - 1, i, .Cols - 1) = 1 Then
                        strTmp = ""
                        For j = 1 To .Cols - 2
                            strTmp = strTmp & "|" & .TextMatrix(0, j)
                        Next
                        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                        strSQL = "Zl_������λ�ƻ�����_Insert(" & mlng�ƻ�Id & ",'" & str������λ & "',Null,Null,'" & strTmp & "')"
                        zlAddArray cllPro, strSQL
                    Else
                        strTmp = ""
                        For j = 1 To .Cols - 1
                            If Val(.TextMatrix(i, j)) <> 0 Then
                                strTmp = strTmp & "|" & .TextMatrix(0, j) & "," & 0 & "," & Val(.TextMatrix(i, j))
                            End If
                        Next
                        If strTmp <> "" Then
                            strTmp = Mid(strTmp, 2)
                            strSQL = "Zl_������λ�ƻ�����_Insert(" & mlng�ƻ�Id & ",'" & str������λ & "','" & strTmp & "')"
                            zlAddArray cllPro, strSQL
                        End If
                    End If
               End If
            Next
        End With
    End If
     
    Err = 0: On Error GoTo Errhand:
    mrsDisable.Filter = 0
    strDisable = ""
    Do While Not mrsDisable.EOF
        strDisable = strDisable & "|" & mrsDisable!������λ
        mrsDisable.MoveNext
    Loop
    If strDisable <> "" Then strDisable = Mid(strDisable, 2)
    zlDatabase.SetPara "���ú�����λ", strDisable, glngSys, 1110
     zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub vsUnits_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUnits.Rows - 1 <= 1 Then Exit Sub
    Call zl_VsGridRowChange(vsUnits, IIf(OldRow = 1, 2, OldRow), NewRow, OldCol, NewCol)
End Sub

Private Sub vsUnits_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
        If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
             Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsUnits_Validate(Cancel As Boolean)
   If mblnChange Then
     MoveUnitReg
   End If
End Sub

Private Sub vsUnits_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 Dim lng�������� As Long, lng�������� As Long
    Dim lng����   As Long, str������Ŀ As String
    Dim str��λ As String, blnʱ�� As Boolean
    Dim strʱ��� As String, strKey As String
    If Not mblnʱ�� Then
        With vsUnits
            If .ColIndex("����") <> Col Then
                strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
                If .Row < 2 Then Exit Sub
                If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
                
                strKey = Format(Abs(Val(strKey)), "####;;;")
                If Val(strKey) > Val(.Cell(flexcpData, 1, Col)) Then
                    MsgBox "�������ܴ����޺���(" & Val(.Cell(flexcpData, 1, Col)) & ")", vbOKOnly + vbDefaultButton2 + vbInformation, gstrSysName
                    Call vsUnits_GotFocus
                    Cancel = True: Exit Sub
                End If
                .EditText = strKey
            Else
                With mrsDisable
                    .Filter = "������λ='" & vsUnits.TextMatrix(Row, 0) & "'"
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        .Delete adAffectCurrent
                        .Update
                    End If
                    If vsUnits.Cell(flexcpChecked, Row, Col, Row, Col) = 2 Then
                        .AddNew
                        !������λ = vsUnits.TextMatrix(Row, 0)
                        .Update
                    End If
                End With
            End If
        End With
        Exit Sub
    End If
    
    If Col Mod 3 <> 2 Then Exit Sub
     strʱ��� = vsUnits.TextMatrix(Row, Col - 2)
     If Trim(strʱ���) = "" Then
         blnʱ�� = False
     Else
         blnʱ�� = True
     End If
     str��λ = lstUnits.Text
     mrsSource.Filter = "������Ŀ='" & str������Ŀ & "'" & IIf(blnʱ��, " And ʱ���='" & Trim(strʱ���) & "'", "")
     lng�������� = Val(vsUnits.EditText)
     str������Ŀ = Mid(vsUnits.TextMatrix(0, Col), 1, InStr(vsUnits.TextMatrix(0, Col), "(") - 1)
     mrsSource.Filter = "������Ŀ='" & str������Ŀ & "'" & IIf(blnʱ��, " And ʱ���='" & Trim(strʱ���) & "'", "")
     mrsUnitsReg.Filter = "������Ŀ='" & str������Ŀ & "' And ������λ='" & str��λ & "'" & IIf(blnʱ��, " And ʱ���='" & Trim(strʱ���) & "'", "")
     
     If mrsSource.RecordCount = 0 Then
         lng���� = 0
     Else
         lng���� = Val(Nvl(mrsSource!����))
     End If
     If mrsSource.RecordCount = 0 Then mrsSource.Filter = 0: Cancel = True: Exit Sub
     lng���� = Val(vsUnits.TextMatrix(Row, Col)) + lng����
     If lng�������� > lng���� Then Cancel = True: Exit Sub
     lng���� = lng���� - lng��������
     If mrsSource.RecordCount = 0 Then
         With mrsSource
             .AddNew
             !�ƻ�Id = mlng�ƻ�Id
             !������Ŀ = str������Ŀ
             !��� = 0
             !���� = lng����
             !ʱ��� = ""
             .Update
         End With
    Else
         With mrsSource
             !���� = lng����
             .Update
         End With
    End If
    vsUnits.TextMatrix(Row, Col - 1) = lng����
    If mrsUnitsReg.RecordCount > 0 Then
             With mrsUnitsReg
                 !���� = lng��������
                 .Update
             End With
     Else
        With mrsUnitsReg
            .AddNew
            !������λ = str��λ
            !�ƻ�Id = mlng�ƻ�Id
            !������Ŀ = str������Ŀ
            !��� = IIf(mrsSource.RecordCount > 0, mrsSource!���, 0)
            !���� = lng��������
            !ʱ��� = strʱ���
            .Update
        End With
     End If
     mrsUnitsReg.Filter = 0
     mrsSource.Filter = 0
End Sub

Private Sub vsUnits_GotFocus()
    Call zl_VsGridGotFocus(vsUnits)
End Sub

Private Sub vsUnits_LostFocus()
        Call zl_VsGridLOSTFOCUS(vsUnits, GRD_LOSTFOCUS_COLORSEL)
End Sub

Private Sub vsUnits_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsUnits
        If Not mblnʱ�� Then
            If KeyCode = vbKeyDelete Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlVsMoveGridCell vsUnits, 0, vsUnits.Cols - 1
End Sub

Private Sub vsUnits_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
   Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If mblnʱ�� Then Exit Sub
    With vsUnits
        If Not mblnʱ�� Then
            If KeyCode = vbKeyDelete Then
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlVsMoveGridCell vsUnits, 0, vsUnits.Cols - 1
End Sub

Private Sub vsUnits_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsUnits
        If mblnʱ�� Then
            If Col Mod 3 <> 2 Then Cancel = True: Exit Sub
            If .TextMatrix(Row, Col - 1) = "" Then Cancel = True: Exit Sub
             Exit Sub
         End If
         If .Cell(flexcpChecked, Row, .Cols - 1, Row, .Cols - 1) = 1 And Col <> .Cols - 1 Then Cancel = True: Exit Sub
         Select Case Col
         Case .ColIndex("������λ")
             Cancel = True: Exit Sub
         Case Else
            If .TextMatrix(Row, .ColIndex("������λ")) = "" Then Cancel = True: Exit Sub
         End Select
    End With
End Sub
Private Sub vsUnits_AfterMoveColumn(ByVal Col As Long, Position As Long)
    If mblnʱ�� = False Then
        zl_vsGrid_Para_Save mlngModule, vsUnits, Me.Caption, "��������_��ʱ���"
    End If
End Sub
Private Sub vsUnits_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If mblnʱ�� = False Then
        zl_vsGrid_Para_Save mlngModule, vsUnits, Me.Caption, "��������_��ʱ���"
    End If
End Sub
