VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ű�����"
   ClientHeight    =   6684
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11280
   Icon            =   "frmPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6684
   ScaleWidth      =   11280
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd���� 
      Caption         =   "�����µ��Ű�(&D)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   1932
   End
   Begin VB.Frame fraCon 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdCur 
         Caption         =   "����(&C)"
         Height          =   350
         Left            =   2760
         TabIndex        =   8
         ToolTipText     =   "�ȼ���F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdRe 
         Caption         =   "ˢ��&R)"
         Height          =   350
         Left            =   9360
         TabIndex        =   7
         ToolTipText     =   "�ȼ���F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "��һ��(&N)"
         Height          =   350
         Left            =   7160
         TabIndex        =   6
         ToolTipText     =   "�ȼ���F2"
         Top             =   215
         Width           =   1215
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "��һ��(&L)"
         Height          =   350
         Left            =   4960
         TabIndex        =   5
         ToolTipText     =   "�ȼ���F2"
         Top             =   215
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Dtp��ʼʱ�� 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3620
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   112066563
         CurrentDate     =   39998
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   2
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPlan 
      Height          =   5145
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10800
      _cx             =   19050
      _cy             =   9075
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16771280
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPlan.frx":6852
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
Attribute VB_Name = "frmPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����id As Long
Private mstr���� As String
Private mdateCur As Date
Private mrs��Һ̨ As Recordset
Private mrs���� As Recordset
Private mrs��Ա As New Recordset
Private mbln���� As Boolean

Private mintRow As Integer
Private mintCol As Integer

Private Sub CheckDate(ByVal Row As Long, ByVal Col As Long)
    With Me.vsfPlan
        If Col = .ColIndex("��ҩ��") Or Col = .ColIndex("��Һ��") Or Col = .ColIndex("�˶���") Or Col = .ColIndex("������") Or Col = .ColIndex("�����") Then
            If Not .TextMatrix(Row, Col) = "" Then
                mrs��Ա.Filter = "���� = '" & .TextMatrix(Row, Col) & "'"
                If mrs��Ա.RecordCount = 0 Then
                    mrs��Ա.Filter = "���� = '" & UCase(.TextMatrix(Row, Col)) & "'"
                    If mrs��Ա.RecordCount = 0 Then
                        MsgBox "δƥ�䵽�����Ա,����������"
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    Else
                        .TextMatrix(Row, Col) = mrs��Ա!����
                        If Col <> .Cols - 1 Then .Col = .Col + 1
                    End If
                Else
                    If Col <> .Cols - 1 Then .Col = .Col + 1
                End If
            End If
        End If
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub InitCom()
    Dim strPreson As String
    Dim strsql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    mdateCur = zldatabase.Currentdate
    Dtp��ʼʱ��.Value = mdateCur
    
    strsql = "Select a.Id, a.����, a.����" & vbNewLine & _
            "From ��Ա�� A, ������Ա B" & vbNewLine & _
            "Where a.Id = b.��Աid And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.����id =[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "initVSF", mlng����id)
    
    Set mrs��Ա = rsTemp
    
    Do While Not rsTemp.EOF
        strPreson = strPreson & IIf(strPreson = "", "|", "") & rsTemp!���� & "|"
        rsTemp.MoveNext
    Loop
    
    With vsfPlan
        .ColComboList(.ColIndex("��ҩ��")) = strPreson
        .ColComboList(.ColIndex("��Һ��")) = strPreson
        .ColComboList(.ColIndex("�˶���")) = strPreson
        .ColComboList(.ColIndex("������")) = strPreson
        .ColComboList(.ColIndex("�����")) = strPreson
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub InitVSF()
    Dim rsTemp As Recordset
    Dim strsql As String
    Dim i As Integer
    Dim j As Integer
    Dim count As Integer
 
    strsql = "select A.��ҩ̨id,B.����,A.����,A.�����,A.��ҩ��,A.�˶���,A.��Һ��,A.������ from ��Һ�������� A,��Һ̨ B where  A.��ҩ̨id=B.id and  A.����id= B.����id and A.����id=[1] and A.����=[2]"
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "initVSF", mlng����id, CDate(Format(Dtp��ʼʱ��.Value, "Short Date")))
    
    With Me.vsfPlan
        
        If rsTemp.EOF Then
            Me.vsfPlan.Editable = flexEDKbdMouse
            .rows = (mrs��Һ̨.RecordCount * mrs����.RecordCount) + 1
            mrs��Һ̨.MoveFirst
            For i = 1 To mrs��Һ̨.RecordCount
                mrs����.MoveFirst
                For j = 1 To mrs����.RecordCount
                    count = count + 1
                    .TextMatrix(count, .ColIndex("��Һ̨��")) = mrs��Һ̨!����
                    .TextMatrix(count, .ColIndex("��Һ̨id")) = mrs��Һ̨!Id
                    .TextMatrix(count, .ColIndex("��ҩ����")) = mrs����!����
                    .TextMatrix(count, .ColIndex("��ҩ��")) = ""
                    .TextMatrix(count, .ColIndex("��Һ��")) = ""
                    .TextMatrix(count, .ColIndex("�˶���")) = ""
                    .TextMatrix(count, .ColIndex("������")) = ""
                    .TextMatrix(count, .ColIndex("�����")) = ""
                    mrs����.MoveNext
                Next
                mrs��Һ̨.MoveNext
            Next
            
            Exit Sub
        End If
        Me.vsfPlan.Editable = flexEDNone
        Do While Not rsTemp.EOF
            i = i + 1
            .rows = i + 1
            .TextMatrix(i, .ColIndex("��Һ̨��")) = rsTemp!����
            .TextMatrix(i, .ColIndex("��Һ̨id")) = rsTemp!��ҩ̨id
            .TextMatrix(i, .ColIndex("��ҩ����")) = rsTemp!����
            .TextMatrix(i, .ColIndex("��ҩ��")) = NVL(rsTemp!��ҩ��)
            .TextMatrix(i, .ColIndex("��Һ��")) = NVL(rsTemp!��Һ��)
            .TextMatrix(i, .ColIndex("�˶���")) = NVL(rsTemp!�˶���)
            .TextMatrix(i, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(i, .ColIndex("�����")) = NVL(rsTemp!�����)

            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal lng����id As Long, ByVal frmParent As Object)
    
    Dim rsTemp As Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "select id,���� from ��Һ̨ where ����id=[1]"
    Set mrs��Һ̨ = zldatabase.OpenSQLRecord(strsql, "initVSF", lng����id)
    
    If mrs��Һ̨.EOF Then
        MsgBox "���Ƚ�����Һ̨�������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strsql = "select ���� from ��ҩ�������� where ��������id=[1]"
    Set mrs���� = zldatabase.OpenSQLRecord(strsql, "initVSF", lng����id)
    
    If mrs����.EOF Then
        MsgBox "���Ƚ��й������ν������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mlng����id = lng����id
    Me.Show 1, frmParent
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCur_Click()
    Dtp��ʼʱ��.Value = mdateCur
    Call cmdRe_Click
End Sub

Private Sub cmdLast_Click()
    Dtp��ʼʱ��.Value = Dtp��ʼʱ��.Value - 1
    Call cmdRe_Click
End Sub

Private Sub cmdNext_Click()
    Dtp��ʼʱ��.Value = Dtp��ʼʱ��.Value + 1
    Call cmdRe_Click
End Sub

Private Sub cmdRe_Click()
    If mbln���� = False Then
        If MsgBox("��ǰ���ݻ�δ���棬�Ƿ������ǰ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Call InitVSF
    
    Me.vsfPlan.Editable = (Me.Dtp��ʼʱ��.Value >= mdateCur)
    cmdSave.Enabled = (Me.Dtp��ʼʱ��.Value >= mdateCur)
End Sub

Private Sub cmdSave_Click()
    Dim strsql As String
    Dim i  As Integer
    Dim arrSql As Variant
    Dim j As Integer
    
    arrSql = Array()
    With Me.vsfPlan
        For i = 1 To .rows - 1
            
            If Val(.TextMatrix(i, .ColIndex("��Һ̨id"))) = 0 Then
                Exit Sub
            End If
            strsql = "Zl_��Һ��������_����("
            strsql = strsql & mlng����id
            strsql = strsql & ",to_date('" & Format(Dtp��ʼʱ��.Value, "Short Date") & "' ,'yyyy-mm-dd')"
            strsql = strsql & "," & Val(.TextMatrix(i, .ColIndex("��Һ̨id")))
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��ҩ����")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("�����")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��ҩ��")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("�˶���")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��Һ��")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("������")) & "'"
            strsql = strsql & "," & i
            strsql = strsql & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strsql
        Next
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "CmdSave_Click")
    Next
    gcnOracle.CommitTrans
    mbln���� = True
    If MsgBox("����ɹ����Ƿ�����Űࣿ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
        Unload Me
    End If
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Call cmdRe_Click
    frmPlanCopy.ShowCard Me, Dtp��ʼʱ��.Value, mlng����id
End Sub

Private Sub Form_Load()
    mbln���� = True
    Call InitCom
End Sub

Private Sub vsfPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call CheckDate(Row, Col)
End Sub

Private Sub vsfPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    
    With Me.vsfPlan
        If KeyAscii = 13 Then
            If Col = .ColIndex("��Һ̨��") Or Col = .ColIndex("��Һ̨id") Or Col = .ColIndex("��ҩ����") Then
                .Col = Col + 1
            Else
                Call CheckDate(Row, Col)
            End If
        End If
    End With
End Sub
