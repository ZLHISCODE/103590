VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmIdentify�����������Ǽ� 
   Caption         =   "����ҽ���������Ǽ�"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "frmIdentify�����������Ǽ�.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   8385
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraҩƷ��ϸ��ѯ 
      Caption         =   "ѡ�����û�"
      Height          =   2955
      Left            =   165
      TabIndex        =   2
      Top             =   105
      Width           =   8055
      Begin VB.TextBox txtҽ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   1965
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1380
         Width           =   1965
      End
      Begin VB.TextBox txt���֤�� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1830
         Width           =   1965
      End
      Begin VB.TextBox txt�Ա� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1830
         Width           =   1965
      End
      Begin VB.TextBox txt��Ա��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1395
         Width           =   1965
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   885
         Width           =   1965
      End
      Begin VB.CommandButton cmdѡ�� 
         Caption         =   "��"
         Height          =   300
         Left            =   7380
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox txt��λ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2340
         Width           =   1965
      End
      Begin VB.TextBox txt��λ���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2340
         Width           =   1965
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2565
         TabIndex        =   12
         Top             =   270
         Width           =   5085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ID��������סԺ��(&I)"
         Height          =   180
         Left            =   255
         TabIndex        =   21
         Top             =   360
         Width           =   2070
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   10385
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   10385
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Label labҽ���� 
         AutoSize        =   -1  'True
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   420
         TabIndex        =   20
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lab���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   5055
         TabIndex        =   19
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lab��Ա��� 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lab�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   600
         TabIndex        =   17
         Top             =   1890
         Width           =   360
      End
      Begin VB.Label lab���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   5055
         TabIndex        =   16
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label lab���֤�� 
         AutoSize        =   -1  'True
         Caption         =   "���֤��"
         Height          =   180
         Left            =   4695
         TabIndex        =   15
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label lab��λ���� 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Left            =   4695
         TabIndex        =   14
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label lab��λ���� 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   5535
      TabIndex        =   1
      Top             =   5655
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   6945
      TabIndex        =   0
      Top             =   5655
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfProject 
      Height          =   2355
      Left            =   165
      TabIndex        =   22
      ToolTipText     =   "Shift+Deleteɾ����ǰ��"
      Top             =   3120
      Width           =   8055
      _cx             =   14208
      _cy             =   4154
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIdentify�����������Ǽ�.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
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
   Begin XtremeCommandBars.CommandBars cbrDelete 
      Left            =   3375
      Top             =   5685
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ�������շ�ϸĿ����ѡ��"
      Height          =   180
      Left            =   210
      TabIndex        =   23
      Top             =   5737
      Width           =   2520
   End
End
Attribute VB_Name = "frmIdentify�����������Ǽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytEditMode            As Byte             '�ж��ǵǼǻ����޸ĵǼ�
Private mstrҽ����              As String
Private mrsD                    As ADODB.Recordset
Private mintInsure              As Integer
Private mstrSortID              As String
Private mblnCancel              As Boolean

Dim rsTmp                       As ADODB.Recordset
Dim strSQL                      As String
Dim sngX                        As Single
Dim sngY                        As Single
Dim sngH                        As Single

Const strSickFields = "select a.ҽ���� as ID,a.ҽ����, a.����,a.��Ա��� as ��Ա���,b.����,b.�Ա�,b.���֤��,b.��ͬ��λid as ��λ����,b.������λ as ��λ���� from �����ʻ� a , ������Ϣ b where a.����ID = b.����id And a.���� =" & TYPE_������
Const strProjectFields = "select id,decode(���,'5','����ҩ','6','�г�ҩ','7','�в�ҩ','�������') as ���,id as �շ�ϸĿID,���� as �շ�ϸĿ����,���� as �շ�ϸĿ����,��� as �շ�ϸĿ��� from �շ�ϸĿ  "

Public Property Get blnCancel()
    blnCancel = mblnCancel
End Property

Public Property Let bytEditMode(ByVal vEditMode As Byte)
    mbytEditMode = vEditMode
End Property

Public Property Let intinsure(ByVal vintInsure As Integer)
    mintInsure = vintInsure
End Property
 
Public Property Let strҽ����(ByVal vstrҽ���� As String)
    mstrҽ���� = vstrҽ����
End Property

Private Sub cbrDelete_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 0
            vsfProject_Delete
    End Select
End Sub

Private Sub cmdCancle_Click()
On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim strTableD       As String
    Dim strWhereD       As String
    Dim i               As Integer
    Dim sFileName       As String
    Dim blnTran         As Boolean
On Error GoTo ErrH
    blnTran = False
    If txtҽ����.Text = "" Then
        MsgBox "δѡ��ҽ����Ա��", vbInformation, gstrSysName
        Exit Sub
    ElseIf vsfProject.Tag <> "TRUE" Then
        MsgBox "��������Ŀδ���ģ�" & vbCrLf & "����ȡ��", vbInformation, gstrSysName
        Exit Sub
    End If
    mstrҽ���� = txtҽ����.Text
    If mbytEditMode = 2 Then
        '��¼�޸�ǰ��־
        ' �����(�÷ֺ�";"����)
        strTableD = "ҽ��������_����;ҽ����������Ŀ_����"
        ' �������(�÷ֺ�";"����)
        strWhereD = "ҽ����='" & txtҽ����.Text & "';ҽ����='" & txtҽ����.Text & "'"
        ' ��¼�޸�ǰ������
        sFileName = EditFormerWriteFileA(strTableD, strWhereD)
    End If
    With gcnGYYB
        .BeginTrans
        blnTran = True
        '������������
        strSQL = "Zl_ҽ��������_����_Update ('" & txtҽ����.Text & "','" & txt����.Text & "','" & txt��Ա���.Text & "','" & txt����.Text & "','" & txt�Ա�.Text & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "','" & txt��λ����.Text & "',1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        strSQL = "Zl_ҽ����������Ŀ_����_Delete ('" & txtҽ����.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        For i = 1 To vsfProject.Rows - 1
            If vsfProject.TextMatrix(i, vsfProject.ColIndex("�շ�ϸĿID")) <> "" Then
                strSQL = "Zl_ҽ����������Ŀ_����_Update ('" & txtҽ����.Text & "','" & vsfProject.TextMatrix(i, vsfProject.ColIndex("�շ�ϸĿID")) & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        blnTran = False
        .CommitTrans
    End With
    
    If mbytEditMode = 2 Then
        '��¼�޸ĺ���־
        Call EditFormerWriteFileA(strTableD, strWhereD, sFileName)
        '�����޸���־
        AddLog "ҽ������", "ҽ��������_����", DBConnLTEdit, , sFileName, mstrҽ����, , , "ҽ��������_����", , True
    End If
    Unload Me
    Exit Sub
ErrH:
    If blnTran Then
        gcnGYYB.RollbackTrans
        gcnGYYB.Errors.Clear
    End If
    Err.Clear
End Sub

Private Sub Form_Load()
On Error GoTo ErrH
    If mbytEditMode = 2 Then
        strSQL = strSickFields
        strSQL = strSQL & " And a.ҽ���� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrҽ����)
        If Not ChkRsState(rsTmp) Then
            txtҽ����.Text = Nvl(rsTmp!ҽ����)
            txt����.Text = Nvl(rsTmp!����)
            txt��Ա���.Text = Nvl(rsTmp!��Ա���)
            txt����.Text = Nvl(rsTmp!����)
            txt�Ա�.Text = Nvl(rsTmp!�Ա�)
            txt���֤��.Text = Nvl(rsTmp!���֤��)
            txt��λ����.Text = Nvl(rsTmp!��λ����)
            txt��λ����.Text = Nvl(rsTmp!��λ����)
        End If
        txtFind.Locked = True
        txtFind.BackColor = &H80000000
        cmdѡ��.Enabled = False
    End If
    Call dDataload
    '
    With cbrDelete.KeyBindings
        .Add 4, vbKeyDelete, 0                 'Shift +Delete
        .Add 0, vbKeyDelete, 0                 'Shift +Delete
    End With
    
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub dDataload()
    On Error GoTo ErrH
    strSQL = "SELECT A.�շ�ϸĿID,���,�շ�ϸĿ����,�շ�ϸĿ����,�շ�ϸĿ��� " & vbCrLf & _
            "FROM ҽ����������Ŀ_���� A ,(SELECT DECODE(���,'5','����ҩ','6','�г�ҩ','�������') AS ���, ID AS �շ�ϸĿID, ���� AS �շ�ϸĿ����,���� AS �շ�ϸĿ����,��� AS �շ�ϸĿ��� FROM �շ�ϸĿ ) B" & vbCrLf & _
            "WHERE A.�շ�ϸĿID = B.�շ�ϸĿID"
    strSQL = strSQL & vbCrLf & "And A.ҽ���� =[1]   "
    Set mrsD = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrҽ����)
    Set vsfProject.DataSource = mrsD
    vsfProject.Rows = vsfProject.Rows + 1
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub cmdѡ��_Click()
    strSQL = strSickFields & " And A.ҽ���� not in (Select ҽ���� From ҽ����������Ŀ_����)"
    Call SickSelect(strSQL)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode <> 13 Then Exit Sub
    Dim strCode As String, strWhere As String
    '����ǿ��������ȡ��ʽ
    strCode = txtFind.Text
    If (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then
        '����ID
        strWhere = " And A.����ID=" & Mid(strCode, 2)
    ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then
        'סԺ��
        strWhere = " And b.סԺ��='" & Mid(strCode, 2) & "'"
    Else
        'ҽ����
        strWhere = " And (b.���� Like '%" & strCode & "%' or A.ҽ���� like '%" & strCode & "%')"
    End If
    strSQL = strSickFields & " And A.ҽ���� not in (Select ҽ���� From ҽ����������Ŀ_����) " & strWhere
    Call SickSelect(strSQL)
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub SickSelect(sSql As String)
    Dim vRect       As RECT
    On Error GoTo ErrH
    vRect = GetControlRect(txtFind.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txtFind.Height
    strSQL = sSql
    Set rsTmp = zlDatabase.ShowSQLSelect( _
            Nothing, strSQL, 0, "ҽ������ѡ��", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, txtFind.Text _
            )
    If Not ChkRsState(rsTmp) Then
        txtҽ����.Text = Nvl(rsTmp!ҽ����)
        txt����.Text = Nvl(rsTmp!����)
        txt��Ա���.Text = Nvl(rsTmp!��Ա���)
        txt����.Text = Nvl(rsTmp!����)
        txt�Ա�.Text = Nvl(rsTmp!�Ա�)
        txt���֤��.Text = Nvl(rsTmp!���֤��)
        txt��λ����.Text = Nvl(rsTmp!��λ����)
        txt��λ����.Text = Nvl(rsTmp!��λ����)
    Else
        MsgBox "û���ҵ�������Ϣ!", vbInformation, gstrSysName
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
 
Private Sub vsfProject_AfterEdit(ByVal Row As Long, ByVal COL As Long)
    On Error GoTo ErrH
    vsfProject.Tag = "TRUE"
    Call vsfProject_KeyPressEdit(Row, COL, 13)
    If COL = vsfProject.ColIndex("�շ�ϸĿ����") Then

        vsfProject.EditText = UCase(vsfProject.EditText)
        strSQL = strProjectFields & "  where (���� like '%' || [1] || '%' or ����  like '%' || [1] || '%' or zlSpellCode(����)  like '%' || [1] || '%')"
        Call CalcPosition(sngX, sngY, vsfProject)
        sngY = sngY - vsfProject.CellHeight
        sngH = vsfProject.CellHeight
        DoEvents
        If Trim(vsfProject.EditText) = "" Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = ""
            Exit Sub
        End If
        Set rsTmp = zlDatabase.ShowSQLSelect( _
                Nothing, strSQL, 0, "�շ�ϸĿ����ѡ��", False, _
                "", "", False, False, True, _
                sngX, sngY, sngH, False, False, _
                False, vsfProject.EditText _
                )
        If ChkRsState(rsTmp) Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = ""
        Else
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = rsTmp!ID
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = rsTmp!�շ�ϸĿ����
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = rsTmp!�շ�ϸĿ����
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = rsTmp!���
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = rsTmp!�շ�ϸĿ���
            If vsfProject.Rows - 1 = vsfProject.Row Then vsfProject.Rows = vsfProject.Rows + 1
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    On Error GoTo ErrH
    With vsfProject
        Select Case COL
            Case .ColIndex("�շ�ϸĿ����")
                vsfProject.ComboList = "|..."
            Case Else
                .ComboList = ""
                Cancel = True
        End Select
        
    End With
    Exit Sub
ErrH:
    Err.Clear
    
End Sub

'==============================================================================
'=���ܣ� �����λ��¼ vsfProject
'==============================================================================
Private Sub vsfProject_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfProject.FindRow(mstrSortID, -1, vsfProject.ColIndex("����ID"), False, True)
    If lngRow > 0 Then vsfProject.Row = lngRow
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_BeforeMoveColumn(ByVal COL As Long, Position As Long)
    If COL = vsfProject.ColIndex("����ID") Then
        Position = -1
    Else
        If Position <= vsfProject.ColIndex("����ID") Then Position = COL
    End If
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С vsfProject[ͼ��]
'==============================================================================
Private Sub vsfProject_BeforeUserResize(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If COL = vsfProject.ColIndex("����ID") Then Cancel = True
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼ID vsfProject
'==============================================================================
Private Sub vsfProject_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("����ID"))
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_CellButtonClick(ByVal Row As Long, ByVal COL As Long)
    On Error GoTo ErrH
    vsfProject.Tag = "TRUE"
    If vsfProject.ColIndex("�շ�ϸĿ����") = COL Then
        strSQL = strProjectFields
             
        Call CalcPosition(sngX, sngY, vsfProject)
        sngY = sngY - vsfProject.CellHeight
        sngH = vsfProject.CellHeight
        
        
        Set rsTmp = zlDatabase.ShowSQLSelect( _
                Nothing, strSQL, 0, "�շ�ϸĿ����ѡ��", False, _
                "", "", False, False, True, _
                sngX, sngY, sngH, False, False, _
                False, "" _
                )
        If ChkRsState(rsTmp) Then
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = ""
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = ""
        Else
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = rsTmp!ID
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = rsTmp!�շ�ϸĿ����
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = rsTmp!�շ�ϸĿ����
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = rsTmp!���
            vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = rsTmp!�շ�ϸĿ���
            If vsfProject.Rows - 1 = vsfProject.Row Then vsfProject.Rows = vsfProject.Rows + 1
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_DblClick()
On Error GoTo ErrH
    If vsfProject.MouseRow = 0 Or vsfProject.MouseCol = 0 Then Exit Sub
    If vsfProject.TextMatrix(vsfProject.MouseRow, vsfProject.MouseCol) = "" Then Exit Sub
    Clipboard.SetText (vsfProject.TextMatrix(vsfProject.MouseRow, vsfProject.MouseCol))
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If vsfProject.ColIndex("�շ�ϸĿ����") = vsfProject.COL Then
        '�ո�༭
        If KeyAscii = vbKeySpace Then
            'KeyAscii = 39
            KeyAscii = 0
            SendKeys "{f2}"
        End If
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub vsfProject_KeyPressEdit(ByVal Row As Long, ByVal COL As Long, KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = asc("'") Then
       KeyAscii = 0
    End If
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Function vsfProject_Delete() As Long
    
    If vsfProject.Row = 1 And vsfProject.Rows = 2 Then
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿID")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ����")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("�շ�ϸĿ���")) = ""
        vsfProject.TextMatrix(vsfProject.Row, vsfProject.ColIndex("���")) = ""
        Exit Function
    End If
    vsfProject.RemoveItem vsfProject.Row

End Function


