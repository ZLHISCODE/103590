VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffLimit 
   Caption         =   "���Ĵ�������"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmStuffLimit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9060
   StartUpPosition =   1  '����������
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3105
      Left            =   3285
      TabIndex        =   14
      Top             =   6015
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   5477
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraFunc 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -15
      TabIndex        =   6
      Top             =   4350
      Width           =   9810
      Begin VB.CommandButton cmdӦ���ڱ����������� 
         Caption         =   "Ӧ���ڱ���(&O)"
         Height          =   350
         Left            =   3990
         TabIndex        =   13
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "�ָ�(&R)"
         Height          =   350
         Left            =   2685
         Picture         =   "frmStuffLimit.frx":058A
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ�����(&C)"
         Height          =   350
         Left            =   1395
         Picture         =   "frmStuffLimit.frx":06D4
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   5700
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   90
         Picture         =   "frmStuffLimit.frx":081E
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�(&X)"
         Height          =   350
         Left            =   6810
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
   End
   Begin ZL9BillEdit.BillEdit msfLimit 
      Height          =   2655
      Left            =   75
      TabIndex        =   1
      Top             =   1380
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -195
      TabIndex        =   5
      Top             =   1035
      Width           =   9810
   End
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      Text            =   "cboRoom"
      Top             =   585
      Width           =   3360
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5715
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffLimit.frx":0968
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   6480
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5760
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ���Ŀⷿ(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   645
      Width           =   1350
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmStuffLimit.frx":11FA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ѡ��ⷿ��ָ���ÿⷿ�������ϵĴ����������������������ϵĹ���Ҫ�󣬿���ͬʱָ�����̵����ԺͿⷿ��λ��"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   135
      Width           =   7725
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "���������ڸ��ⷿ���޶����̵�Ҫ��(&T)��"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   3330
   End
End
Attribute VB_Name = "frmStuffLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ���ʣ���me.tag����4
'   2����ǰ״̬����me.cmdClose.tag���棬�ֱ�Ϊ"�޸�"��"����"�����ϼ�������
'   3��ָ�����ģ���me.lblMedi.tag���棬���ϼ���������Դ��ݣ�Ҳ���Բ�����
'---------------------------------------------------
Public strPrivs As String       '��ǰ�û����еı�����Ȩ��

Dim mobjItem As ListItem
Dim mLngCount As Long

Dim mrsTemp As New ADODB.Recordset
Private Const col���� As Integer = 1
Private Const col���� As Integer = 2
Private Const col��� As Integer = 3
Private Const col���� As Integer = 4
Private Const col�ɱ��� As Integer = 5
Private Const col���ۼ� As Integer = 6
Private Const col������� As Integer = 7
Private Const col��λ As Integer = 8
Private Const col��װ As Integer = 9
Private Const col���� As Integer = 10
Private Const col���� As Integer = 11
Private Const col���� As Integer = 12
Private Const col���� As Integer = 13
Private Const col���� As Integer = 14
Private Const col���� As Integer = 15
Private Const col��λ As Integer = 16
Private mlngFind As Long
Private mblnFind As Boolean             '�Ƿ��ѯ��ֵ
Private mblnFindFrist As Boolean        '����û���ҵ�������
Private mblnNoClick As Boolean

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub cboRoom_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long
    
    err = 0: On Error GoTo ErrHand
    
    If mblnNoClick Then Exit Sub
    
    gstrSQL = "Select ��������,nvl(�������,0) as ������� From ��������˵��" & _
            " Where ����id=[1]"
    gstrSQL = gstrSQL & " and  �������� In ('���Ŀ�','���ϲ���', '����ⷿ') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex))
    
    With rsTemp
        Me.cboRoom.Tag = "�ۼ�"
        Do While Not .EOF
            If InStr(1, !��������, "���Ŀ�") > 0 Or InStr(1, !��������, "����ⷿ") > 0 Then
                Me.cboRoom.Tag = "���Ŀ�"
                Exit Do
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        
        Do While Not .EOF
            If InStr(1, !��������, "���ϲ���") > 0 And (!������� = 1) Then
                Me.cboRoom.Tag = "����"
                Exit Do
            End If
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, !��������, "���ϲ���") > 0 And (!������� = 2 Or !������� = 3) Then Me.cboRoom.Tag = "סԺ": Exit Do
            .MoveNext
        Loop
    End With
    
    Call zlLimitRef
    
    With msfLimit
        If .Rows = 2 And .RowData(1) = 0 Then
            For lngCol = 0 To .Cols - 1
                .ColData(lngCol) = 5
            Next
        ElseIf .Rows > 2 Then
            .Redraw = False
            .ColData(col����) = 4
            .ColData(col����) = 4
            .ColData(col����) = -1
            .ColData(col����) = -1
            .ColData(col����) = -1
            .ColData(col����) = -1
            .ColData(col��λ) = 1
            .SetColColor col����, vbWhite
            .SetColColor col����, vbWhite
            .SetColColor col����, vbWhite
            .SetColColor col����, vbWhite
            .SetColColor col����, vbWhite
            .SetColColor col����, vbWhite
            .SetColColor col��λ, vbWhite
            .SetRowColor 0, &H8000000F
            .Redraw = True
        End If
        For lngCol = 0 To .Cols - 1
            If .ColData(lngCol) = 5 Then
                .SetColColor lngCol, &H8000000F
            End If
        Next
        .SetFocus
        .Col = col����
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRoom.ListCount = 0 Then Call zlControl.ControlSetFocus(msfLimit): Exit Sub
    
    If cboRoom.ListIndex >= 0 Then
        If Val(cboRoom.Tag) = cboRoom.ItemData(cboRoom.ListIndex) Then
            Call zlControl.ControlSetFocus(msfLimit, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboRoom, Trim(cboRoom.Text), "V,K,12,W", IIf(InStr(1, strPrivs, "�����������пⷿ�޶��̵�") = 0, True, False)) = False Then
        Exit Sub
    End If
    If cboRoom.ListIndex >= 0 Then
        cboRoom.Tag = cboRoom.ItemData(cboRoom.ListIndex)
    End If
End Sub


Private Sub cboRoom_LostFocus()
    Dim i As Long
    If cboRoom.ListCount = 0 Then Exit Sub
    If cboRoom.ListIndex < 0 Then
        For i = 0 To cboRoom.ListCount - 1
            If Val(cboRoom.Tag) = cboRoom.ItemData(i) Then
                mblnNoClick = True
                cboRoom.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub


Private Sub cmdClear_Click()
    With Me.msfLimit
        .Redraw = False
        For mLngCount = 1 To .Rows - 1
            .TextMatrix(mLngCount, 0) = ""
            If InStr(1, strPrivs, "�����޿���") > 0 Then
                .TextMatrix(mLngCount, col����) = Format(0, mFMT.FM_����)
                .TextMatrix(mLngCount, col����) = Format(0, mFMT.FM_����)
            End If
            If InStr(1, strPrivs, "�̵���������") > 0 Then
                .TextMatrix(mLngCount, col����) = ""
                .TextMatrix(mLngCount, col����) = ""
                .TextMatrix(mLngCount, col����) = ""
                .TextMatrix(mLngCount, col����) = ""
            End If
            .TextMatrix(mLngCount, col��λ) = ""
        Next
        .Redraw = True
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlLimitRef
End Sub

Private Sub CmdSave_Click()
    Dim strMsgBox As String, strErrors As String
    
    strErrors = ""
    With Me.msfLimit
        For mLngCount = 1 To .Rows - 1
            If Val(.TextMatrix(mLngCount, col����)) <> 0 _
                And Val(.TextMatrix(mLngCount, col����)) < Val(.TextMatrix(mLngCount, col����)) Then
                .TextMatrix(mLngCount, 0) = "��"
                strErrors = strErrors & vbCrLf & .TextMatrix(mLngCount, col����) & "-" & .TextMatrix(mLngCount, col����)
                strMsgBox = "��" & .TextMatrix(mLngCount, col����) & "-" & .TextMatrix(mLngCount, col����) & "���Ĵ������޴��ڴ������ޣ�" & _
                        vbCrLf & vbCrLf & "������������������"
                If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbThis.Panels(2).Text = ""
                    .MsfObj.TopRow = mLngCount: .Row = mLngCount: .SetFocus: Exit Sub
                End If
            ElseIf .RowData(mLngCount) <> 0 Then
                gstrSQL = "zl_�������ϴ����޶�_Update(" & Me.cboRoom.ItemData(Me.cboRoom.ListIndex)
                gstrSQL = gstrSQL & "," & .RowData(mLngCount)
                gstrSQL = gstrSQL & "," & Round(Val(.TextMatrix(mLngCount, col����)) * Val(.TextMatrix(mLngCount, col��װ)), g_С��λ��.obj_ɢװС��.����С��)
                gstrSQL = gstrSQL & "," & Round(Val(.TextMatrix(mLngCount, col����)) * Val(.TextMatrix(mLngCount, col��װ)), g_С��λ��.obj_ɢװС��.����С��)
                gstrSQL = gstrSQL & ",'" & IIf(Trim(.TextMatrix(mLngCount, col����)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col����)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col����)) = "", "0", "1")
                gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(mLngCount, col����)) = "", "0", "1")
                gstrSQL = gstrSQL & "','" & Trim(.TextMatrix(mLngCount, col��λ)) & "')"
                err = 0: On Error Resume Next
                zldatabase.ExecuteProcedure gstrSQL, Me.Caption
                
                If err <> 0 Then
                    Call SaveErrLog
                    err = 0
                    .TextMatrix(mLngCount, 0) = "��"
                    strErrors = strErrors & vbCrLf & .TextMatrix(mLngCount, col����) & "-" & .TextMatrix(mLngCount, col����)
                    strMsgBox = "���桰" & .TextMatrix(mLngCount, col����) & .TextMatrix(mLngCount, col����) & "��ʱ��������" & _
                            vbCrLf & vbCrLf & "������������������"
                    If MsgBox(strMsgBox, vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Me.stbThis.Panels(2).Text = ""
                        .MsfObj.TopRow = mLngCount: .Row = mLngCount: .SetFocus: Exit Sub
                    End If
                End If
                If mLngCount Mod IIf(.Rows > 20, .Rows \ 20, 1) = 0 Then
                    Me.stbThis.Panels(2).Text = "���ڱ��棺" & String(mLngCount \ IIf(.Rows > 20, .Rows \ 20, 1), "��")
                End If
            End If
        Next
    End With
    Me.stbThis.Panels(2).Text = ""
    strMsgBox = "��" & Me.cboRoom.Text & "���������Ա�����ϣ�"
    If strErrors <> "" Then
        strMsgBox = strMsgBox & vbCrLf & "���������ķ����������飺" & strErrors
    End If
    MsgBox strMsgBox, vbExclamation, gstrSysName
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdӦ���ڱ�����������_Click()
    Dim lngRow As Long, lngRows As Long
    Dim strValue As String
    '����ǰ�е�����Ӧ�õ�����ҩƷ��ͬ��
    lngRows = msfLimit.Rows - 1
    strValue = msfLimit.TextMatrix(msfLimit.Row, msfLimit.Col)
    For lngRow = 1 To lngRows
        msfLimit.TextMatrix(lngRow, msfLimit.Col) = strValue
    Next
    msfLimit.SetFocus
End Sub

Private Sub Form_Activate()
    Dim lngCol As Long
    
    With Me.msfLimit
        .Cols = 17: .MsfObj.FixedCols = 1
        .TextMatrix(0, col����) = "����": .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col���) = "���": .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col��λ) = "��λ": .TextMatrix(0, col��װ) = "��װ":
        .TextMatrix(0, col�ɱ���) = "�ɱ���": .TextMatrix(0, col���ۼ�) = "���ۼ�": .TextMatrix(0, col�������) = "�������":
        .TextMatrix(0, col����) = "����": .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col����) = "����": .TextMatrix(0, col����) = "����": .TextMatrix(0, col����) = "����": .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col��λ) = "��λ"

        .ColWidth(0) = 250: .ColWidth(col����) = 900: .ColWidth(col����) = 2200
        .ColWidth(col���) = 1500: .ColWidth(col����) = 1200
        .ColWidth(col��λ) = 500: .ColWidth(col��װ) = 0: .ColWidth(col�ɱ���) = 1200: .ColWidth(col���ۼ�) = 1200: .ColWidth(col�������) = 1200
        If InStr(1, strPrivs, "�����޿���") > 0 Then
            .ColWidth(col����) = 855: .ColWidth(col����) = 855
        Else
            .ColWidth(col����) = 0: .ColWidth(col����) = 0
        End If
        If InStr(1, strPrivs, "�̵���������") > 0 Then
            .ColWidth(col����) = 500: .ColWidth(col����) = 500: .ColWidth(col����) = 500: .ColWidth(col����) = 500
        Else
            .ColWidth(col����) = 0: .ColWidth(col����) = 0: .ColWidth(col����) = 0: .ColWidth(col����) = 0
        End If
        .ColWidth(col��λ) = 1700
        
        .ColAlignment(col����) = 1: .ColAlignment(col����) = 1
        .ColAlignment(col���) = 1: .ColAlignment(col����) = 1
        .ColAlignment(col��λ) = 4: .ColAlignment(col��װ) = 7: .ColAlignment(col�ɱ���) = 7: .ColAlignment(col���ۼ�) = 7: .ColAlignment(col�������) = 7
        .ColAlignment(col����) = 7: .ColAlignment(col����) = 7
        .ColAlignment(col����) = 4: .ColAlignment(col����) = 4: .ColAlignment(col����) = 4: .ColAlignment(col����) = 4
        .ColAlignment(col��λ) = 1
        
        .ColData(col����) = 5: .ColData(col����) = 5
        .ColData(col���) = 5: .ColData(col����) = 5
        .ColData(col��λ) = 5: .ColData(col��װ) = 5: .ColData(col�ɱ���) = 5: .ColData(col���ۼ�) = 5: .ColData(col�������) = 5
        If InStr(1, strPrivs, "�����޿���") > 0 Then
            .ColData(col����) = 4: .ColData(col����) = 4
        Else
            .ColData(col����) = 5: .ColData(col����) = 5
        End If
        If InStr(1, strPrivs, "�̵���������") > 0 Then
            .ColData(col����) = -1: .ColData(col����) = -1: .ColData(col����) = -1: .ColData(col����) = -1
        Else
            .ColData(col����) = 5: .ColData(col����) = 5: .ColData(col����) = 5: .ColData(col����) = 5
        End If
        .ColData(col��λ) = 1
        .PrimaryCol = col��λ:
        
        '��ȷ�����Ȩ��
        If InStr(1, strPrivs, "�����޿���") = 0 And InStr(1, strPrivs, "�����޿���") = 0 Then
        Else
            .LocateCol = IIf(InStr(1, strPrivs, "�����޿���") <> 0, col����, col����)
        End If
                    
        .Row = 1: .Col = IIf(InStr(1, strPrivs, "�����޿���") <> 0, col����, col����)
    End With
    
    If Me.cmdClose.Tag = "����" Then
        Me.msfLimit.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfLimit.Active = True
    End If
    
    err = 0: On Error GoTo ErrHand
    
    gstrSQL = "Select ID, ����, ����" & _
              "  From ���ű� D" & _
               " Where ID In (Select Distinct ����id" & _
                "             From ��������˵�� a" & _
                            " Where �������� In ('���ϲ���', '���ʿⷿ', '���Ŀ�', '�Ƽ���', '����ⷿ')) And" & _
                     " exists (Select 1  b From ����ִ�п��� b where d.id=b.ִ�п���id) and (d.����ʱ�� is null or to_char(d.����ʱ��,'yyyy-mm-dd')='3000-01-01')"
    If InStr(1, strPrivs, "�����������пⷿ�޶��̵�") = 0 Then
        gstrSQL = gstrSQL & "      and ID in (select ����ID from ������Ա R where R.��ԱID=[1])"
    End If
    
    Set mrsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id)
    
    With mrsTemp
        Me.cboRoom.Clear
        Do While Not .EOF
            Me.cboRoom.AddItem !���� & "-" & !����
            Me.cboRoom.ItemData(Me.cboRoom.NewIndex) = !Id
            .MoveNext
        Loop
    End With
    If Me.cboRoom.ListCount <= 0 Then
        MsgBox "δ������صĿⷿ���޷����ô�������", vbExclamation, gstrSysName
        Unload Me: Exit Sub
    End If
    Me.cboRoom.ListIndex = 0

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(1, g_�ɱ���)
        .FM_��� = GetFmtString(1, g_���)
        .FM_���ۼ� = GetFmtString(1, g_�ۼ�)
        .FM_���� = GetFmtString(1, g_����)
    End With
    lbl����.Visible = True
    txt����.Visible = True
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.fraLine.Left = 0: Me.fraLine.Width = Me.ScaleWidth + 100
    Me.msfLimit.Left = 0: Me.msfLimit.Width = Me.ScaleWidth
    Me.msfLimit.Height = Me.ScaleHeight - Me.msfLimit.Top - Me.fraFunc.Height - Me.stbThis.Height
    Me.fraFunc.Left = 0: Me.fraFunc.Width = Me.ScaleWidth: Me.fraFunc.Top = Me.msfLimit.Top + Me.msfLimit.Height
    Me.cmdClose.Left = Me.fraFunc.Width - Me.cmdClose.Width - 90
    Me.cmdSave.Left = Me.cmdClose.Left - Me.cmdSave.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFindFrist = False
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        msfLimit.SetFocus
        msfLimit.Col = col��λ
        Cancel = True
        Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub msfLimit_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub msfLimit_EditKeyPress(KeyAscii As Integer)
    If InStr("'!@#$%^&*|-""", Chr(KeyAscii)) <> 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub msfLimit_EnterCell(Row As Long, Col As Long)
    Dim lngCol As Long
    With msfLimit
      Select Case .Col
          Case col����, col����
              .TxtCheck = True
              .MaxLength = 15
              .TextMask = ".1234567890"
          Case col��λ
               ImeLanguage True
              .MaxLength = 20
              .TextMask = ""
              .TxtSetFocus
          Case Else
              .TxtCheck = False
          End Select
          For lngCol = 0 To .Cols - 1
              If .ColData(lngCol) = 5 Then
                  .SetColColor lngCol, &H8000000F
              End If
          Next
      End With
End Sub

Private Sub msfLimit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Me.msfLimit
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
        If .TextMatrix(.Row, col����) = "" Then
            .Text = " "
            .TextMatrix(.Row, col����) = .Text
            .TextMatrix(.Row, col����) = .Text
            Exit Sub
        End If
        
        If Trim(.Text) = "" Then
            .Text = IIf(.TextMatrix(.Row, .Col) = "", " ", IIf(.TxtVisible, .Text, .TextMatrix(.Row, .Col)))
            .TextMatrix(.Row, .Col) = .Text
        Else
            If .Col = col��λ Then
                If LenB(StrConv(.Text, vbFromUnicode)) > 50 Then
                    MsgBox "��λ���������50���ַ���25������", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
            Else
                If Not IsNumeric(.Text) Then
                    MsgBox "�����к��зǷ��ַ���", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
                If Val(.Text) < 0 Then
                    MsgBox "��������޲���С���㣡", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
                If Val(.Text) > 10000000000000# Then
                    MsgBox "����ֵ�������ֵ��", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
                End If
            End If
        End If
        
        Select Case .Col
        Case col����
            .Text = Format(.Text, mFMT.FM_����): .TextMatrix(.Row, col����) = .Text
        Case col����
            .Text = Format(.Text, mFMT.FM_����): .TextMatrix(.Row, col����) = .Text
        Case col��λ
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, col��λ) = ""
                    End If
                    Exit Sub
                Else
                    Dim strTemp As String
                    strTemp = GetMatchingSting(strKey)
                    
                    gstrSQL = " Select ����,���� From ���Ͽⷿ��λ " & _
                          " Where (���� Like [1]" & _
                          "     Or ���� Like [1]" & _
                          "     Or ���� Like [1])"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
                
                    
                    If rsTemp.EOF Then
                        If MsgBox("û���ҵ����ݣ��Ƿ���������Ϊ[" & .Text & "]�Ŀⷿ��λ�� ", vbYesNo + vbQuestion, Me.Caption) = vbNo Then
                            .Text = ""
                            .TxtSetFocus
                            Cancel = True
                            Exit Sub
                        Else
                            .AllowAddRow = False
                            '��λ_IN    IN ���Ͽⷿ��λ.����%Type
                            gstrSQL = "zl_���Ŀⷿ��λ_Update('" & strKey & "')"
                            Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            If .Cols - 1 = col��λ And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                                Exit Sub
                            End If
                        End If
                    Else
                        .AllowAddRow = False
                        If rsTemp.RecordCount = 1 Then
                            .TextMatrix(.Row, col��λ) = rsTemp.Fields("����")
                            .Text = rsTemp.Fields("����")
                            If .Cols - 1 = col��λ And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                                Exit Sub
                            End If
                        Else
                            Set mshSelect.Recordset = rsTemp
                            Call setSelectLocal
                            Cancel = True
                            If .Cols - 1 = col��λ And .Row = .Rows - 1 Then
                                cmdSave.SetFocus
                            End If
                            Exit Sub
                        End If
                    End If
                End If
                OS.OpenIme False
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlLimitRef()
    '--------------------------------------------------------
    '���ܣ�ˢ�¿���޶�
    '--------------------------------------------------------
    err = 0: On Error GoTo ErrHand
    gstrSQL = "Select  i.Id, i.����, i.����, i.���, i.����, i.��װ��λ As ��λ, i.����ϵ�� As ��װ," & vbNewLine & _
                    "       Nvl(l.����, 0) / i.����ϵ�� As ����,Nvl(l.����, 0) /i.����ϵ�� As ����, l.�̵�����, l.�ⷿ��λ," & vbNewLine & _
                    "       nvl(k.ʵ������,0)/i.����ϵ�� as �������," & vbNewLine & _
                    "       Decode(nvl(k.ʵ������,0),0,i.�ɱ���,(k.ʵ�ʽ��-k.ʵ�ʲ��) / k.ʵ������)*i.����ϵ�� as �ɱ���," & vbNewLine & _
                    "       Decode(i.�Ƿ���, 0, p.�ּ�,Decode(nvl(k.ʵ������,0),0,nvl(i.�ϴ��ۼ�,p.�ּ�),k.ʵ�ʽ�� / k.ʵ������))*i.����ϵ�� as ���ۼ�" & vbNewLine & _
                    "From (Select i.�Ƿ���, i.Id, i.����, i.����, i.���, i.����, i.���㵥λ, s.��װ��λ," & vbNewLine & _
                    "            Decode(s.����ϵ��, 0, 1, Null, 1, s.����ϵ��) as ����ϵ��,s.�ɱ���,s.�ϴ��ۼ�" & vbNewLine & _
                    "       From �շ���ĿĿ¼ I, �������� S, (Select Distinct ������Ŀid From ����ִ�п��� Where ִ�п���id =[1]) E" & vbNewLine & _
                    "       Where i.Id = s.����id And s.����id = e.������Ŀid And i.��� = '4' And" & vbNewLine & _
                    "             (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))) I," & vbNewLine & _
                    "     (Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� L Where �ⷿid =[1]) L," & vbNewLine & _
                    "     (Select ҩƷid, Sum(ʵ������) As ʵ������, Sum(ʵ�ʽ��) As ʵ�ʽ��, Sum(ʵ�ʲ��) As ʵ�ʲ��" & vbNewLine & _
                    "       From ҩƷ���" & vbNewLine & _
                    "       Where ���� = 1 And �ⷿid =[1]" & vbNewLine & _
                    "       Group By ҩƷid) K, �շѼ�Ŀ P" & vbNewLine & _
                    "Where i.Id = p.�շ�ϸĿid And i.Id = l.����id(+) And i.Id = k.ҩƷid(+) And" & vbNewLine & _
                    "      (p.��ֹ���� Is Null Or Sysdate Between p.ִ������ And Nvl(p.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) And" & vbNewLine & _
                    "      p.�۸�ȼ� Is Null" & vbNewLine & _
                    "Order By i.����"

    Set mrsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.cboRoom.ItemData(Me.cboRoom.ListIndex))
        
    With mrsTemp
        Me.msfLimit.ClearBill
        Me.msfLimit.Redraw = False
        Do While Not .EOF
            If Me.msfLimit.Rows < .AbsolutePosition + 1 Then Me.msfLimit.Rows = Me.msfLimit.Rows + 1
            Me.msfLimit.RowData(.AbsolutePosition) = !Id
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = !����
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = !����
            Me.msfLimit.TextMatrix(.AbsolutePosition, col���) = IIf(IsNull(!���), "", !���)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = IIf(IsNull(!����), "", !����)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col��λ) = IIf(IsNull(!��λ), "", !��λ)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col��װ) = zlStr.Nvl(!��װ)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = Format(!����, mFMT.FM_����)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = Format(!����, mFMT.FM_����)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = IIf(Mid(!�̵�����, 1, 1) = "1", "��", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = IIf(Mid(!�̵�����, 2, 1) = "1", "��", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = IIf(Mid(!�̵�����, 3, 1) = "1", "��", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col����) = IIf(Mid(!�̵�����, 4, 1) = "1", "��", "")
            Me.msfLimit.TextMatrix(.AbsolutePosition, col��λ) = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col�ɱ���) = Format(!�ɱ���, mFMT.FM_�ɱ���)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col���ۼ�) = Format(!���ۼ�, mFMT.FM_���ۼ�)
            Me.msfLimit.TextMatrix(.AbsolutePosition, col�������) = Format(!�������, mFMT.FM_����)
            
            If .AbsolutePosition Mod IIf(.RecordCount > 20, .RecordCount \ 20, 1) = 0 Then
                Me.stbThis.Panels(2).Text = "������ȡ��" & String(.AbsolutePosition \ IIf(.RecordCount > 20, .RecordCount \ 20, 1), "��")
            End If
            .MoveNext
        Loop
        Me.msfLimit.Redraw = True
    End With
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt����_Change()
    mblnFindFrist = False
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String, lngStart As Long, lngRows As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim strTemp���� As String, strTemp���� As String, strTemp���� As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txt����.Text))
    If strInput = "" Then Exit Sub
    
    '����ҩƷ
    If strInput = txt����.Tag Then
        '��ʾ������һ����¼
        If mlngFind >= msfLimit.Rows - 1 And mblnFindFrist = True Then
            MsgBox "�Ѳ�ѯ�����", vbInformation, gstrSysName
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '��ʾ�µĲ���
        lngStart = 0
        txt����.Tag = strInput
    End If
    
    '��ʼ����
    lngStart = lngStart + 1
    lngRows = msfLimit.Rows - 1
    mblnFind = False
    For lngStart = lngStart To lngRows
        str���� = Trim(UCase(msfLimit.TextMatrix(lngStart, col����)))
        str���� = Trim(UCase(msfLimit.TextMatrix(lngStart, col����)))
        str���� = UCase(zlStr.GetCodeByVB(str����))
        If str���� Like "*" & strInput & "*" Or _
            str���� Like "*" & strInput & "*" Or _
            str���� Like "*" & strInput & "*" Then
            msfLimit.Row = lngStart
            msfLimit.MsfObj.TopRow = lngStart
            msfLimit.SetFocus
            mblnFind = True '��¼�Ѿ���ѯ��ֵ
            mblnFindFrist = True
            Exit For
        End If
    Next

    mlngFind = lngStart
    If mlngFind = lngRows + 1 And mblnFind = False And mblnFindFrist = False Then
        MsgBox "û���ҵ�����Ҫ�����ݣ�", vbInformation, gstrSysName
        zlControl.TxtSelAll txt����
        Exit Sub
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'Private Sub cmdFind_Click()
'    Dim blnVisible As Boolean
'    '���һ������һ��
'    blnVisible = lbl����.Visible Xor True
'    lbl����.Visible = blnVisible
'    txt����.Visible = blnVisible
'    If blnVisible Then txt����.SetFocus
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If txt����.Visible And KeyCode = vbKeyF3 Then
        Call txt����_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub msfLimit_CommandClick()
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ����,����,���� From ���Ͽⷿ��λ Order by ����"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ���Ͽⷿ��λ")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "���Ͽⷿ��λ��δ��ʼ����[�ֵ����]", vbInformation, gstrSysName
        Exit Sub
    End If
    With msfLimit
        If rsTemp.RecordCount = 1 Then
            .TextMatrix(.Row, col��λ) = rsTemp.Fields("����")
            .Text = rsTemp.Fields("����")
        Else
            Set mshSelect.Recordset = rsTemp
            Call setSelectLocal
            Exit Sub
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    With msfLimit
        If KeyCode = vbKeyEscape Then
            mshSelect.Visible = False
            .SetFocus
        End If
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = mshSelect.TextMatrix(mshSelect.Row, 1)
            mshSelect.Visible = False
            .Col = col��λ
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            End If
            .SetFocus
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    If mshSelect.Visible Then
        mshSelect.Visible = False
    End If
End Sub
Private Sub setSelectLocal()
    '����:����ѡ������λ��
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngTemp As Single
    sngLeft = msfLimit.Left + msfLimit.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = msfLimit.Top + msfLimit.MsfObj.CellTop + msfLimit.MsfObj.CellHeight
    With mshSelect
        .Redraw = False
        If sngLeft + .Width > Me.ScaleWidth Then
            If sngLeft - .Width < 0 Then
                sngLeft = 0
            Else
                sngLeft = sngLeft - .Width + msfLimit.MsfObj.CellWidth
            End If
        End If
        sngTemp = sngTop - msfLimit.MsfObj.CellHeight
        
        If Me.ScaleHeight - sngTop > sngTemp Then
            .Height = Me.ScaleHeight - sngTop
        Else
            .Height = IIf(sngTemp < 0, 0, sngTemp)
            sngTemp = sngTop
            sngTop = sngTop - .Height
            
        End If
        If .Rows * .RowHeight(0) + (.Rows * 15) + .RowHeight(0) <= .Height Then
            .Height = .Rows * .RowHeight(0) + (.Rows * 15) + .RowHeight(0)
            If msfLimit.Top + msfLimit.MsfObj.CellTop > sngTop Then
                sngTop = msfLimit.Top + msfLimit.MsfObj.CellTop - .Height
            End If
        End If
        
        .Left = sngLeft
        .Top = sngTop
        .ColWidth(0) = 1000
        .ColWidth(1) = IIf(.Width - .ColWidth(0) - 15 < 0, 500, .Width - .ColWidth(0) - 15)
        .Row = 1
        .Col = 0
        .TopRow = 1
        .ColSel = .Cols - 1
        .Visible = True
        .SetFocus
        .Redraw = True
        Exit Sub
    End With
End Sub
