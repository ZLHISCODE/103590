VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMediSpecExp 
   Caption         =   "�����չ��Ϣ����"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   Icon            =   "frmMediSpecExp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5040
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2555
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   2
      Top             =   4560
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11850
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   3500
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   4870
      _cx             =   8590
      _cy             =   6174
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMediSpecExp.frx":6852
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
      Editable        =   2
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
   Begin VB.Label lblComment 
      Caption         =   "����ɹ���"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4638
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   �����һ�е���Ŀ�����а��س�����������Ŀ����ѡ���а���Del����ɾ��������Ŀ"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   750
      TabIndex        =   0
      Top             =   173
      Width           =   4125
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmMediSpecExp.frx":68EB
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMediSpecExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintItemNameLength As Integer









Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strItems As String
    Dim n As Integer
    Dim i As Integer
    
    With vsfItem
        For n = 1 To .Rows - 1
            '��������Ƿ�Ϊ��
            If Trim(.TextMatrix(n, .ColIndex("��Ŀ����"))) = "" Then
                MsgBox "��Ŀ���Ʋ���Ϊ�գ���¼�����ƣ�", vbInformation, gstrSysName
                
                .Row = n
                .TopRow = n
                .Col = .ColIndex("��Ŀ����")
                Exit Sub
            End If
            
            '��������Ƿ񳬳�
            If LenB(StrConv(Trim(.TextMatrix(n, .ColIndex("��Ŀ����"))), vbFromUnicode)) > mintItemNameLength Then
                MsgBox "��Ŀ���Ƴ��������" & mintItemNameLength & "���ַ���" & Int(mintItemNameLength / 2) & "������ ��", vbInformation, gstrSysName
                
                .Row = n
                .TopRow = n
                .Col = .ColIndex("��Ŀ����")
                Exit Sub
            End If
            
            '��������Ƿ��ظ�
            For i = 1 To .Rows - 1
                If i <> n And Trim(.TextMatrix(i, .ColIndex("��Ŀ����"))) = Trim(.TextMatrix(n, .ColIndex("��Ŀ����"))) Then
                    MsgBox "��Ŀ�����Ѵ��ڣ�������¼�����ƣ�", vbInformation, gstrSysName
                    
                    .Row = n
                    .TopRow = n
                    .Col = .ColIndex("��Ŀ����")
                    Exit Sub
                End If
            Next
            
            'ƴ����Ŀ��
            If .TextMatrix(n, .ColIndex("��Ŀ����")) <> "" And .TextMatrix(n, .ColIndex("��Ŀ����")) <> .TextMatrix(n, .ColIndex("ԭ��Ŀ����")) Then
                strItems = IIf(strItems = "", "", strItems & "|") & .TextMatrix(n, .ColIndex("����")) & "," & .TextMatrix(n, .ColIndex("��Ŀ����"))
            End If
        Next
        
        If strItems <> "" Then
            gstrSql = "Zl_ҩƷ�����չ��Ŀ_Update("
            '��Ŀ��
            gstrSql = gstrSql & "'" & strItems & "'"
            gstrSql = gstrSql & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSql, "����ҩƷ�����չ��Ŀ")
            
            lblComment.Caption = "����ɹ���"
            lblComment.Visible = True
        End If
        
        For n = 1 To .Rows - 1
            .TextMatrix(n, .ColIndex("ԭ��Ŀ����")) = .TextMatrix(n, .ColIndex("��Ŀ����"))
        Next
        
        .Cell(flexcpForeColor, 1, .ColIndex("��Ŀ����"), .Rows - 1, .ColIndex("��Ŀ����")) = vbBlack
    End With
End Sub

Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    gstrSql = "Select ����, ���� From ҩƷ�����չ��Ŀ Order By ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "��ѯҩƷ�����չ��Ŀ")
    
    mintItemNameLength = rsData.Fields("����").DefinedSize
    
    With vsfItem
        .Rows = 1
        
        If rsData.RecordCount = 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����")) = "1"
            Exit Sub
        End If
        
        Do While Not rsData.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData!����
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ����")) = rsData!����
            .TextMatrix(.Rows - 1, .ColIndex("ԭ��Ŀ����")) = rsData!����
            
            rsData.MoveNext
        Loop
        
    End With
        
End Sub


Private Sub vsfItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 0 Then Exit Sub
    With vsfItem
        If Col <> .ColIndex("��Ŀ����") Then Exit Sub
        If .TextMatrix(Row, .ColIndex("��Ŀ����")) <> .TextMatrix(Row, .ColIndex("ԭ��Ŀ����")) Then
            .Cell(flexcpForeColor, Row, .ColIndex("��Ŀ����")) = vbRed
            lblComment.Visible = False
        Else
            .Cell(flexcpForeColor, Row, .ColIndex("��Ŀ����")) = vbBlack
        End If
    End With
End Sub

Private Sub vsfItem_EnterCell()
    With vsfItem
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .Col = .ColIndex("��Ŀ����") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub vsfItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsData As ADODB.Recordset
    
    With vsfItem
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("��Ŀ����") Then
                .Col = .Col + 1
            ElseIf .Row = .Rows - 1 And .Col = .ColIndex("��Ŀ����") And .TextMatrix(.Row, .ColIndex("��Ŀ����")) <> "" Then
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, .ColIndex("����")) = Val(.TextMatrix(.Rows - 2, .ColIndex("����"))) + 1
                
                .Row = .Rows - 1
                .Col = .ColIndex("��Ŀ����")
                lblComment.Visible = False
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, .ColIndex("ԭ��Ŀ����")) <> "" Then
                gstrSql = "Select 1 From ҩƷ�����չ��Ϣ Where ��Ŀ = [1] And Rownum < 2 "
                Set rsData = zlDatabase.OpenSQLRecord(gstrSql, "��ѯҩƷ�����չ��Ϣ", .TextMatrix(.Row, .ColIndex("��Ŀ����")))
                
                If rsData.RecordCount > 0 Then
                    If MsgBox("����ҩƷ��������չ��Ŀ��" & .TextMatrix(.Row, .ColIndex("��Ŀ����")) & "�����Ƿ�ɾ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Else
                    If MsgBox("�Ƿ�ɾ����չ��Ŀ��" & .TextMatrix(.Row, .ColIndex("��Ŀ����")) & "����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                gstrSql = "Zl_ҩƷ�����չ��Ŀ_Del("
                '����
                gstrSql = gstrSql & Val(.TextMatrix(.Row, .ColIndex("����")))
                gstrSql = gstrSql & ")"
        
                Call zlDatabase.ExecuteProcedure(gstrSql, "ɾ����չ��Ŀ")
            End If
            
            If .Row = .Rows - 1 Then
                .TextMatrix(.Row, .ColIndex("��Ŀ����")) = ""
            Else
                .RemoveItem .Row
            End If
            
            lblComment.Visible = False
        End If
    End With
End Sub


Private Sub vsfItem_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then Exit Sub
    If Col = vsfItem.ColIndex("��Ŀ����") Then
        If InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub


