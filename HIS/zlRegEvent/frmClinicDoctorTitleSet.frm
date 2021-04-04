VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicDoctorTitleSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ְ�Ʊ�ʶ����"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   Icon            =   "frmClinicDoctorTitleSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      Top             =   3060
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Top             =   270
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDoctorTitle 
      Height          =   2985
      Left            =   90
      TabIndex        =   3
      Top             =   510
      Width           =   3975
      _cx             =   7011
      _cy             =   5265
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
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicDoctorTitleSet.frx":000C
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
   Begin VB.Label lbl��ʾ 
      Caption         =   "    �����ø�ְ���ڹҺŰ�����ҽ������ǰ��ʾ�ı�ʶ����"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   4110
   End
End
Attribute VB_Name = "frmClinicDoctorTitleSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnChange As Boolean     '�Ƿ�ı���

Public Function ShowMe(frmParent As Form) As Boolean
    '�������
    mblnOk = False
    On Error Resume Next
    Me.Show 1, frmParent
 
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function TextIsValied(strText As String) As Boolean
    '�����ı��Ƿ���Ч
    Dim intCHeckLen As Integer
    
    intCHeckLen = 5
    If zlCommFun.StrIsValid(strText, intCHeckLen) = False Then Exit Function
    If InStr(strText, ",") > 0 Or InStr(strText, ";") Then
        MsgBox "��ʶ�����зǷ��ַ� ", vbInformation, gstrSysName
        Exit Function
    End If
    TextIsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim i As Integer
    
    On Error GoTo ErrHandler
    '���ݼ��
    With vsfDoctorTitle
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("��ʶ��"))) <> "" Then
                If TextIsValied(Trim(.TextMatrix(i, .ColIndex("��ʶ��")))) = False Then
                    .Row = i: Exit Sub
                End If
            End If
        Next
    End With
    
    If SaveData() = False Then Exit Sub
    
    mblnOk = True
    mblnChange = False
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData() As Boolean
    '����:����༭������
    '����:
    '����ֵ:�ɹ�����True,����ΪFalse
    Dim i As Integer, str��ʶ�� As String
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    '���������õı�ʶ������һ����
    '��ʽ������1,��ʶ��1;����2,��ʶ��2;...
    With vsfDoctorTitle
        For i = .FixedRows To .Rows - 1
            str��ʶ�� = str��ʶ�� & ";"
            str��ʶ�� = str��ʶ�� & .TextMatrix(i, .ColIndex("����")) & ","
            str��ʶ�� = str��ʶ�� & Trim(.TextMatrix(i, .ColIndex("��ʶ��")))
        Next
    End With
    If str��ʶ�� <> "" Then str��ʶ�� = Mid(str��ʶ��, 2)
    
    'Zl_רҵ����ְ��_���±�ʶ��
    strSQL = "Zl_רҵ����ְ��_���±�ʶ��("
    '    ��ʶ��_In In Varchar2 --��ʽ������1,��ʶ��1;����2,��ʶ��2;...
    strSQL = strSQL & "'" & str��ʶ�� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    mblnChange = False
    
    On Error GoTo ErrHandler
    '23* ����������Ա��ҽ�ƣ�
    strSQL = "Select ����, ����, ��ʶ��" & vbNewLine & _
            " From רҵ����ְ��" & vbNewLine & _
            " Where ���� Like '23%' And ���� <> '23'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then
        MsgBox "    �ڡ�רҵ����ְ�񡱱���û���ҵ������ԡ�23����ͷ��ʾ������������Ա��ҽ�ƣ��������ݣ����������Ƿ�������", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    With vsfDoctorTitle
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        .Editable = flexEDKbdMouse
        .GridLines = flexGridInset
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("ְ��")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��ʶ��")) = Nvl(rsTemp!��ʶ��)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        
        .Cell(flexcpBackColor, .FixedRows, .ColIndex("ְ��"), .Rows - 1, .ColIndex("ְ��")) = vbButtonFace
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
End Sub

Private Sub vsfDoctorTitle_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsfDoctorTitle_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDoctorTitle.ColIndex("ְ��") Then Cancel = True
End Sub

Private Sub vsfDoctorTitle_EnterCell()
    vsfDoctorTitle.EditCell
End Sub

Private Sub vsfDoctorTitle_GotFocus()
    With vsfDoctorTitle
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub

Private Sub vsfDoctorTitle_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If TextIsValied(Trim(vsfDoctorTitle.EditText)) = False Then
        Cancel = True
    End If
End Sub
