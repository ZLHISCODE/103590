VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBalanceDefaultSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ȱʡ���㷽ʽ����"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   Icon            =   "frmBalanceDefaultSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4290
      TabIndex        =   4
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   3
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4290
      TabIndex        =   2
      Top             =   180
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf���ʽ 
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      Top             =   630
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   260
      RowHeightMax    =   260
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceDefaultSet.frx":000C
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
      Caption         =   "    �����ø�ҽ�Ƹ��ʽ�ڸó��ϵ�ȱʡ���㷽ʽ��"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4110
   End
End
Attribute VB_Name = "frmBalanceDefaultSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr���� As String
Dim mblnItem As Boolean
Dim mintSuccess As Integer
Dim mblnChange As Boolean     '�Ƿ�ı���

Public Function ShowMe(frmParent As Object, ByVal str���� As String) As Boolean
'����:��������õĽ��㷽ʽ�����ڽ���ͨѶ�ĳ���
'����:str����     ��ǰ�༭�Ľ��㷽ʽ�ı���
'����ֵ:�༭�ɹ�����True,����ΪFalse
    On Error GoTo ErrHandler
    
    mstr���� = str����
    
    Me.Show vbModal, frmParent
    ShowMe = mintSuccess > 0
    
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, 5
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
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
    '����:����༭�����ݵ����㷽ʽ����
    '����:
    '����ֵ:�ɹ�����True,����ΪFalse
    Dim i As Integer, strȱʡ As String
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    '������ȱʡ���㷽ʽ����һ����
    '��ʽ��ҽ�Ƹ��ʽ1:���㷽ʽ1;ҽ�Ƹ��ʽ2:���㷽ʽ2;...
    With vsf���ʽ
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ȱʡ���㷽ʽ"))) <> "" Then
                strȱʡ = strȱʡ & ";"
                strȱʡ = strȱʡ & Trim(.TextMatrix(i, .ColIndex("ҽ�Ƹ��ʽ"))) & ":"
                strȱʡ = strȱʡ & Trim(.TextMatrix(i, .ColIndex("ȱʡ���㷽ʽ")))
            End If
        Next
    End With
    If strȱʡ <> "" Then strȱʡ = Mid(strȱʡ, 2)
    
    '�޸�
    strSQL = "zl_���㷽ʽӦ��_update( '" & mstr���� & "','',1,'" & strȱʡ & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim strSQL As String
    Dim rs���㷽ʽ As New ADODB.Recordset
    Dim rs���ʽ As New ADODB.Recordset, lngRow As Long
    Dim str���㷽ʽ As String
    
    mblnChange = False
    mintSuccess = 0
    
    lbl��ʾ.Caption = Replace(lbl��ʾ.Caption, "�ó���", "��" & mstr���� & "�����㳡��")
    
    '��ѡ���õĽ��㷽ʽ������Ϊ1,2,7,8������Ӧ�����
    strSQL = "Select B.���㷽ʽ" & _
            " From ���㷽ʽ A,���㷽ʽӦ�� B" & _
            " Where A.����=B.���㷽ʽ And B.Ӧ�ó���=[1] And b.���ʽ Is Null" & _
            "       And A.���� In(1,2,7,8) And nvl(A.Ӧ����,0)<>1" & _
            " Order by A.����"
    Set rs���㷽ʽ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����)
    If rs���㷽ʽ.EOF Then
        MsgBox "    ��" & mstr���� & "�����㳡��û�п���������Ϊȱʡ���㷽ʽ�Ľ��㷽ʽ��" & _
            "���ȶԡ�" & mstr���� & "�����㳡����������Ϊ(1,2,7,8)�����Ҳ���Ӧ����Ľ��㷽ʽ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Do Until rs���㷽ʽ.EOF
        str���㷽ʽ = str���㷽ʽ & "|" & Nvl(rs���㷽ʽ!���㷽ʽ)
        rs���㷽ʽ.MoveNext
    Loop
    vsf���ʽ.ColComboList(vsf���ʽ.ColIndex("ȱʡ���㷽ʽ")) = " " & str���㷽ʽ
    
    'ҽ�Ƹ��ʽ���Լ������õ�ȱʡ���㷽ʽ
    strSQL = "Select a.���� As ���ʽ, b.���㷽ʽ" & vbNewLine & _
            " From ҽ�Ƹ��ʽ A, ���㷽ʽӦ�� B" & vbNewLine & _
            " Where a.���� = b.���ʽ(+) And b.Ӧ�ó���(+) = [1]" & vbNewLine & _
            "       And b.���ʽ(+) Is Not Null" & vbNewLine & _
            " Order By a.����"
    Set rs���ʽ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����)
    
    With vsf���ʽ
        .Clear 1
        .Editable = flexEDKbdMouse
        
        .Rows = rs���ʽ.RecordCount + 1
        lngRow = 1
        Do Until rs���ʽ.EOF
            .TextMatrix(lngRow, .ColIndex("ҽ�Ƹ��ʽ")) = Nvl(rs���ʽ!���ʽ)
            .TextMatrix(lngRow, .ColIndex("ȱʡ���㷽ʽ")) = Nvl(rs���ʽ!���㷽ʽ)
            lngRow = lngRow + 1
            rs���ʽ.MoveNext
        Loop
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
End Sub

Private Sub vsf���ʽ_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsf���ʽ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsf���ʽ.ColIndex("ҽ�Ƹ��ʽ") Then Cancel = True
End Sub

Private Sub vsf���ʽ_GotFocus()
    With vsf���ʽ
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub


