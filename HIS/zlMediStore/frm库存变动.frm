VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���䶯 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���䶯"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frm���䶯.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8895
      TabIndex        =   2
      Top             =   0
      Width           =   8895
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "˵�����ô�����ʾ�̵���ж�ӦҩƷ���̵����ں����Ŀ��䶯�����"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   75
         Width           =   5760
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   0
         Picture         =   "frm���䶯.frx":000C
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   345
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf���䶯 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      _cx             =   15690
      _cy             =   8281
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
      BackColorSel    =   16769992
      ForeColorSel    =   0
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���䶯.frx":685E
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
      ExplorerBar     =   5
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
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "ע�⣺��ɫ��ʾ��������ҵ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   5235
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Caption         =   "˵�����ô�����ʾ�̵���ж�ӦҩƷ���̵����ں����Ŀ��䶯�����"
      Height          =   180
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "frm���䶯"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintStyle As Integer '1:��ʾ������ʾ���䶯��2����ʾ������ʾ��������ռ��
Private mlng�ⷿID As Long
Private mlngҩƷID As Long
Private mlng���� As Long
Private mstr�̵�ʱ�� As String
Private mstr��λ As String
Private mdbl����ϵ�� As Double
Private mstr��λС As String
Private mdbl����ϵ��С As Double
Private mbln���ִ�С��λ As Boolean
Private mintNumberDigit As Integer

Public Sub ShowME(ByVal intStyle As Integer, ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long, ByVal lng���� As Long, ByVal str�̵�ʱ�� As String, ByVal frmPar As Form, ParamArray arrInput() As Variant)
    Dim arrPars() As Variant
    arrPars = arrInput
    
    If UBound(arrPars) = 2 Then
        mbln���ִ�С��λ = False
        mstr��λ = arrPars(0)
        mdbl����ϵ�� = arrPars(1)
        mintNumberDigit = arrPars(2)
    Else
        mbln���ִ�С��λ = True
        mstr��λ = arrPars(0)
        mdbl����ϵ�� = arrPars(1)
        mstr��λС = arrPars(2)
        mdbl����ϵ��С = arrPars(3)
        mintNumberDigit = arrPars(4)
    End If
    
    mintStyle = intStyle
    mlng�ⷿID = lng�ⷿID
    mlngҩƷID = lngҩƷID
    mlng���� = lng����
    mstr�̵�ʱ�� = str�̵�ʱ��
    
    Me.Show 1, frmPar
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim int���װ���� As Integer
    
    On Error GoTo ErrHandle
    If mintStyle = 1 Then '���䶯
        gstrSQL = "Select * From (Select a.NO,Decode(a.����," & vbNewLine & _
                    "               1," & vbNewLine & _
                    "               '�⹺���'," & vbNewLine & _
                    "               2," & vbNewLine & _
                    "               '�������'," & vbNewLine & _
                    "               3," & vbNewLine & _
                    "               'Эҩ���'," & vbNewLine & _
                    "               4," & vbNewLine & _
                    "               '�������'," & vbNewLine & _
                    "               6," & vbNewLine & _
                    "               'ҩƷ�ƿ�'," & vbNewLine & _
                    "               7," & vbNewLine & _
                    "               '��������'," & vbNewLine & _
                    "               11," & vbNewLine & _
                    "               '��������'," & vbNewLine & _
                    "               12," & vbNewLine & _
                    "               'ҩƷ�̵�'," & vbNewLine & _
                    "               '������ҩ') As ҵ������,  a.���ϵ�� * a.ʵ������ * a.���� As ��������, To_Char(a.�������, 'yyyy-mm-dd HH24:Mi:SS') As ��������, a.������, a.�����,a.��¼״̬,a.��ҩ��ʽ,a.����" & vbNewLine & _
                    "From ҩƷ�շ���¼ a" & vbNewLine & _
                    "Where  a.�ⷿid = [1] And a.ҩƷid = [2] And nvl(a.����,0) = [3] And a.������� > To_Date([4], 'YYYY-MM-DD HH24:MI:SS') And a.���� not in (5,13)" & vbNewLine & _
                    "Order By a.������� Desc )" & vbNewLine & _
                    "union all " & vbNewLine & _
                    "Select '','�ϼ�' As ҵ������ , sum(a.���ϵ�� * a.ʵ������ * a.����) As ��������, '', '', '',Null,Null,Null" & vbNewLine & _
                    "From ҩƷ�շ���¼ a,�շ���ĿĿ¼ b" & vbNewLine & _
                    "Where a.ҩƷid = b.id and a.�ⷿid = [1] And a.ҩƷid = [2] And nvl(a.����,0) = [3] And a.������� > To_Date([4], 'YYYY-MM-DD HH24:MI:SS')And a.���� not in (5,13)"
    
    
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng�ⷿID, mlngҩƷID, mlng����, mstr�̵�ʱ��)
        
        If rsTemp.RecordCount <= 1 Then Exit Sub
        
        With vsf���䶯
            
            Do While Not rsTemp.EOF
                .rows = .rows + 1
                .Row = .Row + 1
                
                .TextMatrix(.Row, .ColIndex("NO")) = "" & rsTemp!NO
                .TextMatrix(.Row, .ColIndex("ҵ������")) = "" & rsTemp!ҵ������
                
                '��ʾ�������˿ⵥ����ʾ
                If rsTemp!���� = 1 Then '�⹺���
                    If rsTemp!��¼״̬ Mod 3 = 2 Then '��������
                        If rsTemp!��ҩ��ʽ = 1 Then '�˿�
                            .TextMatrix(.Row, .ColIndex("ҵ������")) = .TextMatrix(.Row, .ColIndex("ҵ������")) & "(�ˡ���)"
                        Else
                            .TextMatrix(.Row, .ColIndex("ҵ������")) = .TextMatrix(.Row, .ColIndex("ҵ������")) & "(����)"
                        End If
                    Else
                        If rsTemp!��ҩ��ʽ = 1 Then .TextMatrix(.Row, .ColIndex("ҵ������")) = .TextMatrix(.Row, .ColIndex("ҵ������")) & "(�˿�)" '�˿�
                    End If
                Else
                    If rsTemp!��¼״̬ Mod 3 = 2 Then .TextMatrix(.Row, .ColIndex("ҵ������")) = .TextMatrix(.Row, .ColIndex("ҵ������")) & "(����)" '��������
                End If
                
                '��ɫ�����롢����
                If rsTemp!�������� < 0 Then .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '�����ɫ
               
                
                If Not mbln���ִ�С��λ Then
                    .TextMatrix(.Row, .ColIndex("��������")) = zlStr.FormatEx(Abs(IIf(IsNull(rsTemp!��������), 0, rsTemp!��������)) / mdbl����ϵ��, mintNumberDigit, , True) & mstr��λ
                Else
'                    If rsTemp!�������� < 0 Then .TextMatrix(.Row, .ColIndex("��������")) = "-"
                    
                    int���װ���� = Int(Abs(IIf(IsNull(rsTemp!��������), 0, rsTemp!��������)) / mdbl����ϵ��)
                    .TextMatrix(.Row, .ColIndex("��������")) = .TextMatrix(.Row, .ColIndex("��������")) & zlStr.FormatEx(int���װ����, mintNumberDigit, , True) & mstr��λ
                    .TextMatrix(.Row, .ColIndex("��������")) = .TextMatrix(.Row, .ColIndex("��������")) & zlStr.FormatEx((Abs(IIf(IsNull(rsTemp!��������), 0, rsTemp!��������)) - int���װ���� * mdbl����ϵ��) / mdbl����ϵ��С, mintNumberDigit, , True) & mstr��λС
                End If
                
                .TextMatrix(.Row, .ColIndex("��������")) = "" & rsTemp!��������
                .TextMatrix(.Row, .ColIndex("������")) = "" & rsTemp!������
                .TextMatrix(.Row, .ColIndex("�����")) = "" & rsTemp!�����
                
                rsTemp.MoveNext
    
            Loop
             
            .TextMatrix(.rows - 1, .ColIndex("ҵ������")) = .TextMatrix(.rows - 1, .ColIndex("ҵ������")) & IIf(.Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF, "(����)", "(���)")
        End With
       
        
    Else '��������ռ��
        Me.Caption = "��������ռ��"
        lblComment.Caption = "˵�����ô�����ʾ�̵���ж�ӦҩƷ30��֮�ڵĿ�������ռ�������"
        gstrSQL = "Select * From (Select Decode(a.����," & vbNewLine & _
                    "               1," & vbNewLine & _
                    "               '�⹺���'," & vbNewLine & _
                    "               2," & vbNewLine & _
                    "               '�������'," & vbNewLine & _
                    "               3," & vbNewLine & _
                    "               'Эҩ���'," & vbNewLine & _
                    "               4," & vbNewLine & _
                    "               '�������'," & vbNewLine & _
                    "               6," & vbNewLine & _
                    "               'ҩƷ�ƿ�'," & vbNewLine & _
                    "               7," & vbNewLine & _
                    "               '��������'," & vbNewLine & _
                    "               11," & vbNewLine & _
                    "               '��������'," & vbNewLine & _
                    "               12," & vbNewLine & _
                    "               'ҩƷ�̵�'," & vbNewLine & _
                    "               '������ҩ') As ҵ������, a.ʵ������ * a.���� As ռ������, To_Char(a.��������, 'yyyy-mm-dd HH24:Mi:SS') As ռ������, a.������, a.�����" & vbNewLine & _
                    "From ҩƷ�շ���¼ a" & vbNewLine & _
                    "Where  a.���ϵ�� = -1 And a.�ⷿid = [1] And a.ҩƷid = [2] And nvl(a.����,0) = [3] And a.������� is null And a.�������� > (sysdate - 30) And a.���� not in (5,13)" & vbNewLine & _
                    "Order By a.�������� Desc )" & vbNewLine & _
                    "union all " & vbNewLine & _
                    "Select '�ϼ�' As ҵ������, sum(a.ʵ������ * a.����) As ��������, '', '', ''" & vbNewLine & _
                    "From ҩƷ�շ���¼ a,�շ���ĿĿ¼ b" & vbNewLine & _
                    "Where a.���ϵ�� = -1 And a.ҩƷid = b.id and a.�ⷿid = [1] And a.ҩƷid = [2] And nvl(a.����,0) = [3] And a.������� is null And a.�������� > (sysdate - 30)And a.���� not in (5,13)"
    
    
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng�ⷿID, mlngҩƷID, mlng����)
        
        If rsTemp.RecordCount <= 1 Then Exit Sub
        
        With vsf���䶯
            vsf���䶯.TextMatrix(0, 2) = "ռ������"
            vsf���䶯.TextMatrix(0, 4) = "ռ������"
            
            Do While Not rsTemp.EOF
                .rows = .rows + 1
                .Row = .Row + 1
                
                .TextMatrix(.Row, .ColIndex("ҵ������")) = rsTemp!ҵ������
                
                If Not mbln���ִ�С��λ Then
                    .TextMatrix(.Row, .ColIndex("��������")) = zlStr.FormatEx(IIf(IsNull(rsTemp!ռ������), 0, rsTemp!ռ������) / mdbl����ϵ��, mintNumberDigit, , True) & mstr��λ
                Else
                    If rsTemp!ռ������ < 0 Then .TextMatrix(.Row, .ColIndex("��������")) = "-"
                    
                    int���װ���� = Int(Abs(IIf(IsNull(rsTemp!ռ������), 0, rsTemp!ռ������)) / mdbl����ϵ��)
                    .TextMatrix(.Row, .ColIndex("��������")) = .TextMatrix(.Row, .ColIndex("��������")) & zlStr.FormatEx(int���װ����, mintNumberDigit, , True) & mstr��λ
                    .TextMatrix(.Row, .ColIndex("��������")) = .TextMatrix(.Row, .ColIndex("��������")) & zlStr.FormatEx((Abs(IIf(IsNull(rsTemp!ռ������), 0, rsTemp!ռ������)) - int���װ���� * mdbl����ϵ��) / mdbl����ϵ��С, mintNumberDigit, , True) & mstr��λС
                End If
                
                .TextMatrix(.Row, .ColIndex("��������")) = "" & rsTemp!ռ������
                .TextMatrix(.Row, .ColIndex("������")) = "" & rsTemp!������
                .TextMatrix(.Row, .ColIndex("�����")) = "" & rsTemp!�����
                
              rsTemp.MoveNext
    
            Loop
        End With
        
        
    End If
    
    vsf���䶯.Cell(flexcpFontBold, vsf���䶯.rows - 1, 0, vsf���䶯.rows - 1, vsf���䶯.Cols - 1) = True '���һ������Ӵ֣��ϼ��У�
    vsf���䶯.TopRow = vsf���䶯.Row
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

