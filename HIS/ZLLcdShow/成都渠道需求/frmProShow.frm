VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmProShow 
   BorderStyle     =   0  'None
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timRest 
      Interval        =   8000
      Left            =   7560
      Top             =   6480
   End
   Begin VB.Timer timData 
      Interval        =   1000
      Left            =   8520
      Top             =   6480
   End
   Begin VB.Timer timTime 
      Interval        =   60000
      Left            =   9600
      Top             =   6480
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf��ȡҩ 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   5385
      _cx             =   9499
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf�ѹ��� 
      Height          =   4455
      Left            =   7080
      TabIndex        =   2
      Top             =   1560
      Width           =   4185
      _cx             =   7382
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmProShow.frx":0152
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Label lbl��ܰ��ʾ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ܰ��ʾ�������ĵȺ������Ŷӡ�ף�����տ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   5520
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2019-01-01 00:00"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Top             =   7080
      Width           =   1965
   End
   Begin VB.Label lbl����_���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "һ�Ŵ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   9
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lbl����_�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   240
   End
   Begin VB.Label lbl����_���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lbl����_�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   600
      Width           =   240
   End
   Begin VB.Label lbl����_ȡҩ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȡҩ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblҩ������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ҩ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image img���� 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1065
   End
End
Attribute VB_Name = "frmProShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngҩ��ID As Long          '��ǰҩ��ID
Private mstrWins As String          '��ǰ��ҩ����
Private mbln��ҩ As Boolean         '�Ƿ���ҩ
Private mbln��ҩȷ�� As Boolean     '�Ƿ���ҩȷ��
Private mbln������ʾ As Boolean     '���������Ƿ�������ʾ

Private Const CST_STR_REG As String = "����ģ��\ҩ���Ŷӽк�\Һ������Pro"

Private Type Type_para
    bln������ģʽ As Boolean          '��Ļ������ʾģʽ��True-��������ʾ��False-�ര����ʾ
    str�ര�� As String               '�ര��ģʽ�������õĸ�����
    
    str����ͼƬ��ַ As String         '���ر���ͼƬ���ʵ�ַ
    
    '������λ��
    lngLeft As Long                  '���������ʾ�����λ��
    lngTop  As Long                 '���������ʾ���ϱ�λ��
    lngWidth As Long                '��������
    lngHeight As Long               '������߶�
    
    '����ˢ��
    int������ѯʱ�� As Integer
    int������ʾʱ�� As Integer
    
    'ҩ���������
    lngҩ��_Left As Long
    lngҩ��_Top As Long
    
    strҩ��_���� As String
    strҩ��_�ֺ� As String
    blnҩ��_���� As Boolean
    blnҩ��_б�� As Boolean
    lngҩ��_��ɫ As Long
    
    '���������������
    lng����_Left As Long
    lng����_Top As Long
    
    str����ͨ��_���� As String
    str����ͨ��_�ֺ� As String
    bln����ͨ��_���� As Boolean
    bln����ͨ��_б�� As Boolean
    lng����ͨ��_��ɫ As Long
    
    bln���������������� As Boolean
    str��������_���� As String
    str��������_�ֺ� As String
    bln��������_���� As Boolean
    bln��������_б�� As Boolean
    lng��������_��ɫ As Long
    
    bln���д��ڵ������� As Boolean
    str���д���_���� As String
    str���д���_�ֺ� As String
    bln���д���_���� As Boolean
    bln���д���_б�� As Boolean
    lng���д���_��ɫ As Long
    
    '����ҩ�����������
    bln��ʾ����ҩ As Boolean
    
    lng����ҩ_Left As Long
    lng����ҩ_Top As Long
    
    str����ҩ_���� As String
    str����ҩ_�ֺ� As String
    bln����ҩ_���� As Boolean
    bln����ҩ_б�� As Boolean
    lng����ҩ_��ɫ As Long
    
    lng����ҩ_�п� As Long
    lng����ҩ_�и� As Long
    lng����ҩ_���� As Long
    
    '�ѹ��������������
    bln��ʾ�ѹ��� As Boolean
    
    lng�ѹ���_Left As Long
    lng�ѹ���_Top As Long
    
    str�ѹ���_���� As String
    str�ѹ���_�ֺ� As String
    bln�ѹ���_���� As Boolean
    bln�ѹ���_б�� As Boolean
    lng�ѹ���_��ɫ As Long
    
    lng�ѹ���_�п� As Long
    lng�ѹ���_�и� As Long
    lng�ѹ���_���� As Long
    
    '��ܰ��ʾ����
    bln��ʾ��ʾ As Boolean
    lng��ܰ��ʾ_Left As Long
    lng��ܰ��ʾ_Top As Long
    
    str��ܰ��ʾ_���� As String             '��ʾ���·�����ܰ��ʾ���ݣ�Ĭ��ָ�·�����
    
    str��ܰ��ʾ_���� As String
    str��ܰ��ʾ_�ֺ� As String
    bln��ܰ��ʾ_���� As Boolean
    bln��ܰ��ʾ_б�� As Boolean
    lng��ܰ��ʾ_��ɫ As Long
    
    'ʱ��
    bln��ʾʱ�� As Boolean
    
    lngʱ��_Left As Long
    lngʱ��_Top As Long
    
    strʱ��_���� As String
    strʱ��_�ֺ� As String
    blnʱ��_���� As Boolean
    blnʱ��_б�� As Boolean
    lngʱ��_��ɫ As Long
End Type

Private mPara As Type_para

Public Sub ShowMe(ByVal lngҩ��ID As Long, ByVal strWins As String, ByVal bln��ҩ As Boolean, ByVal bln��ҩȷ�� As Boolean)
    '���ܣ���ģ����г�ʼ��
    
    mlngҩ��ID = lngҩ��ID
    mstrWins = strWins
    mbln��ҩ = bln��ҩ
    mbln��ҩȷ�� = bln��ҩȷ��
    
    '������ʾ
    Me.Show
End Sub

Private Sub Change��������()
    '���ܣ�������ʾ���ĺ�������
    Dim strSQL As String
    Dim strWins As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date

    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbln������ʾ Then Exit Sub
    
    strWins = IIf(mPara.bln������ģʽ, mstrWins, mPara.str�ര��)
    
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select ����, ��ҩ����, NO, ����, �ⷿid" & vbNewLine & _
            "From (Select ����, ��ҩ����, NO, ����, �ⷿid" & vbNewLine & _
            "       From δ��ҩƷ��¼" & vbNewLine & _
            "       Where (�Ŷ�״̬ = 3 Or �Ŷ�״̬ = 4) And Nvl(��ʾ״̬, 0) = 0 And �ⷿid = [1]" & vbNewLine & _
            "       And ��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And �������� Between [3] And [4]" & vbNewLine & _
            "       Order By ����ʱ��) A" & vbNewLine & _
            "Where Rownum < 2"
        
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "��������������", mlngҩ��ID, strWins, date��ʼ����, date��������)
    
    If rsData.EOF Then Exit Sub
    
    lbl����_����.Caption = rsData!����
    lbl����_����.Caption = rsData!��ҩ����
    
    mbln������ʾ = True
    
    timRest.Enabled = False
    timRest.Enabled = True
    
    Call Init��������
    Call Show��������(True)
    
    '�����������������ݣ�����ˢ����ʾ�Ĵ������
    strSQL = "Zl_δ��ҩƷ��¼_��ʾ("
    'NO
    strSQL = strSQL & "'" & rsData!NO & "'"
    '����
    strSQL = strSQL & "," & rsData!����
    'ҩ��id
    strSQL = strSQL & "," & rsData!�ⷿid
    strSQL = strSQL & ")"
    
    Call gobjDatabase.ExecuteProcedure(strSQL, "tmrCall_Timer")
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub Reset()
    '���ܣ����ý��沼��
    
    Call LoadPar
    Call ResetLayout
    Call DrawWallPaper
End Sub

Private Sub LoadPar()
    '���ܣ���ȡ��������
    
    With mPara
        .bln������ģʽ = (Val(GetSetting("ZLSOFT", CST_STR_REG, "����ģʽ", "0")) = 0)
        .str�ര�� = GetSetting("ZLSOFT", CST_STR_REG, "�ര��", "")
        .str����ͼƬ��ַ = GetSetting("ZLSOFT", CST_STR_REG, "ͼƬλ��", "")
        
        '������λ��
        .lngLeft = GetSetting("ZLSOFT", CST_STR_REG, "Һ����_��", "1024")
        .lngTop = GetSetting("ZLSOFT", CST_STR_REG, "Һ����_��", "0")
        .lngWidth = GetSetting("ZLSOFT", CST_STR_REG, "Һ����_���", "1024")
        .lngHeight = GetSetting("ZLSOFT", CST_STR_REG, "Һ����_�߶�", "768")
        
        '����ˢ��
        .int������ѯʱ�� = Val(GetSetting("ZLSOFT", CST_STR_REG, "������ѯʱ��", "1"))
        If .int������ѯʱ�� < 1 Then
            .int������ѯʱ�� = 1
        ElseIf .int������ѯʱ�� > 60 Then
            .int������ѯʱ�� = 60
        End If
        
        '��ʾˢ��
        .int������ʾʱ�� = Val(GetSetting("ZLSOFT", CST_STR_REG, "������ʾʱ��", "1"))
        If .int������ʾʱ�� < 1 Then
            .int������ʾʱ�� = 1
        ElseIf .int������ʾʱ�� > 60 Then
            .int������ʾʱ�� = 60
        End If
        
        'ҩ���������
        .lngҩ��_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "ҩ��_��", "0"))
        .lngҩ��_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "ҩ��_��", "0"))
        
        .strҩ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "ҩ������", "΢���ź�")
        .strҩ��_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "ҩ���ֺ�", "12")
        .blnҩ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "ҩ������", "False")
        .blnҩ��_б�� = GetSetting("ZLSOFT", CST_STR_REG, "ҩ��б��", "False")
        .lngҩ��_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "ҩ����ɫ", vbBlack)
        
        '���������������
        .lng����_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "����_��", "0"))
        .lng����_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "����_��", "0"))
        
        .str����ͨ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "��������_ͨ��", "΢���ź�")
        .str����ͨ��_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "����ͨ������_�ֺ�", "12")
        .bln����ͨ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "����ͨ������_����", "False")
        .bln����ͨ��_б�� = GetSetting("ZLSOFT", CST_STR_REG, "����ͨ������_б��", "False")
        .lng����ͨ��_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "������ɫ_ͨ��", vbBlack)
        
        .bln���������������� = (Val(GetSetting("ZLSOFT", CST_STR_REG, "����������������", "0")) = 1)
        .str��������_���� = GetSetting("ZLSOFT", CST_STR_REG, "��������_����", "΢���ź�")
        .str��������_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "������������_�ֺ�", "12")
        .bln��������_���� = GetSetting("ZLSOFT", CST_STR_REG, "������������_����", "False")
        .bln��������_б�� = GetSetting("ZLSOFT", CST_STR_REG, "������������_б��", "False")
        .lng��������_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "������ɫ_����", vbBlack)
        
        .bln���д��ڵ������� = (Val(GetSetting("ZLSOFT", CST_STR_REG, "���д��ڵ�������", "0")) = 1)
        .str���д���_���� = GetSetting("ZLSOFT", CST_STR_REG, "��������_����", "΢���ź�")
        .str���д���_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "���д�������_�ֺ�", "12")
        .bln���д���_���� = GetSetting("ZLSOFT", CST_STR_REG, "���д�������_����", "False")
        .bln���д���_б�� = GetSetting("ZLSOFT", CST_STR_REG, "���д�������_б��", "False")
        .lng���д���_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "������ɫ_����", vbBlack)
        
        '����ҩ�����������
        .bln��ʾ����ҩ = (Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾ����ҩ", "1")) = 1)
        
        .lng����ҩ_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "����ҩ_��", "0"))
        .lng����ҩ_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "����ҩ_��", "0"))
        
        .str����ҩ_���� = GetSetting("ZLSOFT", CST_STR_REG, "����ҩ����", "΢���ź�")
        .str����ҩ_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "����ҩ�ֺ�", "12")
        .bln����ҩ_���� = GetSetting("ZLSOFT", CST_STR_REG, "����ҩ����", "False")
        .bln����ҩ_б�� = GetSetting("ZLSOFT", CST_STR_REG, "����ҩб��", "False")
        .lng����ҩ_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "����ҩ��ɫ", vbBlack)
        
        .lng����ҩ_�п� = Val(GetSetting("ZLSOFT", CST_STR_REG, "����ҩ_�п�", "800"))
        .lng����ҩ_�и� = Val(GetSetting("ZLSOFT", CST_STR_REG, "����ҩ_�и�", "350"))
        .lng����ҩ_���� = Val(GetSetting("ZLSOFT", CST_STR_REG, "����ҩ_����", "5"))
        
        '�ѹ��������������
        .bln��ʾ�ѹ��� = (Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾ�ѹ���", "1")) = 1)
        
        .lng�ѹ���_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���_��", "0"))
        .lng�ѹ���_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���_��", "0"))
        
        .str�ѹ���_���� = GetSetting("ZLSOFT", CST_STR_REG, "�ѹ�������", "΢���ź�")
        .str�ѹ���_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "�ѹ����ֺ�", "12")
        .bln�ѹ���_���� = GetSetting("ZLSOFT", CST_STR_REG, "�ѹ��Ŵ���", "False")
        .bln�ѹ���_б�� = GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���б��", "False")
        .lng�ѹ���_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "�ѹ�����ɫ", vbBlack)
        
        .lng�ѹ���_�п� = Val(GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���_�п�", "800"))
        .lng�ѹ���_�и� = Val(GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���_�и�", "350"))
        .lng�ѹ���_���� = Val(GetSetting("ZLSOFT", CST_STR_REG, "�ѹ���_����", "5"))
        
        '��ܰ��ʾ����
        .bln��ʾ��ʾ = (Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾ��ʾ", "1")) = 1)
        .lng��ܰ��ʾ_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾ_��", "0"))
        .lng��ܰ��ʾ_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾ_��", "0"))
        
        .str��ܰ��ʾ_���� = GetSetting("ZLSOFT", CST_STR_REG, "��ʾ_����", "")
        
        .str��ܰ��ʾ_���� = GetSetting("ZLSOFT", CST_STR_REG, "��ʾ����", "΢���ź�")
        .str��ܰ��ʾ_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "��ʾ�ֺ�", "12")
        .bln��ܰ��ʾ_���� = GetSetting("ZLSOFT", CST_STR_REG, "��ʾ����", "False")
        .bln��ܰ��ʾ_б�� = GetSetting("ZLSOFT", CST_STR_REG, "��ʾб��", "False")
        .lng��ܰ��ʾ_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "��ʾ��ɫ", vbBlack)
        
        'ʱ��
        .bln��ʾʱ�� = (Val(GetSetting("ZLSOFT", CST_STR_REG, "��ʾʱ��", "1")) = 1)
        
        .lngʱ��_Left = Val(GetSetting("ZLSOFT", CST_STR_REG, "ʱ��_��", "0"))
        .lngʱ��_Top = Val(GetSetting("ZLSOFT", CST_STR_REG, "ʱ��_��", "0"))
        
        .strʱ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "ʱ������", "΢���ź�")
        .strʱ��_�ֺ� = GetSetting("ZLSOFT", CST_STR_REG, "ʱ���ֺ�", "12")
        .blnʱ��_���� = GetSetting("ZLSOFT", CST_STR_REG, "ʱ�����", "False")
        .blnʱ��_б�� = GetSetting("ZLSOFT", CST_STR_REG, "ʱ��б��", "False")
        .lngʱ��_��ɫ = GetSetting("ZLSOFT", CST_STR_REG, "ʱ����ɫ", vbBlack)
            
    End With
    
End Sub

Private Sub Init��������()
    '���ܣ����ú������ݵ�λ�ü���ʽ
    
    With lbl����_��
        .FontName = mPara.str����ͨ��_����
        .FontSize = mPara.str����ͨ��_�ֺ�
        .FontBold = mPara.bln����ͨ��_����
        .FontItalic = mPara.bln����ͨ��_б��
        .ForeColor = mPara.lng����ͨ��_��ɫ
        
        .Left = mPara.lng����_Left
        .Top = mPara.lng����_Top
    End With
    
    With lbl����_����
        .FontName = IIf(mPara.bln����������������, mPara.str��������_����, mPara.str����ͨ��_����)
        .FontSize = IIf(mPara.bln����������������, mPara.str��������_�ֺ�, mPara.str����ͨ��_�ֺ�)
        .FontBold = IIf(mPara.bln����������������, mPara.bln��������_����, mPara.bln����ͨ��_����)
        .FontItalic = IIf(mPara.bln����������������, mPara.bln��������_б��, mPara.bln����ͨ��_б��)
        .ForeColor = IIf(mPara.bln����������������, mPara.lng��������_��ɫ, mPara.lng����ͨ��_��ɫ)
        
        .Left = lbl����_��.Left + lbl����_��.Width + 50
        .Top = lbl����_��.Top + (lbl����_��.Height - .Height) / 2
    End With
    
    With lbl����_��
        .FontName = mPara.str����ͨ��_����
        .FontSize = mPara.str����ͨ��_�ֺ�
        .FontBold = mPara.bln����ͨ��_����
        .FontItalic = mPara.bln����ͨ��_б��
        .ForeColor = mPara.lng����ͨ��_��ɫ
        
        .Left = lbl����_����.Left + lbl����_����.Width + 50
        .Top = lbl����_��.Top
    End With
    
    With lbl����_����
        .FontName = IIf(mPara.bln���д��ڵ�������, mPara.str���д���_����, mPara.str����ͨ��_����)
        .FontSize = IIf(mPara.bln���д��ڵ�������, mPara.str���д���_�ֺ�, mPara.str����ͨ��_�ֺ�)
        .FontBold = IIf(mPara.bln���д��ڵ�������, mPara.bln���д���_����, mPara.bln����ͨ��_����)
        .FontItalic = IIf(mPara.bln���д��ڵ�������, mPara.bln���д���_б��, mPara.bln����ͨ��_б��)
        .ForeColor = IIf(mPara.bln���д��ڵ�������, mPara.lng���д���_��ɫ, mPara.lng����ͨ��_��ɫ)
        
        .Left = lbl����_��.Left + lbl����_��.Width + 50
        .Top = lbl����_��.Top + (lbl����_��.Height - .Height) / 2
    End With
    
    With lbl����_ȡҩ
        .FontName = mPara.str����ͨ��_����
        .FontSize = mPara.str����ͨ��_�ֺ�
        .FontBold = mPara.bln����ͨ��_����
        .FontItalic = mPara.bln����ͨ��_б��
        .ForeColor = mPara.lng����ͨ��_��ɫ
        
        .Left = lbl����_����.Left + lbl����_����.Width + 50
        .Top = lbl����_��.Top
    End With
End Sub

Private Sub ResetLayout()
    '����λ��
    With Me
        .Left = mPara.lngLeft * Screen.TwipsPerPixelY
        .Top = mPara.lngTop * Screen.TwipsPerPixelY
        .Width = mPara.lngWidth * Screen.TwipsPerPixelX
        .Height = mPara.lngHeight * Screen.TwipsPerPixelY
    End With
    
    '��������
    With img����
        .Top = 0
        .Left = 0
        .Height = Me.ScaleHeight
        .Width = Me.ScaleWidth
        
        .Picture = LoadPicture(mPara.str����ͼƬ��ַ)
    End With
    
    'ҩ���������
    With lblҩ������
        .Left = mPara.lngҩ��_Left
        .Top = mPara.lngҩ��_Top
        
        .FontName = mPara.strҩ��_����
        .FontSize = mPara.strҩ��_�ֺ�
        .FontBold = mPara.blnҩ��_����
        .FontItalic = mPara.blnҩ��_б��
        .ForeColor = mPara.lngҩ��_��ɫ
    End With
    
    '������ѯʱ��
    With timData
        .Interval = mPara.int������ѯʱ�� * 1000#
    End With
    
    '������ʾʱ��
    With timRest
        .Interval = mPara.int������ʾʱ�� * 1000#
    End With
    
    '���������������
    Call Init��������
    
    '����ҩ�����������
    Call InitList_����ҩ
    
    '�ѹ��������������
    Call InitList_�ѹ���
    
    '��ܰ��ʾ����
    With lbl��ܰ��ʾ
        .Visible = mPara.bln��ʾ��ʾ
        .Caption = mPara.str��ܰ��ʾ_����
        
        .FontName = mPara.str��ܰ��ʾ_����
        .FontSize = mPara.str��ܰ��ʾ_�ֺ�
        .FontBold = mPara.bln��ܰ��ʾ_����
        .FontItalic = mPara.bln��ܰ��ʾ_б��
        .ForeColor = mPara.lng��ܰ��ʾ_��ɫ
        
        .Left = mPara.lng��ܰ��ʾ_Left
        .Top = mPara.lng��ܰ��ʾ_Top
    End With
    
    'ʱ��
    With lblTime
        .Visible = mPara.bln��ʾʱ��
        
        .Caption = Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
        
        .FontName = mPara.strʱ��_����
        .FontSize = mPara.strʱ��_�ֺ�
        .FontBold = mPara.blnʱ��_����
        .FontItalic = mPara.blnʱ��_б��
        .ForeColor = mPara.lngʱ��_��ɫ
        
        .Left = mPara.lngʱ��_Left
        .Top = mPara.lngʱ��_Top
    End With
    
    With timTime
        .Enabled = mPara.bln��ʾʱ��
    End With
    
End Sub

Private Sub Form_Load()
    mbln������ʾ = False
    
    Call Show��������(False)
        
    '��ȡ��������
    Call LoadPar
    
    '���ò���
    Call ResetLayout
    
    '�ػ��񱳾�
    Call DrawWallPaper
End Sub

Private Sub Show��������(ByVal bln��ʾ As Boolean)
    '���ܣ��Ƿ���ʾ��������
    
    lbl����_��.Visible = bln��ʾ
    lbl����_����.Visible = bln��ʾ
    lbl����_��.Visible = bln��ʾ
    lbl����_����.Visible = bln��ʾ
    lbl����_ȡҩ.Visible = bln��ʾ
End Sub


Private Sub InitList_����ҩ()
    '���ܣ���ʼ������ҩ�б�
    Dim str���ڴ� As String
    Dim i As Integer
    
    str���ڴ� = IIf(mPara.bln������ģʽ, mstrWins, mPara.str�ര��)
    
    With vsf��ȡҩ
        .Visible = mPara.bln��ʾ����ҩ
        .Left = mPara.lng����ҩ_Left
        .Top = mPara.lng����ҩ_Top
        
        '�������
        .Clear
        
        '��������
        If mPara.bln������ģʽ Then
            .Rows = 1 + mPara.lng����ҩ_����
        Else
            .Rows = 4 + mPara.lng����ҩ_����
        End If
        
        '�����������
        .Cols = UBound(Split(str���ڴ�, ",")) + 1
        
        '������ÿ��Keyֵ
        For i = 0 To UBound(Split(str���ڴ�, ","))
            .ColKey(i) = Split(str���ڴ�, ",")(i)
            
            If mPara.bln������ģʽ Then
                '��ʾ���ȴ�ȡҩ��
                .TextMatrix(0, i) = "�ȴ�ȡҩ"
            Else
                '��ʾ����ҩ���ڡ�
                .TextMatrix(0, i) = Split(str���ڴ�, ",")(i)
                
                '��ʾ������ȡҩ��
                .TextMatrix(1, i) = "��ǰȡҩ"
                
                '��ʾ���ȴ�ȡҩ��
                .TextMatrix(3, i) = "�ȴ�ȡҩ"
            End If
        Next
        
        '������ʾ
        .ColAlignment(-1) = flexAlignCenterCenter
        
        '��������
        .FontName = mPara.str����ҩ_����
        .FontSize = mPara.str����ҩ_�ֺ�
        .FontBold = mPara.bln����ҩ_����
        .FontItalic = mPara.bln����ҩ_б��
        .ForeColor = mPara.lng����ҩ_��ɫ
        
        '��Ԫ������
        .ColWidth(-1) = mPara.lng����ҩ_�п�
        .RowHeight(-1) = mPara.lng����ҩ_�и�
        
        .Width = mPara.lng����ҩ_�п� * .Cols
        .Height = mPara.lng����ҩ_�и� * .Rows
        
    End With
    
    '���ݳ�ʼ��
    Call LoadData_����ҩ
    
End Sub

Private Sub InitList_�ѹ���()
    '���ܣ���ʼ������ҩ�б�
    Dim str���ڴ� As String
    Dim i As Integer
    
    str���ڴ� = IIf(mPara.bln������ģʽ, mstrWins, mPara.str�ര��)
    
    With vsf�ѹ���
        .Visible = mPara.bln��ʾ�ѹ���
        .Left = mPara.lng�ѹ���_Left
        .Top = mPara.lng�ѹ���_Top
        
        '�������
        .Clear
    
        '��������
        If mPara.bln������ģʽ Then
            .Rows = 1 + mPara.lng����ҩ_����
        Else
            .Rows = 2 + mPara.lng�ѹ���_����
        End If
        
        '�����������
        .Cols = UBound(Split(str���ڴ�, ",")) + 1
    
        '������ÿ��Keyֵ
        For i = 0 To UBound(Split(str���ڴ�, ","))
            .ColKey(i) = Split(str���ڴ�, ",")(i)
            
            '��ʾ������������
            .TextMatrix(0, i) = "��������"
            
            If Not mPara.bln������ģʽ Then
                '��ʾ����ҩ���ڡ�
                .TextMatrix(1, i) = Split(str���ڴ�, ",")(i)
            End If
            
        Next
        
        '������ʾ
        .ColAlignment(-1) = flexAlignCenterCenter
        
        '�ϲ�
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
        
        '��������
        .FontName = mPara.str�ѹ���_����
        .FontSize = mPara.str�ѹ���_�ֺ�
        .FontBold = mPara.bln�ѹ���_����
        .FontItalic = mPara.bln�ѹ���_б��
        .ForeColor = mPara.lng�ѹ���_��ɫ
        
        '��Ԫ������
        .ColWidth(-1) = mPara.lng�ѹ���_�п�
        .RowHeight(-1) = mPara.lng�ѹ���_�и�
        
        .Width = mPara.lng�ѹ���_�п� * .Cols
        .Height = mPara.lng�ѹ���_�и� * .Rows
        
    End With
    
    '���ݳ�ʼ��
    Call LoadData_�ѹ���
    
End Sub

Private Sub RefreshList_������(ByVal rsData As ADODB.Recordset)
    '���ܣ�ˢ�´���ҩ�б�
    Dim i As Integer
    Dim n As Integer
    Dim int�̶��� As Integer       'Ĭ�ϵڶ���Ϊ��������Ϣ
    
    int�̶��� = 2
    
    With vsf��ȡҩ
        For i = 0 To .Cols - 1
            rsData.Filter = "��ҩ���� = '" & .ColKey(i) & "'"
            
            '��д����
            If Not rsData.EOF Then
                .TextMatrix(int�̶���, i) = rsData!����
            Else
                .TextMatrix(int�̶���, i) = ""
            End If
        Next
    End With
End Sub

Private Sub RefreshList_����ҩ(ByVal rsData As ADODB.Recordset)
    '���ܣ�ˢ�´���ҩ�б�
    Dim i As Integer
    Dim n As Integer
    Dim int���� As Integer
    Dim intTemp As Integer
    
    int���� = IIf(mPara.bln������ģʽ, 0, 3)
    
    With vsf��ȡҩ
        For i = 0 To .Cols - 1
            rsData.Filter = "��ҩ���� = '" & .ColKey(i) & "'"
            
            '��ʱ��ո�������
            For n = int���� + 1 To .Rows - 1
                .TextMatrix(n, i) = ""
            Next
            
            intTemp = IIf(rsData.RecordCount > mPara.lng����ҩ_����, mPara.lng����ҩ_����, rsData.RecordCount)
            
            '���¼��ظ�������
            For n = 1 To intTemp
                .TextMatrix(n + int����, i) = rsData!����
                
                rsData.MoveNext
            Next
        Next
    End With
End Sub

Private Sub RefreshList_�ѹ���(ByVal rsData As ADODB.Recordset)
    '���ܣ�ˢ���ѹ����б�
    Dim i As Integer
    Dim n As Integer
    Dim int���� As Integer
    Dim intTemp As Integer
    
    int���� = IIf(mPara.bln������ģʽ, 0, 1)
    
    With vsf�ѹ���
        For i = 0 To .Cols - 1
            rsData.Filter = "��ҩ���� = '" & .ColKey(i) & "'"
            
            '��ʱ��ո�������
            For n = int���� + 1 To .Rows - 1
                .TextMatrix(n, i) = ""
            Next
            
            intTemp = IIf(rsData.RecordCount > mPara.lng�ѹ���_����, mPara.lng�ѹ���_����, rsData.RecordCount)
            
            '���¼��ظ�������
            For n = 1 To intTemp
                .TextMatrix(n + int����, i) = rsData!����
                
                rsData.MoveNext
            Next
        Next
    End With
    
End Sub

Private Sub LoadData_������()
    '���ܣ���������������
    '��Σ���strWins������ķ�ҩ����
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    If mPara.bln������ģʽ Then Exit Sub
    
    strWins = mPara.str�ര��
    
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select ����,��ҩ����" & vbNewLine & _
            "From δ��ҩƷ��¼" & vbNewLine & _
            "Where �Ŷ�״̬ = 3 And �ⷿid = [1] And ��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) And �������� Between [3] And [4]"

    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "��������������", mlngҩ��ID, strWins, date��ʼ����, date��������)
    
    'ˢ�������н�������
    Call RefreshList_������(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData_����ҩ()
    '���ܣ����ش���ҩ����
    '��Σ���strWins������ķ�ҩ����
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    strWins = IIf(mPara.bln������ģʽ, mstrWins, mPara.str�ര��)
    
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select A.����ID,A.����,B.��ҩ����,B.ǩ��ʱ��,B.��������,A.��ҩ���� " & _
            "From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B " & _
            "Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid And (A.����=8 or A.����=9 or A.����=10) "

    If mbln��ҩ Then
        strSQL = strSQL & " and A.�Ŷ�״̬=2 and A.�ⷿid=[1] and A.��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    ElseIf mbln��ҩȷ�� And mbln��ҩ = False Then
        strSQL = strSQL & " and (A.�Ŷ�״̬=1 or A.�Ŷ�״̬=2) and A.�ⷿid=[1] and A.��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    ElseIf mbln��ҩ = False And mbln��ҩȷ�� = False Then
        strSQL = strSQL & " and (A.�Ŷ�״̬<3 or A.�Ŷ�״̬ is null) and A.�ⷿid=[1] and A.��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)"
    End If
    
    strSQL = "Select Rownum ���,����,����,��ҩ���� " & _
            "From ( " & _
            "Select min(" & IIf(mbln��ҩ, "��ҩ����", "Nvl(ǩ��ʱ��,��������)") & ") ����,����id,����,��ҩ���� " & _
            "From (" & strSQL & ") " & _
            "Where ����ID Not In (Select distinct A.����ID From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C " & _
            "Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) " & _
            "  and (A.�Ŷ�״̬=4 or A.�Ŷ�״̬ = 3) and A.�ⷿid=[1] and A.��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)) " & _
            "Group By ����,����id,��ҩ���� " & _
            "Order by ���� " & _
            ")"
                    
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "���ش���ҩ����", mlngҩ��ID, strWins, date��ʼ����, date��������)
    
    'ˢ�´���ҩ��������
    Call RefreshList_����ҩ(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub LoadData_�ѹ���()
    '���ܣ������ѹ�������
    '��Σ���strWins������ķ�ҩ����
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim date��ʼ���� As Date
    Dim date�������� As Date
    Dim strWins As String
    
    On Error GoTo errHandle
    
    strWins = IIf(mPara.bln������ģʽ, mstrWins, mPara.str�ര��)
    
    date��ʼ���� = gobjDatabase.Currentdate
    date��ʼ���� = CDate(Format(date��ʼ����, "yyyy-mm-dd") & " 00:00:00")

    date�������� = gobjDatabase.Currentdate
    date�������� = CDate(Format(date��������, "yyyy-mm-dd") & " 23:59:59")
    
    strSQL = "Select distinct A.����ID,A.����,A.����ʱ��, A.��ҩ���� From δ��ҩƷ��¼ A,ҩƷ�շ���¼ B,������ü�¼ C" & _
            " Where A.����=B.���� And A.No=B.NO And A.�ⷿid=B.�ⷿid and B.����id=C.id and (A.����=8 or A.����=9 or A.����=10) " & _
            "   and A.�Ŷ�״̬=4 and A.�ⷿid=[1] and A.��ҩ���� In (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) and A.�������� between [3] and [4] And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0) " & _
            " And a.����ʱ�� < Sysdate - 1 / 24 / 60 / 60 * [5] "
            
    strTemp = Replace(strSQL, "������ü�¼", "סԺ���ü�¼")
            
    strSQL = strSQL & " union all " & strTemp
    
    strSQL = "Select rownum ���,����,��ҩ���� " & _
            "From (Select ����ID,����,min(����ʱ��) ����ʱ��,��ҩ���� " & _
            "From (" & strSQL & ") " & _
            "Group by ����ID,����,��ҩ���� " & _
            "Order by ����ʱ�� asc " & _
            ")"
    
    Set rsData = gobjDatabase.OpenSQLRecord(strSQL, "�����ѹ�������", mlngҩ��ID, strWins, date��ʼ����, date��������, 1)
    
    'ˢ���ѹ��Ž�������
    Call RefreshList_�ѹ���(rsData)
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub DrawWallPaper()
    '���ܣ����ñ��ؼ��ı���ͼƬΪ��񱳺������ͼƬ
    Dim std As StdPicture
    Dim strTempFile As String
    
    On Error GoTo errHandle
    
    '����ҩ�����
    '--------------------------------
    picDraw.cls
    picDraw.Width = vsf��ȡҩ.Width
    picDraw.Height = vsf��ȡҩ.Height
    
    picDraw.PaintPicture img����.Picture, -vsf��ȡҩ.Left, -vsf��ȡҩ.Top, Me.ScaleWidth, Me.ScaleHeight
    
    Set vsf��ȡҩ.WallPaper = picDraw.Image
    '--------------------------------
    
    '�ѹ��ű����
    '--------------------------------
    picDraw.cls
    picDraw.Width = vsf�ѹ���.Width
    picDraw.Height = vsf�ѹ���.Height
    
    picDraw.PaintPicture img����.Picture, -vsf�ѹ���.Left, -vsf�ѹ���.Top, Me.ScaleWidth, Me.ScaleHeight
    
    Set vsf�ѹ���.WallPaper = picDraw.Image
    '--------------------------------

    Exit Sub
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub timData_Timer()
    '���ܣ���ʱˢ����ʾ������
    
    Call LoadData_������
    Call LoadData_����ҩ
    Call LoadData_�ѹ���
    
    Call Change��������
End Sub

Private Sub timRest_Timer()
    mbln������ʾ = False
    
    Call Show��������(False)
End Sub

Private Sub timTime_Timer()
    '���ܣ�ˢ��ʱ��
    lblTime.Caption = Format(gobjDatabase.Currentdate, "yyyy-mm-dd  hh:mm")
End Sub
