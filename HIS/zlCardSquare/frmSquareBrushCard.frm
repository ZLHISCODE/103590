VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareBrushCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ѿ�ˢ��"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareBrushCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10530
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      Caption         =   "������Ϣ"
      Height          =   2115
      Left            =   285
      TabIndex        =   23
      Top             =   420
      Width           =   3900
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1155
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1155
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   315
         Width           =   2550
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ���"
         Height          =   240
         Index           =   10
         Left            =   90
         TabIndex        =   6
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&W)"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   0
         Top             =   390
         Width           =   840
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   3
         Top             =   735
         Width           =   1005
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   1155
         TabIndex        =   7
         Top             =   1635
         Width           =   2550
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ˢ������"
      Height          =   2115
      Left            =   4290
      TabIndex        =   18
      Top             =   435
      Width           =   3990
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1830
         TabIndex        =   9
         Top             =   1455
         Width           =   2025
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   1830
         TabIndex        =   22
         Top             =   360
         Width           =   2025
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ܶ�"
         Height          =   240
         Index           =   6
         Left            =   810
         TabIndex        =   21
         Top             =   435
         Width           =   960
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�������ˢ����"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   975
         Width           =   1680
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   1830
         TabIndex        =   19
         Top             =   893
         Width           =   2025
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&X)"
         Height          =   240
         Index           =   4
         Left            =   450
         TabIndex        =   8
         Top             =   1530
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ����ǰ��(&D)"
      Height          =   465
      Left            =   8535
      TabIndex        =   17
      Top             =   3285
      Width           =   1860
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   465
      Left            =   8655
      TabIndex        =   14
      Top             =   765
      Width           =   1185
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   465
      Left            =   8655
      TabIndex        =   13
      Top             =   255
      Width           =   1185
   End
   Begin VB.Frame fra 
      Height          =   3540
      Left            =   135
      TabIndex        =   15
      Top             =   75
      Width           =   8295
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   3
         Left            =   1095
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2565
         Width           =   7035
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "����ˢ��һ�ſ�(&K)"
         Height          =   420
         Left            =   5835
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2970
         Width           =   2325
      End
      Begin VB.Label lblʧЧ�� 
         Height          =   240
         Left            =   420
         TabIndex        =   24
         Top             =   3105
         Width           =   4455
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ע(&S)"
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   840
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3060
      Left            =   120
      TabIndex        =   16
      Top             =   3810
      Width           =   10305
      _cx             =   18177
      _cy             =   5397
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareBrushCard.frx":030A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   120
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
      ExplorerBar     =   7
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
Attribute VB_Name = "frmSquareBrushCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String, mintCallType As Integer, mintSucces As Integer, mblnChange As Boolean
Private mrsData As ADODB.Recordset, mrsFeeList As ADODB.Recordset
Private mlng�ӿڱ�� As Long
Private mblnCardNoSHowPW As Boolean '�Ƿ�������ʾ����
Private Type CardInfor
    lng���ѿ�ID As Long
    str���� As String
    dbl��� As Double
    dbl������Ѷ� As Double
    dblʧЧ��� As Double '��ȡ��ԭ����,�Ƚ��ȳ��ķ���:�����ѿ����,��������ֵ��:�˽��Ϊ,���ں�δ���ѵĽ��
    str������� As String
    str�ӿ����� As String
    str���㷽ʽ As String
End Type
Private mdbl������Ѷ� As Double, mdbl�������ۼ� As Double
Private mTyCurCardInfor As CardInfor
Private Enum mtxtIdx
    idx_txt���� = 0
    idx_txt���� = 1
    idx_txt�������� = 2
    idx_txt��ע = 3
End Enum
Private Enum mlblIdx
    idx_lbl������ = 0
    idx_lbl��� = 1
    idx_lbl������� = 2
    idx_lbl�ܷ��� = 3
End Enum
Private mrsRequare As New ADODB.Recordset
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mobjKeyboard As Object
 

Private Function CheckDepended() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĹ�����
    '����:���˺�
    '����:2009-12-24 12:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet���ѿ��ӿ�
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "���=" & mlng�ӿڱ��, , , 1
    If rsTemp.EOF Then
        ShowMsgbox "�ӿ�δ�ҵ�(���Ϊ" & mlng�ӿڱ�� & "),����!"
        Exit Function
    End If
    With mTyCurCardInfor
        .str�ӿ����� = Nvl(rsTemp!����)
        .str���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
        txtEdit(mtxtIdx.idx_txt����).MaxLength = Len(Nvl(rsTemp!ǰ׺�ı�)) + Val(Nvl(rsTemp!���ų���))
    End With
    CheckDepended = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lng�ӿڱ�� As Long, ByVal intCallType As Integer, _
    ByVal rsFeeList As ADODB.Recordset, dbl������Ѷ� As Double, rsRequare As ADODB.Recordset) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ���ӿ�
    '���:frmMain-���õ�������
    '     lngModule-���õ�ģ���
    '     strPrivs-���õ�Ȩ�޴�
    '     dbl������Ѷ�-����ˢ�������ˢ����
    '     rsFeeList-������ϸ��Ϣ()
    '����:rsRequare-���ؽ�����Ϣ
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 10:33:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    mdbl������Ѷ� = dbl������Ѷ�: mlng�ӿڱ�� = lng�ӿڱ��: mintCallType = intCallType
    Set mrsFeeList = rsFeeList  '������ϸ:
    Set mrsRequare = rsRequare
    
    If CheckDepended = False Then Exit Function
        
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowBrushCard = mintSucces > 0
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
'���ϴ�ˢ������Ϣ���ص�����
Private Function LoadPreBrushCardToVsGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ϴ�ˢ������Ϣ���ص�����
    '����:�ӳɳɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2009-12-24 11:42:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl������ As Double
    
    Err = 0: On Error GoTo ErrHand:
    lngRow = 1
    With vsGrid
        .Rows = 2
        .Clear 1
        .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = &H80000008
        lngRow = 1
        If mrsRequare.RecordCount <> 0 Then mrsRequare.MoveFirst
        
        Do While Not mrsRequare.EOF
            If Val(Nvl(mrsRequare!�ӿڱ��)) = mlng�ӿڱ�� Then
                If lngRow > 1 Then
                    .Rows = .Rows + 1
                End If
                
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(mblnCardNoSHowPW, "*******", Nvl(mrsRequare!����))
                .Cell(flexcpData, lngRow, .ColIndex("����")) = Nvl(mrsRequare!���ѿ�ID) & "-" & Nvl(mrsRequare!����)
                
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = Nvl(mrsRequare!���㷽ʽ)
                .Cell(flexcpData, lngRow, .ColIndex("���㷽ʽ")) = Nvl(mrsRequare!������)
                .TextMatrix(lngRow, .ColIndex("�����")) = Format(Val(Nvl(mrsRequare!���)), gVbFmtString.FM_���)
                .Cell(flexcpData, lngRow, .ColIndex("�����")) = Val(Nvl(mrsRequare!���))
                .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(Nvl(mrsRequare!������)), gVbFmtString.FM_���)
                dbl������ = dbl������ + Val(Nvl(mrsRequare!������))
                .TextMatrix(lngRow, .ColIndex("��ע")) = Nvl(mrsRequare!��ע)
                lngRow = lngRow + 1
            End If
            mrsRequare.MoveNext
        Loop
        mdbl�������ۼ� = dbl������
         If .Rows - 1 >= 1 And dbl������ <> 0 Then
            If .TextMatrix(1, .ColIndex("����")) <> "" Then
                .Rows = .Rows + 1
                lngRow = .Rows - 1
                .TextMatrix(lngRow, .ColIndex("����")) = "�ϼ�"
                .Cell(flexcpData, lngRow, .ColIndex("����")) = ""
                .TextMatrix(lngRow, .ColIndex("�����")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("�����")) = ""
                .TextMatrix(lngRow, .ColIndex("��������")) = Format(dbl������, gVbFmtString.FM_���)
                .TextMatrix(lngRow, .ColIndex("��ע")) = ""
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
            End If
        End If
    End With
    LoadPreBrushCardToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cmdDel_Click()
    Call MoveVsGridRowData
End Sub

Private Sub cmdNext_Click()
    If zlInsertDataToGrid = False Then Exit Sub
    zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
End Sub

Private Sub cmdȡ��_Click()
    mintSucces = 0: Unload Me
End Sub
Private Sub cmdȷ��_Click()
    Dim lngRow As Long, blnInputNotGrid As Boolean  '�������Ϣ,û��������������,ֱ�Ӵ���
    Dim dt����ʱ�� As Date, blnHaveData As Boolean
    blnInputNotGrid = False
    If Trim(txtEdit(mtxtIdx.idx_txt����).Text) <> "" Then
        '��Ҫ���
        If CheckCardNotExists(Trim(txtEdit(mtxtIdx.idx_txt����).Text), False) Then
            '������,��ʾ��Ҫ����Ƿ�Ϸ�
            If CheckInput = False Then Exit Sub
            blnInputNotGrid = True
        End If
         
    Else
        '����Ƿ�������
        blnHaveData = False
        With vsGrid
            For lngRow = 1 To .Rows - 1
                If Val(Split(.Cell(flexcpData, lngRow, .ColIndex("����")) & "-", "-")(0)) <> 0 Then
                    blnHaveData = True: Exit For
                End If
            Next
        End With
'        If blnHaveData = False Then
'            ShowMsgbox "������ˢ������,����"
'            zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
'            Exit Sub
'        End If
    End If
   
    '������صĽ�����Ϣ
    '��ɾ��
    With mrsRequare
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Val(Nvl(mrsRequare!�ӿڱ��)) = mlng�ӿڱ�� Then
               .Delete adAffectCurrent
               .Update
               .MoveNext
               If .RecordCount <> 0 Then .MoveFirst
            Else
               .MoveNext
            End If
        Loop
    End With
    
    Dim varData As Variant
    dt����ʱ�� = zlDatabase.Currentdate
    With vsGrid
        
        For lngRow = 1 To .Rows - 1
            varData = Split(.Cell(flexcpData, lngRow, .ColIndex("����")) & "-", "-")
            If Val(varData(0)) <> 0 Then
                mrsRequare.AddNew
                mrsRequare!�ӿڱ�� = mlng�ӿڱ��
                mrsRequare!���ѿ�ID = Val(varData(0))
                mrsRequare!���� = Trim(varData(1))
                mrsRequare!���㷽ʽ = mTyCurCardInfor.str���㷽ʽ
                mrsRequare!������ = mTyCurCardInfor.str�ӿ�����
                mrsRequare!��� = Val(.Cell(flexcpData, lngRow, .ColIndex("�����")))
                mrsRequare!������ = Val(.TextMatrix(lngRow, .ColIndex("��������")))
                mrsRequare!����ʱ�� = dt����ʱ��
                mrsRequare!��ע = Trim(.TextMatrix(lngRow, .ColIndex("��ע")))
                mrsRequare!�����־ = 0
                mrsRequare.Update
            End If
        Next
    End With
    If blnInputNotGrid Then
        mrsRequare.AddNew
        mrsRequare!�ӿڱ�� = mlng�ӿڱ��
        mrsRequare!���ѿ�ID = mTyCurCardInfor.lng���ѿ�ID
        mrsRequare!���� = mTyCurCardInfor.str����
        mrsRequare!���㷽ʽ = mTyCurCardInfor.str���㷽ʽ
        mrsRequare!������ = mTyCurCardInfor.str�ӿ�����
        mrsRequare!��� = mTyCurCardInfor.dbl���
        mrsRequare!������ = Val(txtEdit(mtxtIdx.idx_txt��������).Text)
        mrsRequare!����ʱ�� = dt����ʱ��
        mrsRequare!��ע = Trim(txtEdit(mtxtIdx.idx_txt��ע).Text)
        mrsRequare!�����־ = 0
        mrsRequare.Update
    End If
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim dbl�ܷ��� As Double
    '����Ƿ���������ص�ˢ������
     Call CreateObjectKeyboard
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng�ӿڱ��)
    '��ȡ�ܶ�
    mblnCardNoSHowPW = zlIsCardNoShowPW(mlng�ӿڱ��)
    If mblnCardNoSHowPW Then
        txtEdit(mtxtIdx.idx_txt����).PasswordChar = "*"
    Else
        txtEdit(mtxtIdx.idx_txt����).PasswordChar = ""
    End If
    
    Call LoadPreBrushCardToVsGrid
    lblInfor(mlblIdx.idx_lbl�ܷ���).Caption = Format(grsStatic.dbl�����ܶ�, gVbFmtString.FM_���)
    Call vsGrid_LostFocus
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index <> mtxtIdx.idx_txt���� Then txtEdit(Index).Tag = ""
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt����
        gTy_TestBug.BytType = 2
        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")
        zlControl.TxtSelAll txtEdit(Index)
    Case mtxtIdx.idx_txt��ע
        zlControl.TxtSelAll txtEdit(Index)
        zlCommFun.OpenIme True
    Case Else
        zlControl.TxtSelAll txtEdit(Index)
        zlCommFun.OpenIme False
        If Index = idx_txt���� Then
            Call OpenPassKeyboard(txtEdit(Index))
        End If
    End Select
End Sub
Private Function CheckInputPassWord() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������Ƿ���ȷ
    '����:��ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txtEdit(mtxtIdx.idx_txt����).Tag) <> "" And Trim(txtEdit(mtxtIdx.idx_txt����).Text) = "" Then
        ShowMsgbox "����δ����,����!"
        Exit Function
    End If
    
    If Trim(txtEdit(mtxtIdx.idx_txt����).Tag) <> Trim(txtEdit(mtxtIdx.idx_txt����).Text) Then
        ShowMsgbox "�����������,����!"
        Exit Function
    End If
    CheckInputPassWord = True
End Function

Private Function CheckInputSquareMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ı������ѽ���Ƿ���ȷ
    '����:��ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt��������).Text), 16, True, True, 0, "�����") = False Then
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl�������).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt��������).Text)) Then
        ShowMsgbox "�������������:" & Format(Val(lblInfor(mlblIdx.idx_lbl�������).Tag), gVbFmtString.FM_���) & "Ԫ,����!"
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl���).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt��������).Text)) Then
        ShowMsgbox "������(" & Format(Val(lblInfor(mlblIdx.idx_lbl���).Caption), gVbFmtString.FM_���) & "Ԫ),����!"
        Exit Function
    End If
    
    CheckInputSquareMoney = True
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str���� As String, str���� As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt����
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
         
        '���ǿ��ܴ��ڲ���Ա��ˢ�������,����ݲ��������¹���:
        If IsDesinMode = False Then Exit Sub
        If txtEdit(Index).Text = "" Then
            'ֱ�ӵ�����
            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
                Exit Sub
            End If
            txtEdit(Index).Text = strCardNo
            txtEdit(Index).Tag = strCardNo
        End If
        
        If zlBrusCard(Trim(txtEdit(Index))) = False Then
            zlCtlSetFocus txtEdit(Index)
        Else
            If txtEdit(mtxtIdx.idx_txt����).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt��ע).Enabled And txtEdit(mtxtIdx.idx_txt��ע).Visible Then txtEdit(mtxtIdx.idx_txt��ע).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
        
    Case mtxtIdx.idx_txt��ע
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    Case mtxtIdx.idx_txt����
        If CheckInputPassWord = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt��������
        If CheckInputSquareMoney = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean
    
    Select Case Index
    Case mtxtIdx.idx_txt����
        If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If IsDesinMode Then Exit Sub
        
        Call BrushCard(txtEdit(Index), KeyAscii)
    Case mtxtIdx.idx_txt��ע
        blnCard = zlInputIsCard(txtEdit(Index), KeyAscii, glngSys, mblnCardNoSHowPW)
        If blnCard = True Then KeyAscii = 0
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    Case mtxtIdx.idx_txt����
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    Case mtxtIdx.idx_txt��������
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
    Case Else
    End Select
End Sub
Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������(Ŀǰֻ֧���п�����ˢ��)
    '����:���˺�
    '����:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 And objEdit.SelLength <> Len(objEdit.Text)
    
    If blnCard Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        'ˢ������:
        If zlBrusCard(Trim(objEdit)) = False Then
            zlCtlSetFocus objEdit
        Else
            If txtEdit(mtxtIdx.idx_txt����).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt��ע).Enabled And txtEdit(mtxtIdx.idx_txt��ע).Visible Then txtEdit(mtxtIdx.idx_txt��ע).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub
Private Function CheckIsBreshCard(ByVal objEdit As Object, KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�ˢ������
    '����:��ˢ��,����true,���򷵻�False
    '����:���˺�
    '����:2010-10-25 09:52:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
    End If
    '��ȫˢ�����
    If KeyAscii <> 0 And KeyAscii > 32 Then
        sngNow = Timer
        If objEdit.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") > 0.04 Then '>0.007>=0.01
            '����ˢ����
            blnCard = False
             sngBegin = sngNow
        Else
            blnCard = True
        End If
    End If
End Function

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = idx_txt���� Then
        Call ClosePassKeyboard(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt���� Then Exit Sub
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt���� Then Exit Sub
    If Button = 2 Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case mtxtIdx.idx_txt����
    Case mtxtIdx.idx_txt��ע
    Case mtxtIdx.idx_txt����
        If CheckInputPassWord = False Then
        End If
    Case mtxtIdx.idx_txt��������
        If CheckInputSquareMoney = False Then
           'Cancel = 1
        End If
    Case Else
    End Select
End Sub

Private Function zlBrusCard(ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������
    '����:���˺�
    '����:2009-12-16 10:33:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    
    With mTyCurCardInfor
        .dblʧЧ��� = 0
        .dbl��� = 0
        .dbl������Ѷ� = 0
        .str���� = ""
        .lng���ѿ�ID = 0
    End With
    
    gstrSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��,  a.����," & _
    "          to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬, " & _
    "          to_char(a.������," & gOraFmtString.FM_��� & ") as ������ ," & _
    "          to_char(a.���۽��," & gOraFmtString.FM_��� & ") as ���۽�� ," & _
    "          to_char(a.��ֵ�ۿ���," & gOraFmtString.FM_�ۿ��� & ") as ��ֵ�ۿ��� ," & _
    "          to_char(a.���," & gOraFmtString.FM_��� & ") as ��� ," & _
    "          to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.������� " & _
    "   From ���ѿ�Ŀ¼ A  " & _
    "   Where A.���� = [1] and A.�ӿڱ��=[2] And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��)  " & _
    "   Order by a.���"
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng�ӿڱ��)
    If rsTemp.EOF Then
       ShowMsgbox "δ�ҵ���ص����ѿ���¼,����!"
        Exit Function
    End If
    '���:
    '�Ƿ����
    If Nvl(rsTemp!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "�����ѿ��Ѿ���" & Nvl(rsTemp!��ǰ״̬) & ",������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "�����ѿ��Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "�����ѿ��Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    
    '���Ч��
    mTyCurCardInfor.dbl��� = Val(Nvl(rsTemp!���))
    lblʧЧ��.Visible = False
    If Nvl(rsTemp!��Ч��, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
       '������Ч��
       If Val(Nvl(rsTemp!�ɷ��ֵ)) = 1 Then
          '������ֵ��,���ڵ�,�������ѿ�����,ֻ��������ֵ����
          mTyCurCardInfor.dblʧЧ��� = zlGetʧЧ���(Val(Nvl(rsTemp!ID)), mlng�ӿڱ��)
          mTyCurCardInfor.dbl��� = IIf(mTyCurCardInfor.dbl��� - mTyCurCardInfor.dblʧЧ��� < 0, 0, mTyCurCardInfor.dbl��� - mTyCurCardInfor.dblʧЧ���)
          If mTyCurCardInfor.dblʧЧ��� <> 0 Then
            lblʧЧ��.Caption = "��ǰ����ʧЧ���(�����)Ϊ��" & Format(mTyCurCardInfor.dblʧЧ���, gVbFmtString.FM_���) & "Ԫ"
            lblʧЧ��.Visible = True
            lblʧЧ��.ForeColor = vbRed
          End If
          
       Else
            '��������ֵ��,�����ٽ�������
            ShowMsgbox "����Ϊ" & strCardNo & "�����ѿ��Ѿ�ʧЧ,������ˢ��"
            Exit Function
       End If
    End If
    If mTyCurCardInfor.dbl��� <= 0 Then
        ShowMsgbox "����Ϊ" & strCardNo & "�����ѿ��Ѿ�û�����,������ˢ��"
        Exit Function
    End If
    If CheckCardNotExists(strCardNo, True) = False Then
    
        Exit Function
    End If
    
    With mTyCurCardInfor
        .lng���ѿ�ID = Val(Nvl(rsTemp!ID))
        .str���� = Nvl(rsTemp!����)
        .str������� = Nvl(rsTemp!�������)
        .dbl������Ѷ� = zl��ȡ������Ѷ�(.str�������, mdbl������Ѷ�, mdbl�������ۼ�)
    End With
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����).Tag = Nvl(rsTemp!����)
    lblInfor(mlblIdx.idx_lbl���).Caption = Format(Val(Nvl(rsTemp!���)), gVbFmtString.FM_���)
    lblInfor(mlblIdx.idx_lbl������).Caption = Nvl(rsTemp!������)
    txtEdit(mtxtIdx.idx_txt����).Tag = Nvl(rsTemp!����)
    lblInfor(mlblIdx.idx_lbl�������).Caption = Format(mTyCurCardInfor.dbl������Ѷ�, gVbFmtString.FM_���)
    
    'ȱʡֵ:����,ȱʡ���,����Ϊ������Ѷ�
    If mTyCurCardInfor.dbl��� < mTyCurCardInfor.dbl������Ѷ� Then
        txtEdit(mtxtIdx.idx_txt��������).Text = Format(mTyCurCardInfor.dbl���, gVbFmtString.FM_���)
    Else
        txtEdit(mtxtIdx.idx_txt��������).Text = Format(mTyCurCardInfor.dbl������Ѷ�, gVbFmtString.FM_���)
    End If
    zlBrusCard = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function CheckCardNotExists(ByVal strCardNo As String, Optional blnMsgbox As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鿨Ƭ��Ϣ�Ƿ���ˢ���д���
    '     strCardNO-����
    '����:������,����True,���򷵻�False
    '����:���˺�
    '����:2009-12-23 17:07:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, varData As Variant
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        For lngRow = 1 To .Rows - 1
            varData = Split(.Cell(flexcpData, lngRow, .ColIndex("����")) & "-", "-")
            If Trim(varData(0)) <> "" Then
                If Trim(Trim(varData(1))) = strCardNo Then
                    If blnMsgbox Then
                        If mblnCardNoSHowPW Then
                            ShowMsgbox "��ǰ�����Ѿ��ڵ�" & lngRow & "�д�����,������ˢ��!"
                        Else
                            ShowMsgbox "����Ϊ" & strCardNo & "�Ѿ��ڵ�" & lngRow & "�д�����,������ˢ��!"
                        End If
                        .Row = lngRow: .Col = .ColIndex("����")
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    CheckCardNotExists = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 17:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If txtEdit(mtxtIdx.idx_txt����).Text <> Trim(txtEdit(mtxtIdx.idx_txt����).Tag) Or Trim(txtEdit(mtxtIdx.idx_txt����).Text) = "" Then
        ShowMsgbox "δˢ����ˢ������ȷ,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        Exit Function
    End If
    If CheckInputPassWord = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        Exit Function
    End If
    If CheckInputSquareMoney = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt��������)
        Exit Function
    End If
    '����������Ƿ������ͬ�Ŀ�
    If CheckCardNotExists(Trim(txtEdit(mtxtIdx.idx_txt����))) = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        Exit Function
    End If
    CheckInput = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub ClearCtlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2009-12-24 11:11:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    txtEdit(mtxtIdx.idx_txt��������) = "0.00"
    txtEdit(mtxtIdx.idx_txt����) = ""
    txtEdit(mtxtIdx.idx_txt����) = ""
    txtEdit(mtxtIdx.idx_txt����).Tag = ""
    
    lblInfor(mlblIdx.idx_lbl������).Caption = ""
    lblInfor(mlblIdx.idx_lbl���).Caption = "0.00"
    lblInfor(mlblIdx.idx_lbl�������).Caption = "0.00"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume

End Sub

Private Function zlInsertDataToGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ˢ�����ݣ�������������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 17:00:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl���������ܶ� As Double
    Err = 0: On Error GoTo ErrHand:
    If CheckInput = False Then Exit Function
    
    With vsGrid
        If .Rows - 1 = 1 Then
            If Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("����"))) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        Else
            If Trim(.Cell(flexcpData, .Rows - 1, .ColIndex("����"))) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
        End If
        .TextMatrix(lngRow, .ColIndex("����")) = IIf(mblnCardNoSHowPW, "******", mTyCurCardInfor.str����)
        .Cell(flexcpData, lngRow, .ColIndex("����")) = mTyCurCardInfor.lng���ѿ�ID & "-" & mTyCurCardInfor.str����
        .TextMatrix(lngRow, .ColIndex("�����")) = Format(mTyCurCardInfor.dbl���, gVbFmtString.FM_���)
        .Cell(flexcpData, lngRow, .ColIndex("�����")) = mTyCurCardInfor.dbl���
        
        .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = mTyCurCardInfor.str���㷽ʽ
        .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(txtEdit(mtxtIdx.idx_txt��������).Text), gVbFmtString.FM_���)
        .TextMatrix(lngRow, .ColIndex("��ע")) = Trim(txtEdit(mtxtIdx.idx_txt��ע).Text)
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lblEdit(0).ForeColor
        grsStatic.dbl��ˢ�ۼƶ� = grsStatic.dbl��ˢ�ۼƶ� + Val(txtEdit(mtxtIdx.idx_txt��������).Text)
        dbl���������ܶ� = 0
        For lngRow = 1 To .Rows - 1
            dbl���������ܶ� = dbl���������ܶ� + Val(.TextMatrix(lngRow, .ColIndex("��������")))
        Next
        
        If .Rows - 1 >= 1 And dbl���������ܶ� <> 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("����")) = "�ϼ�"
            .Cell(flexcpData, lngRow, .ColIndex("����")) = ""
            .TextMatrix(lngRow, .ColIndex("�����")) = ""
            .Cell(flexcpData, lngRow, .ColIndex("�����")) = ""
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(dbl���������ܶ�, gVbFmtString.FM_���)
            .TextMatrix(lngRow, .ColIndex("��ע")) = ""
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
        End If
        
        mdbl�������ۼ� = dbl���������ܶ�
        Call ClearCtlData
    End With
    
    Call SetDelRowCtrlEnabled
    zlInsertDataToGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetDelRowCtrlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɾ���е�Eanbled����
    '����:���˺�
    '����:2009-12-24 10:50:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        If .Row < 0 Then cmdDel.Enabled = False: Exit Sub
        cmdDel.Enabled = Trim(.Cell(flexcpData, .Row, .ColIndex("����"))) <> ""
    End With
End Sub
Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow = NewRow Then Exit Sub
    Call SetDelRowCtrlEnabled
End Sub
Private Sub MoveVsGridRowData(Optional lngRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƴ�������
    '���:lngRow-ָ����(-1��ʾɾ����ǰ��)
    '����:���˺�
    '����:2009-12-24 10:52:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCurRow As Long, dbl�������� As Double, blnDeleCurRow As Long
    Dim lng�ϼ�Row As Long
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        If lngRow < 0 Then lngRow = .Row
        If lngRow < 0 Then Exit Sub
        blnDeleCurRow = lngRow = .Row
        lngCurRow = .Row
        
        dbl�������� = Val(.TextMatrix(lngRow, .ColIndex("��������")))
        grsStatic.dbl��ˢ�ۼƶ� = IIf(grsStatic.dbl��ˢ�ۼƶ� - dbl�������� < 0, 0, grsStatic.dbl��ˢ�ۼƶ� - dbl��������)
        
        If .Rows - 1 <= 1 Then
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            .Row = 1
        Else
            If .Cell(flexcpData, lngRow, .ColIndex("����")) = "" Then Exit Sub
            .RemoveItem lngRow
            If blnDeleCurRow Then
                If lngCurRow >= .Rows - 1 Then
                    .Row = .Rows - 1
                Else
                    .Row = lngCurRow + 1
                End If
            End If
        End If
        '���¼���ϼ���
        dbl�������� = 0
        For lngCurRow = 1 To .Rows - 1
            If .Cell(flexcpData, lngCurRow, .ColIndex("����")) <> "" Then
                dbl�������� = dbl�������� + Val(.TextMatrix(lngCurRow, .ColIndex("��������")))
            End If
            If .TextMatrix(lngCurRow, .ColIndex("����")) = "�ϼ�" Then lng�ϼ�Row = lngCurRow
        Next
        If dbl�������� = 0 And .Rows - 1 <= 2 Then
            .Clear 1: .Rows = 2: .Row = 1
            .Cell(flexcpForeColor, 1, 0, 1, .Cols - 1) = &H80000008
        Else
            '���Ӻϼ�
            If lngCurRow < 1 Then
                .Rows = .Rows + 1
                lng�ϼ�Row = .Rows - 1
            End If
            .TextMatrix(lng�ϼ�Row, .ColIndex("����")) = "�ϼ�"
            .Cell(flexcpData, lng�ϼ�Row, .ColIndex("����")) = ""
            .TextMatrix(lng�ϼ�Row, .ColIndex("�����")) = ""
            .Cell(flexcpData, lng�ϼ�Row, .ColIndex("�����")) = ""
            .TextMatrix(lng�ϼ�Row, .ColIndex("��������")) = Format(dbl��������, gVbFmtString.FM_���)
            .TextMatrix(lng�ϼ�Row, .ColIndex("��ע")) = ""
            .Cell(flexcpForeColor, lng�ϼ�Row, 0, lng�ϼ�Row, .Cols - 1) = vbBlue
        End If
        mdbl�������ۼ� = dbl��������
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub vsGrid_GotFocus()
  zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_LostFocus()
  zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

