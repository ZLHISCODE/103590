VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCardSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ�������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmCardSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleMode       =   0  'User
   ScaleWidth      =   5969.773
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkDebug 
      Caption         =   "��¼��־"
      Height          =   225
      Left            =   150
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      Top             =   2400
      Width           =   1020
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "����(&E)"
      Height          =   350
      Index           =   2
      Left            =   2325
      TabIndex        =   4
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "����(&S)"
      Height          =   350
      Index           =   1
      Left            =   1215
      TabIndex        =   3
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Index           =   3
      Left            =   4620
      TabIndex        =   2
      Top             =   2340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Index           =   0
      Left            =   3510
      TabIndex        =   1
      Top             =   2340
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2145
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   5685
      _cx             =   10028
      _cy             =   3784
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
Attribute VB_Name = "frmCardSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCardN0 As String
Private mLasterRow As Integer

Friend Function SelectCard(ByVal colCards As Collection, ByVal intCount As Integer, Optional ByVal FrmMain As Object) As Integer
    Dim objCard As clsCard, lstItem As ListItem
    On Error GoTo errHandle
    If intCount > 0 Then
        '����״̬��ֻ��ʾ�����õĿ�
        Call LoadData(colCards, True)
    ElseIf intCount = 0 Then
        'δ���õ�����¶�������ʾ���п�
        Call LoadData(colCards, False)
    ElseIf intCount = -1 Then
        '����״̬�£���ʾ���еĿ�
        Call LoadData(colCards, False)
    
    End If
    
    If intCount <> -1 Then
        Me.cmdCard(1).Visible = False
        Me.cmdCard(2).Visible = False
        chkDebug.Visible = False
    End If
    If FrmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, FrmMain
    End If
    SelectCard = Val(mstrCardN0)
    mstrCardN0 = ""
    Exit Function
errHandle:
    Call WritLog("CardSelect.SelectCard", "", err.Description)
End Function

Private Sub cmdCard_Click(Index As Integer)
    With vfgList
    Select Case Index
        Case 0 '-ȷ��
            If Val(.TextMatrix(.Row, .ColIndex("����"))) > 0 Then
                mstrCardN0 = Val(.TextMatrix(.Row, .ColIndex("����")))
                
                Call SaveSetting("ZLSOFT", "����ģ��\zlICCard", "����", chkDebug.value)
                mLasterRow = .Row
                If mLasterRow <= 0 Then mLasterRow = 1
                Call SaveSetting("ZLSOFT", "����ģ��\zlICCard", "LastSelect", mLasterRow)
               
                Unload Me
            End If
        Case 1 '-����
            Dim objCard As clsCardDev
            If Val(.TextMatrix(.Row, .ColIndex("����"))) > 0 Then
                Set objCard = CreateObject(.TextMatrix(.Row, .ColIndex("�ӿ�")))
                mLasterRow = .Row
                objCard.SetCard
            End If
        Case 2 '-����,ͣ��
            Call CardEnable
        Case 3 '-�˳�
            Unload Me
    End Select
    End With
End Sub


Private Sub CardEnable()
    Dim i As Integer
    Dim intCardNo As Integer
    With vfgList
            If Val(.TextMatrix(.Row, .ColIndex("����"))) > 0 Then
                intCardNo = Val(.TextMatrix(.Row, .ColIndex("����")))
                If .TextMatrix(.Row, .ColIndex("����")) = "��" Then
                    Call SaveSetting("ZLSOFT", "����ģ��\zlICCard", Val(.TextMatrix(.Row, .ColIndex("����"))), 0)
                    .TextMatrix(.Row, .ColIndex("����")) = "��"
                    cmdCard(2).Caption = "����(&S)"
                Else
                    Call SaveSetting("ZLSOFT", "����ģ��\zlICCard", Val(.TextMatrix(.Row, .ColIndex("����"))), 1)
                    .TextMatrix(.Row, .ColIndex("����")) = "��"
                    cmdCard(2).Caption = "ͣ��(&E)"
                End If
            End If

    End With
    For i = 1 To Cards.Count
        If Item(i).���� = intCardNo Then
            Item(i).���� = IIf(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("����")) = "��", False, True)
        End If
    Next
    
End Sub

Private Sub Form_Activate()
    chkDebug.value = GetSetting("ZLSOFT", "����ģ��\zlICCard", "����", 0)
    
    mLasterRow = Val(GetSetting("ZLSOFT", "����ģ��\zlICCard", "LastSelect", 1))
    If mLasterRow = 0 Then mLasterRow = 1
    
    Call vfgList_EnterCell
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdCard_Click(0)
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdCard_Click(3)
    ElseIf KeyAscii = vbKeySpace Then
        If cmdCard(1).Visible Then
            'Call lvwCardType_DblClick
            Call CardEnable
        End If
    End If
End Sub

Private Sub Form_Load()
    frmTimer.tmrMain.Enabled = False
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If cmdCard(1).Visible Then
            If Val(.TextMatrix(.MouseRow, .ColIndex("����"))) > 0 Then
                .Select .MouseRow, .ColIndex("����")
                Call CardEnable
            End If
        Else
            Call cmdCard_Click(0)
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
    cmdCard(1).Enabled = False
    With vfgList
        If .ColIndex("����") < 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("����"))) > 0 Then
            If .TextMatrix(.Row, .ColIndex("����")) = "1" Then
                cmdCard(1).Enabled = True
            End If
            If .TextMatrix(.Row, .ColIndex("����")) = "��" Then
                cmdCard(2).Caption = "ͣ��(E)"
            Else
                cmdCard(2).Caption = "����(S)"
            End If
        End If
    End With
End Sub

'---


Public Sub LoadData(ByVal objCards As Collection, ByVal bln���� As Boolean)
    
    Dim strHead As String, objCard As clsCard
    '1 ����� 4 ���� 7 �Ҷ���
    On Error GoTo errHandle
    If bln���� Then
        strHead = "����,600,4;����,4200,1;����,1,4;�Զ�,1,4;�ӿ�,1,0;����,1,0;����,1,0"
    Else
        strHead = "����,600,4;����,3600,1;����,600,4;�Զ�,600,4;�ӿ�,1,0;����,1,0;����,1,0"
    End If
    
    With vfgList
        .Clear
        Call SetVsFlexGridHead(strHead, vfgList)
        
        For Each objCard In objCards
             If bln���� = True Then
                 If objCard.���� Then
                    '����ʾ����
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(objCard.����, "��", "��")
                    .TextMatrix(.Rows - 1, .ColIndex("�Զ�")) = IIf(objCard.�Ƿ��Զ���ȡ, "��", "��")
                    .TextMatrix(.Rows - 1, .ColIndex("�ӿ�")) = objCard.�ӿڳ�����
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.�ɷ�����
                    .Rows = .Rows + 1
                End If
            Else
                'ȫ��
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(objCard.����, "��", "��")
                .TextMatrix(.Rows - 1, .ColIndex("�Զ�")) = IIf(objCard.�Ƿ��Զ���ȡ, "��", "��")
                .TextMatrix(.Rows - 1, .ColIndex("�ӿ�")) = objCard.�ӿڳ�����
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.����
                .TextMatrix(.Rows - 1, .ColIndex("����")) = objCard.�ɷ�����
                .Rows = .Rows + 1
            End If
            
        Next
        
        If .Rows > 0 Then
            .Rows = .Rows - 1
        End If
        '��ѡ��
        .SelectionMode = flexSelectionByRow
        
        If mLasterRow > 0 And mLasterRow < .Rows Then
            .Select mLasterRow, 1
            .TopRow = mLasterRow
        End If
    End With
    Exit Sub
errHandle:
    Call WritLog("CardSelect.LoadData", "", err.Description)
End Sub
Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub
