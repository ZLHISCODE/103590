VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIDKindSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����������"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7875
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboDefaultCard 
      Height          =   300
      Left            =   1380
      TabIndex        =   22
      Text            =   "cboDefaultCard"
      Top             =   5400
      Width           =   2115
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "�ָ�ȱʡ(&D)"
      Height          =   405
      Left            =   6465
      TabIndex        =   20
      Top             =   4905
      Width           =   1320
   End
   Begin VB.Frame fraSplitRight 
      Caption         =   "Frame2"
      Height          =   5910
      Left            =   6375
      TabIndex        =   19
      Top             =   -165
      Width           =   30
   End
   Begin VB.ComboBox cboFastkey 
      Height          =   300
      Index           =   2
      Left            =   2310
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1455
      Width           =   1140
   End
   Begin VB.ComboBox cboFun 
      Height          =   300
      Index           =   2
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1455
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������������"
      Height          =   990
      Left            =   90
      TabIndex        =   3
      Top             =   240
      Width           =   6240
      Begin VB.ComboBox cboFastkey 
         Height          =   300
         Index           =   1
         Left            =   5250
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   960
      End
      Begin VB.ComboBox cboFun 
         Height          =   300
         Index           =   1
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
      Begin VB.ComboBox cboFastkey 
         Height          =   300
         Index           =   0
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   345
         Width           =   1155
      End
      Begin VB.ComboBox cboFun 
         Height          =   300
         Index           =   0
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   885
      End
      Begin VB.Label lblFun 
         Caption         =   "���:Ctrl+F4"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   3750
         TabIndex        =   13
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   150
         Index           =   1
         Left            =   5145
         TabIndex        =   11
         Top             =   450
         Width           =   210
      End
      Begin VB.Label lblEdit 
         Caption         =   "������"
         Height          =   210
         Index           =   1
         Left            =   3390
         TabIndex        =   9
         Top             =   405
         Width           =   780
      End
      Begin VB.Label lblFun 
         Caption         =   "���:Ctrl+F4"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   8
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         Height          =   150
         Index           =   0
         Left            =   1980
         TabIndex        =   6
         Top             =   435
         Width           =   210
      End
      Begin VB.Label lblEdit 
         Caption         =   "��ǰ����"
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6645
      TabIndex        =   1
      Top             =   315
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6645
      TabIndex        =   0
      Top             =   750
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3465
      Left            =   120
      TabIndex        =   2
      Top             =   1905
      Width           =   6195
      _cx             =   10927
      _cy             =   6112
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIDKindSet.frx":0000
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
   Begin VB.Label lblDefaultCard 
      Caption         =   "ȱʡ�������"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   5460
      Width           =   1155
   End
   Begin VB.Label lblFun 
      Caption         =   "���:Ctrl+F4"
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   3810
      TabIndex        =   18
      Top             =   1515
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   150
      Index           =   2
      Left            =   2070
      TabIndex        =   16
      Top             =   1545
      Width           =   210
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   14
      Top             =   1500
      Width           =   720
   End
End
Attribute VB_Name = "frmIDKindSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mobjCards As Cards
Private mblnOk As Boolean
Private mstrNotContainFastKey As String
Private mvarNotKey As Variant
Private mstrMustSelectItems As String
Private mRegType As gRegType
Private mblnConn As Boolean
Private mstrPrivs As String
Private mcnOracle As ADODB.Connection

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ���Ч��
    '����: ���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-23 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFind As Boolean, i As Long
    On Error GoTo errHandle
    
    With vsGrid
        For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("����"))) <> 0 And .RowData(i) <> "" Then
                blnFind = True: Exit For
             End If
        Next
    End With
    If Not blnFind Then
        MsgBox "������������һ�����,����", vbInformation + vbOKOnly, gstrSysName
        If vsGrid.Enabled And vsGrid.Visible Then vsGrid.SetFocus
        Exit Function
    End If
    For i = 0 To cboFun.UBound
        If Trim(cboFun(i).Text) <> "" And Trim(cboFastkey(i).Text) = "" Then
            MsgBox lblEdit(i).Caption & "δ���ÿ��,����", vbInformation + vbOKOnly, gstrSysName
            If cboFastkey(i).Enabled And cboFastkey(i).Visible Then cboFastkey(i).SetFocus
            Exit Function
        End If
        
    Next
    '78768
    With vsGrid
        For i = 1 To .Rows - 1
            If cboDefaultCard.Text = .TextMatrix(i, .ColIndex("����")) And .TextMatrix(i, .ColIndex("����")) <> 1 Then
                MsgBox "ѡ�õ�ȱʡ��������ѱ�ͣ�ã����������á�", vbInformation + vbOKOnly, gstrSysName
                'ˢ�¿��б�
                Call InitDefaultCard
            isValied = False: Exit Function
        End If
    Next
End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function


Private Sub SaveParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2018-12-05 14:58:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, strValue As String
    Dim objPubOneCard As Object
    
    On Error GoTo errHandle
    
    Call zlGetPubOneCard(mcnOracle, objPubOneCard)
    strValue = Trim(cboFun(0).Text)
    If mblnConn Then Call objPubOneCard.SetPara("��ǰ����-���ܼ�", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���ܼ�", strValue)
    strValue = Trim(cboFastkey(0).Text)
    If mblnConn Then Call objPubOneCard.SetPara("��ǰ����-���", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���", strValue)
    
    strValue = Trim(cboFun(1).Text)
    If mblnConn Then Call objPubOneCard.SetPara("������-���ܼ�", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "������-���ܼ�", strValue)
     strValue = Trim(cboFastkey(1).Text)
     If mblnConn Then Call objPubOneCard.SetPara("������-���", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "������-���", strValue)
    
    strValue = Trim(cboFun(2).Text)
    If mblnConn Then Call objPubOneCard.SetPara("����-���ܼ�", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "����-���ܼ�", strValue)
    strValue = Trim(cboFastkey(2).Text)
    If mblnConn Then Call objPubOneCard.SetPara("����-���", strValue, glngSys, 1153, True)
    Call SaveRegInFor(mRegType, "ҽ�ƿ����", "����-���", strValue)
    With vsGrid
        For i = 1 To .Rows - 1
            If Left(.RowData(i), 1) = "K" Then
                strValue = IIf(Val(.TextMatrix(i, .ColIndex("����"))) = 0, 0, 1)
                Call SaveRegInFor(mRegType, "ҽ�ƿ����\" & .TextMatrix(i, .ColIndex("����")), "����", strValue)
                '103310:���ϴ�,2016/12/7,���ûس������ӿ��ų���
                strValue = Trim(.TextMatrix(i, .ColIndex("�س���")))
                Call SaveRegInFor(mRegType, "ҽ�ƿ����\" & .TextMatrix(i, .ColIndex("����")), "�س���", strValue)
                strValue = Trim(.TextMatrix(i, .ColIndex("���ܼ�")))
                Call SaveRegInFor(mRegType, "ҽ�ƿ����\" & .TextMatrix(i, .ColIndex("����")), "����-���ܼ�", strValue)
                strValue = Trim(.TextMatrix(i, .ColIndex("���")))
                Call SaveRegInFor(mRegType, "ҽ�ƿ����\" & .TextMatrix(i, .ColIndex("����")), "����-���", strValue)
            End If
        Next
    End With
    '78768:���ϴ�,2014/11/26,ȱʡ�������
    '103309�����ϴ���2016/12/7����ֹ¼��ȱʡ�������
    With cboDefaultCard
        If InStr(1, mstrPrivs, ";��������;") > 0 Then
            If .ListIndex < 0 Then .ListIndex = 0
            Call objPubOneCard.SetPara("ȱʡ�������", .ItemData(.ListIndex), glngSys, 1153, True)
            Call SaveRegInFor(mRegType, "ҽ�ƿ����", "ȱʡ�������", .ItemData(.ListIndex))
        End If
        If Not mblnConn Then Call SaveRegInFor(mRegType, "ҽ�ƿ����", "ȱʡ�������", .ItemData(.ListIndex))
    End With
    Set objPubOneCard = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function InitData(Optional blnRestoreDefault As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 14:59:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strValue As String, strA1 As String
    Dim objPubOneCard As Object
    
    On Error GoTo errHandle
    
    Call zlGetPubOneCard(mcnOracle, objPubOneCard)
 
        
    If Not blnRestoreDefault Then
        '78768:���ϴ�,2014/11/26,��������浽������
        If mblnConn Then
            strA1 = objPubOneCard.getPara("��ǰ����-���ܼ�", glngSys, 1153)
            strValue = objPubOneCard.getPara("��ǰ����-���", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���ܼ�", strA1)
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���", strValue)
        End If
    End If
    
    strA1 = IIf(strA1 = "", "SHIFT", strA1)
    Call SetCboFun(cboFun(0), strA1)
    Call SetCboFastkey(cboFastkey(0), cboFun(0).Text, IIf(strValue = "", "F4", strValue))
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("������-���ܼ�", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "������-���ܼ�", strValue)
        End If
    End If
    Call SetCboFun(cboFun(1), strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("������-���", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "������-���", strValue)
        End If
    End If
    strValue = IIf(strValue = "", "F4", strValue) 'ȱʡF4
    Call SetCboFastkey(cboFastkey(1), cboFun(1).Text, strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("����-���ܼ�", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "����-���ܼ�", strValue)
        End If
    End If
    Call SetCboFun(cboFun(2), strValue)
    
    strValue = ""
    If Not blnRestoreDefault Then
        If mblnConn Then
            strValue = objPubOneCard.getPara("����-���", glngSys, 1153)
        Else
            Call GetRegInFor(mRegType, "ҽ�ƿ����", "����-���", strValue)
        End If
    End If
    strValue = IIf(strValue = "", "�ո��", strValue) 'ȱʡ�ո��
    Call SetCboFastkey(cboFastkey(2), cboFun(2).Text, strValue)
    With vsGrid
        .Clear 1

        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .Rows = mobjCards.Count + 1
        For i = 1 To mobjCards.Count
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "ҽ�ƿ����\" & mobjCards(i).����, "����", strValue)
            End If
            If strValue = "" Then   'ȱʡ����
                .TextMatrix(i, .ColIndex("����")) = 1
            Else
                .TextMatrix(i, .ColIndex("����")) = Val(strValue)
            End If
            
            '103310:���ϴ�,2016/12/7,���ûس������ӿ��ų���
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "ҽ�ƿ����\" & mobjCards(i).����, "�س���", strValue)
            End If
            If strValue = "" Then   'ȱʡ����
                .TextMatrix(i, .ColIndex("�س���")) = "ȱʡ"
            Else
                .TextMatrix(i, .ColIndex("�س���")) = strValue
            End If
            
            strValue = ""
            If Not blnRestoreDefault Then
                Call GetRegInFor(mRegType, "ҽ�ƿ����\" & mobjCards(i).����, "����-���ܼ�", strValue)
            End If
            .TextMatrix(i, .ColIndex("���ܼ�")) = IIf(strValue = "", " ", strValue)
            
            strValue = ""
            If Not blnRestoreDefault Then
                    Call GetRegInFor(mRegType, "ҽ�ƿ����\" & mobjCards(i).����, "����-���", strValue)
            End If
            .TextMatrix(i, .ColIndex("���")) = IIf(strValue = "", " ", strValue)
            .TextMatrix(i, .ColIndex("����")) = mobjCards(i).����
            If InStr(1, "," & mstrMustSelectItems & ",", "," & .TextMatrix(i, .ColIndex("����")) & ",") > 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000011
                .Cell(flexcpForeColor, i, .ColIndex("�س���")) = &H80000008
            End If
            '78768:���ϴ�,2014/11/26,ȱʡ�������
            .TextMatrix(i, .ColIndex("ˢ��")) = IIf(mobjCards(i).�Ƿ�ˢ��, 1, 0)
            .Cell(flexcpData, i, .ColIndex("����")) = mobjCards(i).�ӿ����
            .TextMatrix(i, .ColIndex("ɨ��")) = IIf(mobjCards(i).�Ƿ�ɨ��, 1, 0)
            .TextMatrix(i, .ColIndex("�Ӵ�ʽ����")) = IIf(mobjCards(i).�Ƿ�Ӵ�ʽ����, 1, 0)
            .TextMatrix(i, .ColIndex("�ǽӴ�ʽ����")) = IIf(mobjCards(i).�Ƿ�ǽӴ�ʽ����, 1, 0)
            .RowData(i) = "K" & i
            
        Next
        .Editable = flexEDKbdMouse
    End With
    InitData = True
    Set objPubOneCard = Nothing
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AddCboKey(ByVal cboFun As ComboBox, ByVal strFunKey As String, ByVal strKey As String, strDeult As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӺϷ��Ŀ��
    '���:strKey - String
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-27 16:31:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varTemp As Variant
    For i = 0 To UBound(mvarNotKey)
        varTemp = Split(mvarNotKey(i) & "+", "+")
        If strFunKey = "" Then
            If varTemp(0) = "" And varTemp(1) = strKey Then Exit Sub
        ElseIf strFunKey = "CTRL" Then
            If varTemp(0) = "CTRL" And varTemp(1) = strKey Then Exit Sub
        ElseIf strFunKey = "SHIFT" Then
            If varTemp(0) = "SHIFT" And varTemp(1) = strKey Then Exit Sub
        End If
    Next
    With cboFun
        .AddItem strKey
        If strKey = strDeult Then .ListIndex = .NewIndex
    End With
End Sub

Private Sub SetCboFastkey(ByVal cboFun As ComboBox, ByVal strFunKey As String, ByVal strDefaultValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ��
    '����:���˺�
    '����:2012-08-22 15:54:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFastkey As String
    Dim i As Long
    With cboFun
        .Clear
        .AddItem " "
        Call AddCboKey(cboFun, strFunKey, "�ո��", strDefaultValue)
    End With
    For i = 1 To 12
        Call AddCboKey(cboFun, strFunKey, "F" & i, strDefaultValue)
    Next
    Call AddCboKey(cboFun, strFunKey, "��", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "��", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "��", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "��", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "<", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, ">", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "[", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "]", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "/", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "\", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "*", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "`", strDefaultValue)
        
    Call AddCboKey(cboFun, strFunKey, "+", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "-", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "=", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "(", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, ")", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "?", strDefaultValue)
    For i = Asc("A") To Asc("Z")
        Call AddCboKey(cboFun, strFunKey, Chr(i), strDefaultValue)
    Next
    For i = Asc("0") To Asc("9")
        Call AddCboKey(cboFun, strFunKey, Chr(i), strDefaultValue)
    Next
        
    '���ּ���
    For i = Asc("0") To Asc("9")
        Call AddCboKey(cboFun, strFunKey, "NUM" & Chr(i), strDefaultValue)
    Next
    Call AddCboKey(cboFun, strFunKey, "NUM*", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM+", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM-", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM/", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUM.", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, "NUMENTER", strDefaultValue)
    Call AddCboKey(cboFun, strFunKey, UCase("KeyLButton"), strDefaultValue)      'vbKeyLButton 0x1 ������
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMButton"), strDefaultValue)     'vbKeyRButton 0x2 ����Ҽ�'
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMButton"), strDefaultValue)     'vbKeyMButton 0x4 ����м�
    Call AddCboKey(cboFun, strFunKey, UCase("KeyCancel"), strDefaultValue)     ' vbKeyCancel 0x3 CANCEL ��
    Call AddCboKey(cboFun, strFunKey, UCase("KeyBack"), strDefaultValue)     ' vbKeyBack 0x8 BACKSPACE ��
    Call AddCboKey(cboFun, strFunKey, UCase("KeyTab"), strDefaultValue)     ' vbKeyTab 0x9 TAB ��
    Call AddCboKey(cboFun, strFunKey, UCase("KeyClear"), strDefaultValue)     ' vbKeyClear 0xC CLEAR ��
    Call AddCboKey(cboFun, strFunKey, UCase("ENTER"), strDefaultValue)     ' vbKeyReturn 0xD ENTER ��
    Call AddCboKey(cboFun, strFunKey, UCase("SHIFT"), strDefaultValue)     ' vbKeyShift 0x10 SHIFT ��
    Call AddCboKey(cboFun, strFunKey, UCase("CTRL"), strDefaultValue)     ' vbKeyControl 0x11 CTRL ��
    Call AddCboKey(cboFun, strFunKey, UCase("KeyMenu"), strDefaultValue)     ' vbKeyMenu 0x12 MENU ��
    Call AddCboKey(cboFun, strFunKey, UCase("KeyPause"), strDefaultValue)     ' vbKeyPause 0x13 PAUSE ��
    Call AddCboKey(cboFun, strFunKey, UCase(" CAPS LOCK"), strDefaultValue)     ' vbKeyCapital 0x14 CAPS LOCK ��
    Call AddCboKey(cboFun, strFunKey, UCase("ESC"), strDefaultValue)     ' vbKeyEscape 0x1B ESC ��
    Call AddCboKey(cboFun, strFunKey, UCase("SELECT"), strDefaultValue)     ' vbKeySelect 0x29 SELECT ��
    Call AddCboKey(cboFun, strFunKey, UCase("PRINT SCREEN"), strDefaultValue)     ' vbKeyPrint 0x2A PRINT SCREEN ��
    Call AddCboKey(cboFun, strFunKey, UCase("EXECUTE"), strDefaultValue)     ' vbKeyExecute 0x2B EXECUTE ��
    Call AddCboKey(cboFun, strFunKey, UCase("SNAPSHOT"), strDefaultValue)     ' vbKeySnapshot 0x2C SNAPSHOT ��
    Call AddCboKey(cboFun, strFunKey, UCase("INSERT"), strDefaultValue)    ' vbKeyInsert 0x2D INSERT ��
    Call AddCboKey(cboFun, strFunKey, UCase("DELETE"), strDefaultValue)     ' vbKeyDelete 0x2E DELETE ��
    Call AddCboKey(cboFun, strFunKey, UCase("HELP"), strDefaultValue)     ' vbKeyHelp 0x2F HELP ��
    Call AddCboKey(cboFun, strFunKey, UCase("NUM LOCK"), strDefaultValue)     ' vbKeyNumlock 0x90 NUM LOCK ��
    If cboFun.ListIndex < 0 Then cboFun.ListIndex = 0
End Sub

Private Sub SetCboFun(ByVal cboFun As ComboBox, ByVal strDefaultValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù��ܼ�
    '����:���˺�
    '����:2012-08-22 15:54:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With cboFun
        .Clear
        .AddItem " "
        If strDefaultValue = "" Then .ListIndex = .NewIndex
        .AddItem "CTRL"
        If strDefaultValue = "CTRL" Then .ListIndex = .NewIndex
        .AddItem "SHIFT"
        If strDefaultValue = "SHIFT" Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub

Public Function ShowSetWin(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, _
    ByVal objCards As Cards, ByVal RegType As gRegType, _
    Optional strNotContainFastKey As String = "", _
    Optional strMustSelectItems As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û�ͣ�ÿ����
    '���:strNotContainFastKey-���ܰ����Ŀ��
    '����:���õ��ȷ����,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-25 23:20:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjCards = objCards: mstrNotContainFastKey = strNotContainFastKey
    Set mcnOracle = cnOracle
    
    mstrMustSelectItems = strMustSelectItems
    mRegType = RegType
    mblnOk = False
    If frmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmMain
    End If
    ShowSetWin = mblnOk
End Function

Private Sub cboDefaultCard_KeyPress(KeyAscii As Integer)
    '����������
    KeyAscii = 0
End Sub

Private Sub cboFastkey_Click(Index As Integer)
    lblFun(Index).Caption = "���:" & cboFun(Index).Text & IIf(Trim(cboFun(Index).Text) = "", " ", " + ") & cboFastkey(Index).Text
End Sub

Private Sub cboFun_Click(Index As Integer)
    lblFun(Index).Caption = "���:" & cboFun(Index).Text & IIf(Trim(cboFun(Index).Text) = "", " ", " + ") & cboFastkey(Index).Text
    Call SetCboFastkey(cboFastkey(Index), cboFun(Index).Text, cboFastkey(Index).Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefault_Click()
   Call InitData(True)
   If isValied = False Then Exit Sub
  Call SaveParaSet
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    Call SaveParaSet
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    Call InitDefaultCard
    If vsGrid.Enabled And vsGrid.Visible Then vsGrid.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        Unload Me: Exit Sub
    End If
End Sub
Private Sub Form_Load()
    Dim objPubOneCard As Object
    
    mblnFirst = True
    mvarNotKey = Split(mstrNotContainFastKey, ";")
    mblnConn = Not gcnOracle Is Nothing
    Call InitFace
    If mblnConn Then
        Call zlGetPubOneCard(mcnOracle, objPubOneCard)
        mstrPrivs = ";" & objPubOneCard.GetPrivFunc(glngSys, 1153) & ";"
        Set objPubOneCard = Nothing
    End If
    
End Sub

Private Sub InitFace()
    With vsGrid
        .Clear 1
        .ColComboList(.ColIndex("�س���")) = "ȱʡ|����|����"
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("����"), .ColIndex("���ܼ�"), .ColIndex("���")
            If Left(.RowData(Row), 1) <> "K" Then Cancel = True: Exit Sub
            If InStr(1, "," & mstrMustSelectItems & ",", "," & .TextMatrix(Row, .ColIndex("����")) & ",") > 0 Then Cancel = True
        Case .ColIndex("�س���")
            If Left(.RowData(Row), 1) <> "K" Then Cancel = True: Exit Sub
            If mobjCards(Row).�ӿ���� <= 0 Or Not (mobjCards(Row).�Ƿ�ˢ�� Or mobjCards(Row).�Ƿ�ɨ��) Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub InitDefaultCard()
'��ȱʡ�������ͱ����ڲ������ע���
    Dim lngDefaultCardID As Long, strValue As String, i As Integer, j As Integer
    Dim objPubOneCard As Object
    If mblnConn Then
        Call zlGetPubOneCard(mcnOracle, objPubOneCard)
        strValue = objPubOneCard.getPara("ȱʡ�������", glngSys, 1153, 0, Array(lblDefaultCard, cboDefaultCard), InStr(1, mstrPrivs, ";��������;") > 0)
    Else
        Call GetRegInFor(mRegType, "ҽ�ƿ����", "ȱʡ�������", strValue)
    End If
    lngDefaultCardID = Val(strValue)
    cboDefaultCard.Clear
    cboDefaultCard.AddItem ""
    cboDefaultCard.ItemData(cboDefaultCard.NewIndex) = 0
    cboDefaultCard.ListIndex = cboDefaultCard.NewIndex
    With vsGrid
        For i = 1 To .Rows - 1
            '85565,���ϴ�,2015/7/10:��������
            'If .TextMatrix(i, .ColIndex("����")) <> 0 And .TextMatrix(i, .ColIndex("ˢ��")) = 0 Then
            If .TextMatrix(i, .ColIndex("����")) <> 0 And (.TextMatrix(i, .ColIndex("�Ӵ�ʽ����")) = 1 Or .TextMatrix(i, .ColIndex("�ǽӴ�ʽ����")) = 1) Then
                cboDefaultCard.AddItem .TextMatrix(i, .ColIndex("����"))
                cboDefaultCard.ItemData(cboDefaultCard.NewIndex) = .Cell(flexcpData, i, .ColIndex("����"))
                If lngDefaultCardID = .Cell(flexcpData, i, .ColIndex("����")) Then cboDefaultCard.ListIndex = cboDefaultCard.NewIndex
            End If
        Next
    End With
End Sub

Private Sub vsGrid_Validate(Cancel As Boolean)
    If cboDefaultCard.Enabled And cboDefaultCard.Visible Then
        Call InitDefaultCard
    End If
End Sub
