VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBlackListTypeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϊ����"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7440
   Icon            =   "frmBlackListTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7440
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   195
      Width           =   4890
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "��Ч�ڼ�"
         Top             =   1095
         Width           =   600
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "����"
         Top             =   1095
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "����"
         Top             =   705
         Width           =   3675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "����"
         Top             =   345
         Width           =   1500
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   4200
         TabIndex        =   8
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���(&U)"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   405
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��Ч�ڼ�(&Q)"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   6
         Top             =   1155
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5235
      TabIndex        =   10
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   11
      Top             =   765
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Height          =   120
      Left            =   1155
      TabIndex        =   12
      Top             =   1935
      Width           =   6420
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGridRule 
      Height          =   2835
      Left            =   60
      TabIndex        =   9
      Top             =   2235
      Width           =   7305
      _cx             =   12885
      _cy             =   5001
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBlackListTypeEdit.frx":06EA
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
   Begin VB.Label Label1 
      Caption         =   "������Ϊ����"
      Height          =   270
      Left            =   45
      TabIndex        =   14
      Top             =   1935
      Width           =   1170
   End
End
Attribute VB_Name = "frmBlackListTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gTypeEdit
    EM_Ty_���� = 0
    EM_Ty_�޸�
    EM_Ty_ɾ��
    EM_Ty_�������
    EM_Ty_�鿴
End Enum
Private mbytEditType As gTypeEdit
Private mfrmMain As Object
Private mstrCode As String
Private mblnChange As Boolean     '�Ƿ�ı���
Private mintSuccess As Integer
Private mblnFirst As Boolean
Private mblnSys As Boolean
Private mblnUnLoad As Boolean

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytEditType As gTypeEdit, Optional strCode As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�༭������Ϊ���
    '���:frmMain-���õ�������
    '    bytEditType-�༭���:0-����;1-�޸�;2-���޸Ŀ��Ʒ�ʽ;3-�鿴;
    '     strCode-����,����ʱ�����
    '����:�༭�ɹ�����True,����ΪFalse
    '����:���˺�
    '����:2018-11-08 17:01:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytEditType = bytEditType: mintSuccess = 0
    Set mfrmMain = frmMain: mstrCode = strCode: mblnFirst = True
    mblnUnLoad = False
    
    Me.Show 1, frmMain
    zlShowEdit = mintSuccess > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetInputDefineSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ؿؼ����볤�ȣ��õ����ݿ�ı��ֶεĳ��ȣ�
    '����:���˺�
    '����:2018-11-09 17:06:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "SELECT ����,����,����,��Ч���� FROM ������Ϊ���� Where Rownum<0 "
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "������Ϊ����")
    
    txtEdit(1).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(2).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(3).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(4).MaxLength = rsTemp.Fields("��Ч����").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub SetCtrolEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ��Ƶ�enable����
    '����:���˺�
    '����:2018-11-13 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, i As Long
    On Error GoTo errHandle
    blnEdit = (mbytEditType = EM_Ty_���� Or mbytEditType = EM_Ty_�޸�) And mblnSys = False
    txtEdit(1).Enabled = mbytEditType = EM_Ty_����
    txtEdit(2).Enabled = blnEdit
    txtEdit(3).Enabled = blnEdit
    txtEdit(4).Enabled = mbytEditType = EM_Ty_���� Or mbytEditType = EM_Ty_�޸�
    
    For i = 1 To txtEdit.UBound
        txtEdit(i).BackColor = IIf(txtEdit(i).Enabled, &H80000005, &H8000000F)
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

 
Private Function ReadData(ByVal strCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ����ȡ����
    '���:strCode-��ǰ����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 17:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
   
    mblnSys = False
    If mbytEditType = 0 Then
        '����
        txtEdit(1).Text = zlDatabase.GetMax("������Ϊ����", "����", txtEdit(1).MaxLength)
        Call LoadRuleData("")
        Call SetCtrolEnabled
        ReadData = True
        Exit Function
    End If
     
    strSQL = "" & _
    "   SELECT ����,����,����,��Ч���� ,�Ƿ�̶� FROM ������Ϊ����  Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
            
    If rsTemp.EOF Then
        MsgBox "δ�ҵ�����Ϊ��" & strCode & "���Ĳ�����Ϊԭ�����ݣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    txtEdit(1).Text = Nvl(rsTemp!����)
    txtEdit(2).Text = Nvl(rsTemp!����)
    txtEdit(3).Text = Nvl(rsTemp!����)
    txtEdit(4).Text = IIf(Val(Nvl(Nvl(rsTemp!��Ч����))) = 0, "", Val(Nvl(Nvl(rsTemp!��Ч����))))
    
    mblnSys = Val(Nvl(rsTemp!�Ƿ�̶�)) = 1
    
    '���ؿ��ƹ���
    Call LoadRuleData(Nvl(rsTemp!����))
    Call SetCtrolEnabled
     
    ReadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetSplitRuleValue(ByVal str���ƹ��� As String, Optional str���Ʒ�_out As String, Optional str����_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݿ��ƹ���ֵ,����ָ������
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-12 15:22:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���Ʒ� As String, str���� As String, varTemp As Variant
    
    On Error GoTo errHandle
    If InStr(1, str���ƹ���, ">=") > 0 Then
        varTemp = Split(str���ƹ���, ">=")
        str���Ʒ� = ">="
        str���� = Val(varTemp(1))
    ElseIf InStr(1, str���ƹ���, ">") > 0 Then
        varTemp = Split(str���ƹ���, ">")
        str���Ʒ� = ">"
        str���� = Val(varTemp(1))
    ElseIf InStr(1, str���ƹ���, "<=") > 0 Then
        varTemp = Split(str���ƹ���, "<=")
        str���Ʒ� = "<="
        str���� = Val(varTemp(1))
    ElseIf InStr(1, str���ƹ���, "<") > 0 Then
        varTemp = Split(str���ƹ���, "<")
        str���Ʒ� = "<"
        str���� = Val(varTemp(1))
    ElseIf InStr(1, str���ƹ���, "=") > 0 Then
        varTemp = Split(str���ƹ���, "=")
        str���Ʒ� = "="
        str���� = Val(varTemp(1))
    ElseIf InStr(1, str���ƹ���, "-") > 0 Then
        str���Ʒ� = "���η�Χ"
        str���� = str���ƹ���
    Else
        Exit Function
    End If
    str���Ʒ�_out = str���Ʒ�
    str����_Out = str����
    GetSplitRuleValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetRuleRowFromRuleValue(ByVal str���Ʒ� As String, ByVal str���� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ��򣬻�ȡָ������
    '���:str���Ʒ�-���Ʒ�:>=;<=��
    '     str����
    '����:�ҵ�����ָ�����У����򷵻�-1
    '����:���˺�
    '����:2018-11-12 15:27:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo errHandle
    With vsGridRule
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���Ʒ�")) = str���Ʒ� And .TextMatrix(i, .ColIndex("����")) = str���� Then
                GetRuleRowFromRuleValue = i: Exit Function
            End If
        Next
    End With
    GetRuleRowFromRuleValue = -1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadRuleData(ByVal str��Ϊ��� As String) As Boolean  '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ϊ�������Ӧ�Ŀ��ƹ���
    '����:���˺�
    '����:2018-11-12 14:35:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllType As Collection, blnFind As Boolean
    Dim varTemp As Variant, strTemp As String, lngRow As Long
    Dim str���Ʒ� As String, str���� As String
    Dim rsTemp As ADODB.Recordset, rsԤԼ��ʽ As ADODB.Recordset
    On Error GoTo errHandle
    
     
    strSQL = "Select ����,���� From ԤԼ��ʽ order by ����"
    Set rsԤԼ��ʽ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set cllType = New Collection
    With rsԤԼ��ʽ
        Do While Not .EOF
            cllType.Add Nvl(!����)
            .MoveNext
        Loop
    End With
    
    
    strSQL = "" & _
    "   Select a.Ӧ�ó���,a.��Ϊ���,a.ԤԼ��ʽ,a.���,a.���ƹ���,a.���Ʒ�ʽ  " & _
    "   From ������Ϊ���� A" & _
    "   where ��Ϊ���=[1]" & _
    "   Order by ��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str��Ϊ���)
   
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!ԤԼ��ʽ) <> "" Then
                blnFind = False
                rsԤԼ��ʽ.Filter = "����='" & !ԤԼ��ʽ & "'"
                If rsԤԼ��ʽ.EOF Then
                    cllType.Add Nvl(!ԤԼ��ʽ)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    vsGridRule.Redraw = flexRDNone
    Call InitRuleGridColumHead(cllType)
    
    rsTemp.Sort = "���"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsGridRule
        .Clear 1
        Do While Not rsTemp.EOF
            strTemp = Nvl(rsTemp!���ƹ���)
            If GetSplitRuleValue(strTemp, str���Ʒ�, str����) Then
                lngRow = GetRuleRowFromRuleValue(str���Ʒ�, str����)
                If lngRow = -1 Then
                    If .TextMatrix(.Rows - 1, .ColIndex("���Ʒ�")) <> "" Then .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    
                    .TextMatrix(lngRow, .ColIndex("���Ʒ�")) = str���Ʒ�
                    .TextMatrix(lngRow, .ColIndex("����")) = str����
                End If
                If Nvl(rsTemp!Ӧ�ó���) = "ԤԼ" Then
                    If Trim(Nvl(rsTemp!ԤԼ��ʽ)) = "" Then
                        .TextMatrix(lngRow, .ColIndex("����ԤԼ")) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
                    Else
                        .TextMatrix(lngRow, .ColIndex(Trim(Nvl(rsTemp!ԤԼ��ʽ)))) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
                    End If
                Else
                     .TextMatrix(lngRow, .ColIndex(Trim(Nvl(rsTemp!Ӧ�ó���)))) = decode(Val(Nvl(rsTemp!���Ʒ�ʽ)), 1, "��ֹ", 2, "��ʾ", "")
                End If
            End If
           rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadRuleData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
 
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    If mbytEditType <> 0 Then
        mblnChange = False: Unload Me
        Exit Sub
    End If
    
    mstrCode = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = zlDatabase.GetMax("������Ϊ����", "����", txtEdit(1).MaxLength)
    '�������ϴεĲ���
    
    mblnChange = False
    txtEdit(1).SetFocus
End Sub

Private Function IsValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������й������Ƿ���Ч
    '����:��Ч����True,����ΪFalse
    '����:���˺�
    '����:2018-11-09 17:22:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, varTemp As Variant, varData As Variant, blnHaveData As Boolean
    Dim strTemp As String
    
    On Error GoTo errHandle
    For i = 1 To 3
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox txtEdit(i).Tag & "���ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox txtEdit(i).Tag & "�к��зǷ��ַ���", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    If IsNumeric(txtEdit(4).Text) = False And txtEdit(4).Text <> "" Then
        MsgBox "��Ч������������������͡�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    If Val(txtEdit(4).Text) > 99999 Then
        MsgBox "��Ч�������ֻ������99999���¡�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    
    If Val(txtEdit(4).Text) < 0 Then
        MsgBox "��Ч�������������ڵ���0���¡�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    
    If Len(txtEdit(1).Text) = 0 Or Trim(txtEdit(1).Text) = "" Then
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(1)
        Exit Function
    End If
    
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(2)
        Exit Function
    End If
    With vsGridRule
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("���Ʒ�"))) <> "" Then
                If Trim(.TextMatrix(i, .ColIndex("����"))) = "" Then
                    MsgBox "��" & i & "��δ��������!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = .ColIndex("����")
                    .SetFocus
                    Exit Function
                End If
                If .TextMatrix(i, .ColIndex("���Ʒ�")) = "���η�Χ" Then
                    If InStr(.TextMatrix(i, .ColIndex("����")), "-") = 0 Then
                        MsgBox "��" & i & "�е����θ�ʽ����ȷ��Ӧ������ĸ�ʽΪ:000-999!", vbInformation + vbOKOnly, gstrSysName
                        .Row = i: .Col = .ColIndex("����")
                        .SetFocus
                        Exit Function
                    End If
                End If
                blnHaveData = False
                For j = .ColIndex("����") + 1 To .Cols - 1
                    If .TextMatrix(i, .ColIndex("����ԤԼ")) = "��ֹ" Then
                       If InStr(";����ԤԼ;�Һ�;����;��Ժ;��Ժ;����;", ";" & .ColKey(j) & ";") = 0 And .TextMatrix(0, j) = "ԤԼ��ʽ" Then
                            If .TextMatrix(i, j) <> "��ֹ" And j <> .ColIndex("����ԤԼ") And Trim(.TextMatrix(i, j)) <> "" Then
                                MsgBox "��" & i & "�е�����ԤԼҵ���ǽ�ֹ״̬������ԤԼҲӦ��Ϊ��ֹ������!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = j
                                .SetFocus
                                Exit Function
                            End If
                       End If
                    End If
                                    
                    If .TextMatrix(i, j) <> "" Then blnHaveData = True
                Next
                If Not blnHaveData Then
                    MsgBox "��" & i & "��δ������صĿ��Ʒ�ʽ!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = .ColIndex("����")
                    .SetFocus
                    Exit Function
                End If
           
                For j = i + 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���Ʒ�")) = .TextMatrix(j, .ColIndex("���Ʒ�")) Then
                        If .TextMatrix(i, .ColIndex("���Ʒ�")) <> "���η�Χ" Then
                            If Val(.TextMatrix(i, .ColIndex("����"))) = Val(.TextMatrix(j, .ColIndex("����"))) Then
                                MsgBox "��" & i & "���ƶ��Ĺ������" & j & "�еĹ�����ͬ����ϲ�!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = .ColIndex("���Ʒ�")
                                .SetFocus
                                Exit Function
                            End If
                        Else
                            If .TextMatrix(i, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) Then
                                If Val(.TextMatrix(i, .ColIndex("����"))) = Val(.TextMatrix(j, .ColIndex("����"))) Then
                                    MsgBox "��" & i & "���ƶ��Ĺ������" & j & "�еĹ�����ͬ����ϲ�!", vbInformation + vbOKOnly, gstrSysName
                                    .Row = i: .Col = .ColIndex("���Ʒ�")
                                    .SetFocus
                                    Exit Function
                                End If
                            End If
                            varData = Split(.TextMatrix(i, .ColIndex("����")) & "-", "-")
                            varTemp = Split(.TextMatrix(j, .ColIndex("����")) & "-", "-")
                            If Val(varData(0)) = Val(varTemp(0)) And Val(varData(1)) = Val(varTemp(1)) Then
                                MsgBox "��" & i & "���ƶ��Ĺ������" & j & "�еĹ�����ͬ����ϲ�!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = .ColIndex("���Ʒ�")
                                .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                Next
                If CheckRuleDataValid(i) = False Then
                    MsgBox "��" & i & "��δ������Ƴ��ϣ�����!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = 2
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 17:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Long, cllPro As Collection, strSQL As String, blnDel As Boolean
    Dim blnTran As Boolean, strTemp As String, str���� As String, strRule As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    
    If mbytEditType <> EM_Ty_������� Then
        
        '    Zl_������Ϊ����_Update
        strSQL = "Zl_������Ϊ����_Update("
        '  ����_In     Number, 0-����;1-�޸�
        strSQL = strSQL & "" & IIf(mbytEditType = 0, 0, 1) & ","
        '  ����_In     ������Ϊ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  ����_In     ������Ϊ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  ����_In     ������Ϊ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(3).Text & "',"
        '  �Ƿ�̶�_In ������Ϊ����.�Ƿ�̶�%Type := 0,
        strSQL = strSQL & "" & IIf(mblnSys, "1", "0") & ","
        '  ��Ч����_In ������Ϊ����.��Ч����%Type := Null
        strSQL = strSQL & "" & IIf(Val(txtEdit(4).Text) = 0, "NULL", Val(txtEdit(4).Text)) & ")"
        zlAddArray cllPro, strSQL
        
    End If
    'ȡ���򱣴�
    blnDel = True
    With vsGridRule
        strTemp = "": str���� = ""
        For i = .FixedRows To .Rows - 1
            
            '����1|���Ʒ�ʽ1||����2|���Ʒ�ʽ1||....
            '  --                 ����:��:>=10;<10��
            '  --                 ���Ʒ�ʽ:��ʽΪӦ�ó���:���Ʊ�־(0-������(�����Ƶģ���ʼ���룬Ҳ������);1-��ֹ;2-����):ԤԼ��ʽ
            
            If Trim(.TextMatrix(i, .ColIndex("���Ʒ�"))) <> "" Then
                strTemp = Trim(.TextMatrix(i, .ColIndex("���Ʒ�")))
                If strTemp = "���η�Χ" Then
                    strTemp = .TextMatrix(i, .ColIndex("����"))
                Else
                    strTemp = strTemp & Val(.TextMatrix(i, .ColIndex("����")))
                End If
                  
                strRule = strTemp
                For j = .ColIndex("����") + 1 To .Cols - 1
                   
                   If Trim(.TextMatrix(i, j)) <> "" Then
                        'strRule = strTemp
                        Select Case .ColKey(j)
                        Case "����ԤԼ"
                            strRule = strRule & "|" & "ԤԼ:" & IIf(.TextMatrix(i, j) = "��ֹ", 1, 2)
                        Case "�Һ�", "��Ժ", "��Ժ", "����"
                            strRule = strRule & "|" & .ColKey(j) & ":" & IIf(.TextMatrix(i, j) = "��ֹ", 1, 2)
                        Case Else
                            strRule = strRule & "|" & "ԤԼ:" & IIf(.TextMatrix(i, j) = "��ֹ", 1, 2) & ":" & .ColKey(j)
                        End Select
                    End If
                Next
                If strRule <> "" Then
                    If zlCommFun.ActualLen(str���� & "||" & strRule) > 4000 Then
                        str���� = Mid(str����, 3)
                        'Zl_������Ϊ���ƹ���_Update
                        strSQL = "Zl_������Ϊ���ƹ���_Update("
                        '  ��Ϊ���_In ������Ϊ����.��Ϊ���%Type,
                        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
                         '  ���ƹ���_In Varchar2,
                        strSQL = strSQL & "'" & str���� & "',"
                        '  �Ƿ�ɾ��_In Number:=1
                        strSQL = strSQL & "" & IIf(blnDel, 1, 0) & ")"
                        zlAddArray cllPro, strSQL
                        blnDel = False
                        str���� = ""
                    End If
                    str���� = str���� & "||" & strRule
                End If
            End If
       Next
    End With
    
    str���� = Mid(str����, 3)
    'Zl_������Ϊ���ƹ���_Update
    strSQL = "Zl_������Ϊ���ƹ���_Update("
    '  ��Ϊ���_In ������Ϊ����.��Ϊ���%Type,
    strSQL = strSQL & "'" & txtEdit(2).Text & "',"
     '  ���ƹ���_In Varchar2,
    strSQL = strSQL & "'" & str���� & "',"
    '  �Ƿ�ɾ��_In Number:=1
    strSQL = strSQL & "" & IIf(blnDel, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    
    blnTran = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    If txtEdit(2).Enabled And txtEdit(2).Visible Then txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is vsGridRule Then Exit Sub
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    
    Call SetInputDefineSize '����ȱʡ�����볤��
    mblnUnLoad = Not ReadData(mstrCode) '��ȡ����
    
    mblnChange = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
     

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    If Index = 4 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m����ʽ
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub InitRuleGridColumHead(ByVal cllType As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ƹ���������ͷ
    '����:���˺�
    '����:2018-11-08 18:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    
    On Error GoTo errHandle
    With vsGridRule
        .Clear:
        .Rows = .FixedRows + 1
        .Cols = cllType.Count + 7
        i = 0
        
        .TextMatrix(0, i) = "���ƹ���"
        .TextMatrix(1, i) = "���Ʒ�": .ColWidth(i) = 800: i = i + 1
        
        .TextMatrix(0, i) = "���ƹ���"
        .TextMatrix(1, i) = "����": .ColWidth(i) = 800: i = i + 1
        
        .TextMatrix(0, i) = "����ԤԼ"
        .TextMatrix(1, i) = "����ԤԼ": .ColWidth(i) = 800: i = i + 1
        For j = 1 To cllType.Count
            .TextMatrix(0, i) = "ԤԼ��ʽ"
            .TextMatrix(1, i) = cllType(j): .ColWidth(i) = 800
            If Me.TextWidth(" " & cllType(j)) > 800 Then
                 .ColWidth(i) = Me.TextWidth(" " & cllType(j))
            End If
            i = i + 1
        Next
        
        .TextMatrix(0, i) = "�Һ�":
        .TextMatrix(1, i) = "�Һ�": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "��Ժ":
        .TextMatrix(1, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "��Ժ"
        .TextMatrix(1, i) = "��Ժ": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "����"
        .TextMatrix(1, i) = "����": .ColWidth(i) = 800: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(1, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            Select Case .ColKey(i)
            Case "����"
                .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .MergeCol(i) = True
        Next
        .ColComboList(.ColIndex("���Ʒ�")) = " |>=|>|=|<=|<|���η�Χ"
        
        .MergeCells = flexMergeRestrictAll
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True
        .MergeRow(1) = True
        .Editable = IIf(mbytEditType = EM_Ty_�鿴 Or mbytEditType = EM_Ty_ɾ��, flexEDNone, flexEDKbdMouse) '0-����;1-�޸�;2-���޸Ŀ��Ʒ�ʽ;3-�鿴;
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub DeleteRuleRow(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��ָ���Ĺ�����
    '���:lngRow-ָ������
    '����:���˺�
    '����:2018-11-12 16:54:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsGridRule
    
        If lngRow > .Rows - 1 Or lngRow < .FixedRows Then Exit Sub
        If lngRow = .FixedRows And lngRow = .Rows - 1 Then
            .Clear 1
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = ""
            Exit Sub
        End If
        If lngRow < .Rows - 1 Then
            .RemoveItem lngRow
            .Row = lngRow
            Exit Sub
        End If
        .RemoveItem lngRow
        .Row = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsGridRule_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim varData As Variant
    
    If mbytEditType = gTypeEdit.EM_Ty_�鿴 Then Exit Sub
    
    With vsGridRule
        Select Case Col
        Case .ColIndex("ɾ��")
             Call DeleteRuleRow(Row)
        Case Else
        End Select
    End With
End Sub

 
Private Sub vsGridRule_ChangeEdit()
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    With vsGridRule
       Select Case .Col
       Case .ColIndex("����")
       Case Else
       End Select
    End With
End Sub

Private Sub vsGridRule_DblClick()
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    With vsGridRule
        If .Row < 0 Then Exit Sub
        Select Case .Col
        Case .ColIndex("����")
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        Case .ColIndex("���Ʒ�")
        Case .ColIndex("ɾ��")
        
        Case Else
            If .TextMatrix(.Row, .Col) = "��ֹ" Then
                .TextMatrix(.Row, .Col) = "��ʾ"
            ElseIf .TextMatrix(.Row, .Col) = "��ʾ" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "��ֹ"
            End If
        End Select
    End With
End Sub
 

Private Sub vsGridRule_EnterCell()
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
   
    With vsGridRule
        If .Row < 0 Then Exit Sub
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        Select Case .Col
        Case .ColIndex("����")
        End Select
    End With
End Sub

Private Sub vsGridRule_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varTemp As Variant, lngRow As Long
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    
    With vsGridRule
        If .Row > .Rows - 1 Or .Row < 1 Then Exit Sub
    
         If KeyCode <> vbKeyReturn And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                Call vsGridRule_CellButtonClick(.Row, .Col)
            Else
                If .Col = .ColIndex("���Ʒ�") Then .ColComboList(.Col) = ""
            End If
        End If
        'ɾ��
        If KeyCode = vbKeyDelete Then
            Call vsGridRule_CellButtonClick(.Row, .Col)
            Exit Sub
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
     
    With vsGridRule
        Select Case .Col
        Case .ColIndex("����")
            If (Trim(.TextMatrix(.Row, .ColIndex("����"))) = "" And Trim(.TextMatrix(.Row, .ColIndex("���Ʒ�"))) = "" <> "") And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("����"))) = "" And Trim(.TextMatrix(.Row, .ColIndex("���Ʒ�"))) = "" <> "") And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
            Call zlVsMoveGridCell(vsGridRule, .ColIndex("���Ʒ�"), , IIf(mbytEditType = EM_Ty_�鿴, False, True), lngRow)
    End With
    Call vsGridRule_EnterCell
End Sub
 


Private Sub vsGridRule_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

    Dim strKey As String, lngRow As Long
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsGridRule
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '�ݲ���������
        Select Case Col
        Case .ColIndex("���Ʒ�")
        Case .ColIndex("����")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsGridRule, .ColIndex("���Ʒ�"), -1, True, lngRow)
    End With
End Sub

Private Sub vsGridRule_KeyPress(KeyAscii As Integer)
 
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
   
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    
    With vsGridRule
        If .Col <> .ColIndex("����") Then KeyAscii = 0: Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsGridRule, vsGridRule.Row, vsGridRule.Col, KeyAscii, m����ʽ)
End Sub

Private Sub vsGridRule_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngCashRow As Long
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
     
    With vsGridRule
        Select Case .Col
        Case .ColIndex("����")
            Call VsFlxGridCheckKeyPress(vsGridRule, Row, Col, KeyAscii, m����ʽ)
        End Select
    End With
End Sub


Private Sub vsGridRule_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsGridRule
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '�ݲ���������
        Select Case Col
        Case .ColIndex("���Ʒ�")
            .Col = .ColIndex("����")
        Case Else
            'Call zlVsMoveGridCell(vsGridRule, .ColIndex("����"), -1, True, Row)
        End Select

    End With
   
End Sub
Private Sub vsGridRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error GoTo Errhand:
    With vsGridRule
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
       If .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol)) Then Exit Sub
       .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol))
    End With
Errhand:
    Exit Sub
End Sub

Private Sub vsGridRule_LeaveCell()
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    OS.OpenIme False
End Sub



Private Sub vsGridRule_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 
    
    If mbytEditType = EM_Ty_�鿴 Then Exit Sub
    
    '���õ�Ԫ��ı༭����
    With vsGridRule
       Select Case .Col
           Case .ColIndex("����")
               .EditMaxLength = 50
       End Select
    End With
End Sub

Private Sub vsGridRule_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String, varTemp As Variant
    
    With vsGridRule
        If Row <= 0 Then Exit Sub
        
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        
        Select Case Col
        Case .ColIndex("���Ʒ�")
        Case .ColIndex("����")
            If .TextMatrix(Row, .ColIndex("���Ʒ�")) <> "���η�Χ" And .TextMatrix(Row, .ColIndex("���Ʒ�")) <> "" Then
                If Not IsNumeric(strInput) And strInput <> "" Then
                    MsgBox "��������α���Ϊ���֣�", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                If zlDblIsValid(strInput, 5, False, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
            ElseIf .TextMatrix(Row, .ColIndex("���Ʒ�")) <> "" Then
                If InStr(strInput, "-") = 0 Then
                     MsgBox "��������α�����ϸ�ʽ(XXXXX-XXXX)�ķ�Χ��ʽ,���磺1-5��", vbInformation, gstrSysName
                     Cancel = True: Exit Sub
                End If
                varTemp = Split(strInput, "-")
                If Val(varTemp(0)) > Val(varTemp(1)) Then
                     MsgBox "��������η�Χ���ߴ��������ߣ�", vbInformation, gstrSysName
                     Cancel = True: Exit Sub
                End If
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsGridRule_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGridRule
        Select Case Col
        Case .ColIndex("���Ʒ�")
            If .ComboIndex < 0 Then .TextMatrix(Row, Col) = ""
        Case Else
        End Select
    End With
End Sub
Private Sub vsGridRule_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytEditType = EM_Ty_�鿴 Then Cancel = True: Exit Sub
    With vsGridRule
         .ComboList = ""
        Select Case Col
        Case .ColIndex("ɾ��")
            .ComboList = "..."
           ' .CellButtonPicture = imgDel
            Exit Sub
        Case .ColIndex("���Ʒ�")
            .ComboList = ">=|>|=|<=|<|���η�Χ"
        Case .ColIndex("����")
            
        Case Else
              .ComboList = " |��ֹ|��ʾ"
        End Select
    End With
End Sub

Private Function CheckRuleDataValid(ByVal intRow As Integer) As Boolean
    '���ܣ���鲻����Ϊ���ƹ������������ݵĺϷ���
    '��Σ�intRow-������Ϊ���ƹ��������
    Dim i As Integer, strTemp As String
    
    With vsGridRule
        For i = 2 To .Cols - 1
            If Trim(.TextMatrix(intRow, i)) <> "" Then
                strTemp = strTemp & .TextMatrix(intRow, i)
            End If
        Next
    End With
    CheckRuleDataValid = strTemp <> ""
End Function
