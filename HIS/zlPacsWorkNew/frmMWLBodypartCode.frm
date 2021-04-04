VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMWLBodypartCode 
   Caption         =   "Worklist��λ��������"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmMWLBodypartCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��"
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6000
      TabIndex        =   3
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdImportParts 
      Caption         =   "����PACS��λ"
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   350
      Left            =   7560
      TabIndex        =   0
      Top             =   5880
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsListBodyParts 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _cx             =   16748
      _cy             =   9975
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   200
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
      AutoResize      =   0   'False
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
   Begin VB.Menu menuPopup 
      Caption         =   "ѡ������"
      Visible         =   0   'False
      Begin VB.Menu menuType 
         Caption         =   "��"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMWLBodypartCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng����ID As Long
Private mfrmParent As Form

Private Enum ColReturn
    col��� = 0
    ColID
    Col����ID
    ColPACS��λ����
    Col�豸��λ����
    Col�豸��λ����
End Enum

Public Sub zlSohwMe(frmParent As Form, lng����ID As Long)
    mlng����ID = lng����ID
    Set mfrmParent = frmParent
    
    Call InitList
    Call FillList
    Me.Show , mfrmParent
End Sub

Private Sub InitList()
'��ʼ����λ�����б�

    With vsListBodyParts
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 1
        .Cols = 6
        
        .ColWidth(col���) = 500
        .ColWidth(ColID) = 0
        .ColWidth(Col����ID) = 0
        .ColWidth(ColPACS��λ����) = 3000
        .ColWidth(Col�豸��λ����) = 3000
        .ColWidth(Col�豸��λ����) = 3000
        
        .TextMatrix(0, col���) = "���"
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col����ID) = "����ID"
        .TextMatrix(0, ColPACS��λ����) = "PACS��λ����"
        .TextMatrix(0, Col�豸��λ����) = "�豸��λ����"
        .TextMatrix(0, Col�豸��λ����) = "�豸��λ����"
        
        .ColAlignment(col���) = flexAlignLeftCenter
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col����ID) = flexAlignLeftCenter
        .ColAlignment(ColPACS��λ����) = flexAlignLeftCenter
        .ColAlignment(Col�豸��λ����) = flexAlignLeftCenter
        .ColAlignment(Col�豸��λ����) = flexAlignLeftCenter
        
        .Editable = flexEDKbdMouse
    
    End With
End Sub

Private Sub FillList()
'��䲿λ�����
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "Select id,����ID,PACS��λ����,�豸��λ����,�豸��λ���� From Ӱ��MWL��λ���� Where ����ID =[1] order by id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��λ����", mlng����ID)
    With vsListBodyParts
        .Rows = rsTemp.RecordCount + 1
        While rsTemp.EOF = False
            .TextMatrix(rsTemp.AbsolutePosition, col���) = rsTemp.AbsolutePosition
            .TextMatrix(rsTemp.AbsolutePosition, ColID) = rsTemp!ID
            .TextMatrix(rsTemp.AbsolutePosition, Col����ID) = Nvl(rsTemp!����ID)
            .TextMatrix(rsTemp.AbsolutePosition, ColPACS��λ����) = Nvl(rsTemp!PACS��λ����)
            .TextMatrix(rsTemp.AbsolutePosition, Col�豸��λ����) = Nvl(rsTemp!�豸��λ����)
            .TextMatrix(rsTemp.AbsolutePosition, Col�豸��λ����) = Nvl(rsTemp!�豸��λ����)
            rsTemp.MoveNext
        Wend
    End With
    cmdSave.Enabled = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    'ɾ��ָ����
    Dim strSQL As String
    Dim lngResult As Long
    
    On Error GoTo err
    
    cmdDelete.Enabled = False
    
    '���ж��Ƿ���Ҫ��ʾ����
    If cmdSave.Enabled = True Then
        lngResult = MsgBoxD(mfrmParent, "�����ݱ��޸�û�б��棬�Ƿ���Ҫ���棿", vbYesNoCancel, "��ʾ��Ϣ")
        If lngResult = vbYes Then
            Call SaveDate
        ElseIf lngResult = vbCancel Then
            cmdDelete.Enabled = True
            Exit Sub
        End If
    End If
    
    '���û��ѡ���κ��У��򲻶���
    '���ѡ����û��ID��˵��û�б��浽���ݿ��У�ֱ�����б���ɾ���������ID ���������ݿ���ɾ�������ڱ���ɾ��
    
    If Val(vsListBodyParts.TextMatrix(vsListBodyParts.Row, ColID)) <> 0 Then
        strSQL = "ZL_Ӱ��MWL��λ����_ɾ��(" & vsListBodyParts.TextMatrix(vsListBodyParts.Row, ColID) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "ɾ����λ����")
    End If
    
    '����װ������
    Call FillList
    
    cmdDelete.Enabled = True
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    If cmdSave.Enabled = True Then
        If MsgBoxD(mfrmParent, "��λ�����иĶ���ȷ��Ҫ�����Ķ��˳���", vbYesNo, "��ʾ��Ϣ") = vbNo Then
            Exit Sub
        End If
    End If
    '�رմ���
    Unload Me
End Sub

Private Sub cmdImportParts_Click()
'����PACS�ļ�鲿λ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim iCount As Integer
    Dim i As Integer
    
    '��ж��ԭ���Ĳ˵�
    On Error Resume Next
    iCount = menuType.Count
    If iCount > 1 Then
        menuType(0).Visible = True
        For i = 1 To iCount - 1
            Unload menuType(i)
        Next i
    End If
    
    On Error GoTo err
    '�Ȳ�ѯ��λ�����ͣ�ѡ�����ͺ�������͵�ȫ����λ
    strSQL = "Select distinct ���� from ���Ƽ�鲿λ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��鲿λ����")
    
    While rsTemp.EOF = False
        iCount = menuType.Count
        Load menuType(iCount)
        menuType(iCount).Caption = Nvl(rsTemp!����)
        rsTemp.MoveNext
    Wend
    
    If menuType.Count = 1 Then
        MsgBoxD mfrmParent, "���Ƽ�鲿λ����û�в�λ��Ϣ�����ȵ������Ʋ�λ����ģ�����ò�λ��"
    Else
        menuType(0).Visible = False
        PopupMenu menuPopup
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    '������벿λ
    Call SaveDate
    cmdSave.Enabled = False
    '����װ��
    Call FillList
End Sub

Private Sub SaveDate()
    '������벿λ
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo err
    
    For i = 1 To vsListBodyParts.Rows - 1
        '�����ݲű���
        If vsListBodyParts.TextMatrix(i, ColPACS��λ����) <> "" Then
            
            strSQL = "Zl_Ӱ��MWL��λ����_����(" & _
                    IIf(vsListBodyParts.TextMatrix(i, ColID) = "", "NULL", vsListBodyParts.TextMatrix(i, ColID)) & _
                    "," & mlng����ID & ",'" & vsListBodyParts.TextMatrix(i, ColPACS��λ����) & "','" & _
                    vsListBodyParts.TextMatrix(i, Col�豸��λ����) & "','" & _
                    vsListBodyParts.TextMatrix(i, Col�豸��λ����) & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "������벿λ")
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    '�������ڿؼ���λ��
    vsListBodyParts.Left = 0
    vsListBodyParts.Top = 0
    vsListBodyParts.Width = Me.ScaleWidth
    vsListBodyParts.Height = Me.ScaleHeight - cmdExit.Height - 200
    
    cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 100
    cmdDelete.Top = cmdExit.Top
    cmdSave.Top = cmdExit.Top
    cmdImportParts.Top = cmdExit.Top
    
    cmdExit.Left = Me.ScaleWidth - cmdExit.Width - 100
    cmdSave.Left = cmdExit.Left - cmdSave.Width - 200
    cmdDelete.Left = cmdSave.Left - cmdDelete.Width - 200
End Sub

Private Sub menuType_Click(Index As Integer)
'    ��λ�����Ľ���˵��
'    ���ӣ���0���ⷽ��,0���ӿ�ѡ����,1���ӿ�ѡ����2;1���ⷽ��2   0����λ;0����λ;1­��λ;1���÷�����
'
'    1����TAB�� ���ֻ�����������͹��÷���
'    2����;������ÿһ����������
'    3����,�����ֻ��������еĸ��ӷ���
'    4������ǰ��1λ���ִ����Ƿ���Ӱ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str���� As String
    Dim str���� As String
    Dim arrItem() As String     '�������������
    Dim arrChild() As String    '���渽�ӷ�����
    Dim lngItem As Long         '��������������
    Dim lngChild As Long        '���ӷ���������
    Dim strTemp As String
    
    
    On Error GoTo err
    
    If menuType(Index).Caption <> "" Then
        '�������Ͳ����λ���������ݼ�
        strSQL = "Select ����,���� From ���Ƽ�鲿λ Where ����=[1] order by  ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��λ����", CStr(menuType(Index).Caption))
        
        '��������д��λ����
        While rsTemp.EOF = False
            str���� = Nvl(rsTemp!����)
            str���� = Nvl(rsTemp!����)
            If UBound(Split(str����, vbTab)) >= 0 Then  '=0��ʾû�л��ⷽ����>0��ʾ�л��ⷽ��
                arrItem() = Split(Split(str����, vbTab)(0), ";")    '�õ�ÿһ����������
                For lngItem = 0 To UBound(arrItem)
                    strTemp = Mid(arrItem(lngItem), 2)
                    If InStr(1, strTemp, ",") > 0 Then  '����С������ţ���ʾ�������ӷ�������Ҫ��һ������
                        arrChild = Split(strTemp, ",")
                        strTemp = ""
                        Call AddOneBodypart(str���� & arrChild(0))
                        For lngChild = 1 To UBound(arrChild)
                            Call AddOneBodypart(str���� & Mid(arrChild(lngChild), 2))
                        Next lngChild
                    Else
                        Call AddOneBodypart(str���� & strTemp)
                    End If
                Next lngItem
            End If
            If UBound(Split(str����, vbTab)) > 0 Then   '>0��ʾ�л��ⷽ��,����͸��Ź��÷����������÷���
                arrItem() = Split(Split(str����, vbTab)(1), ";")
                For lngItem = 0 To UBound(arrItem)
                    Call AddOneBodypart(str���� & Mid(arrItem(lngItem), 2))
                Next lngItem
            End If
            rsTemp.MoveNext
        Wend
        cmdSave.Enabled = True
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddOneBodypart(strBodypartName As String)
'���б������һ����λ����
'������ strBodypartName ---��λ��������
    Dim i As Integer
    
    On Error GoTo err
    
    '���ȼ���Ƿ���ͬ���ģ�����У��Ͳ����
    For i = 1 To vsListBodyParts.Rows - 1
        If vsListBodyParts.TextMatrix(i, ColPACS��λ����) = strBodypartName Then
            Exit Sub
        End If
    Next i
    
    '��Ӳ�λ����
    With vsListBodyParts
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, col���) = .Rows - 1
        .TextMatrix(.Rows - 1, ColPACS��λ����) = strBodypartName
    End With
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsListBodyParts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    cmdSave.Enabled = True
End Sub

Private Sub vsListBodyParts_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error Resume Next
    
    '�س�ת�Ƶ���һ���༭��
    If KeyAscii = vbKeyReturn Then
        If Col = ColPACS��λ���� Then
            vsListBodyParts.Selecte Row, Col�豸��λ����
        ElseIf Col = Col�豸��λ���� Then
            vsListBodyParts.Selecte Row, Col�豸��λ����
        ElseIf Col = Col�豸��λ���� Then   '�س�������һ������,��ת����һ���༭��
            If vsListBodyParts.Row = vsListBodyParts.Rows - 1 Then
                vsListBodyParts.TextMatrix(vsListBodyParts.Row, col���) = vsListBodyParts.Rows - 1
                vsListBodyParts.Rows = vsListBodyParts.Rows + 1
                vsListBodyParts.Select Row + 1, ColPACS��λ����
            End If
        End If
'        vsListBodyParts.EditCell
    End If
End Sub
