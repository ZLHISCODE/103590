VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmWorklist 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraWorkList 
      Height          =   6330
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.Frame frmResultFilter 
         Caption         =   "��ѯ��������"
         Height          =   1215
         Left            =   3120
         TabIndex        =   15
         ToolTipText     =   "���ִ�е���һ��֮��Worklist�в�������ȡ�������Ϣ"
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optResultFilter 
            Caption         =   "������"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "��鱨����ɺ�Worklist��ѯ���ٷ��ظü��"
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton optResultFilter 
            Caption         =   "ͼ��ɼ���Ĭ�ϣ�"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "���յ��豸���ص�ͼ���Worklist��ѯ���ٷ��ظü��"
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��λ���룺�ಿλ����"
         Height          =   1215
         Left            =   5760
         TabIndex        =   9
         Top             =   245
         Width           =   4215
         Begin VB.OptionButton optMultiParts 
            Caption         =   "������"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtPatsSpliter 
            Height          =   300
            Left            =   960
            TabIndex        =   13
            Top             =   817
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "�ָ���"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "���¼"
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   11
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "��"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBodypartCode 
         Caption         =   "��λ����"
         Height          =   350
         Left            =   10080
         TabIndex        =   8
         Top             =   630
         Width           =   1100
      End
      Begin VB.ComboBox cboMatchOther 
         Height          =   300
         ItemData        =   "frmWorkList.frx":0000
         Left            =   1140
         List            =   "frmWorkList.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   247
         Width           =   1665
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   900
         MaxLength       =   4
         TabIndex        =   3
         Top             =   680
         Width           =   435
      End
      Begin VB.CheckBox chkForceResult 
         Caption         =   "ʹ��ǿ�ƽ��"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1380
      End
      Begin VB.CommandButton cmdResetWLResult 
         Caption         =   "�ָ�Ĭ��ֵ"
         Height          =   350
         Left            =   10080
         TabIndex        =   1
         Top             =   247
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   4560
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   11085
         _cx             =   19553
         _cy             =   8043
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
         Rows            =   3
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ƥ��(&A)"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "�ò������[���ݿ���Ŀ]��""���˱�ʶ��""/""����""ƥ����Ч"
         Top             =   307
         Width           =   990
      End
      Begin VB.Label LblSe 
         Caption         =   "�������      �������"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   733
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWorklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColReturn
    ColID = 0
    Col����ID
    Col���
    Col�ϼ�ID
    Col���ı���
    ColӢ�ı���
    Col����ֵ
    Col�Ƿ�Ƕ������
    Col�Ƿ����
    Colֵ����
    Colѡ��
    ColԪ������
    Colǿ�ƽ��ֵ
    ColĬ��ֵ
    ColĬ��ѡ��
    ColĬ��ǿ�ƽ��ֵ
End Enum
Private mlngSrvID As Long
Private Const mstrDBItem As String = "|[CallingAET]|[�״�����]|[�״�ʱ��]|[Ӱ�����]|[ִ�м�]|[ִ�й���]|[ҽ��ID]|[���ͺ�]|[ҽ��ID]_[���ͺ�]|[����]|[��ʶ��]|[Ӣ����]|[�Ա�]|[����]|[��������]|[������]|[����豸]|[��鲿λ]|[����]|[��������]|[���벿λ����]|[���벿λ����]"

Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraWorkList.Caption = "�Ϸ��б�����ѡ������δ���棬���ܽ������ã�"
        fraWorkList.Enabled = False
    Else
        fraWorkList.Caption = ""
        fraWorkList.Enabled = True
    End If
    Call RefreshPara
    Call ReFreshReturnData
End Sub

Private Sub cmdBodypartCode_Click()
    '�򿪲�λ�������ô���
    frmMWLBodypartCode.zlSohwMe Me, mlngSrvID
End Sub

Private Sub cmdResetWLResult_Click()
Dim i As Integer
    With vfgList
        For i = 1 To .Rows - 1
            .TextMatrix(i, Col����ֵ) = .TextMatrix(i, ColĬ��ֵ)
            .TextMatrix(i, Colǿ�ƽ��ֵ) = .TextMatrix(i, ColĬ��ǿ�ƽ��ֵ)
            .TextMatrix(i, Col�Ƿ����) = ""
            .TextMatrix(i, Colѡ��) = .TextMatrix(i, ColĬ��ѡ��)
        Next
    End With
End Sub

Public Sub SavePara()
    Dim i As Long
    Dim iMatch As Integer
    
    On Error GoTo errHandle
    zlCommFun.ShowFlash "���ڱ�������", Me
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'WorkList���˷�ʽ','" & NeedNo(cboMatchOther.Text) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����WorkList���豸����")
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'WorkListʹ��ǿ�ƽ��','" & chkForceResult.value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����WorkListʹ��ǿ�ƽ��")
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'WorkList��������','" & txtSearch.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����WorkList��������")
    
    If optMultiParts(1).value = True Then
        iMatch = 1
    ElseIf optMultiParts(2).value = True Then
        iMatch = 2
    ElseIf optMultiParts(3).value = True Then
        iMatch = 3
    Else
        iMatch = 0
    End If
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'Worklist�ಿλ��ʽ','" & iMatch & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Worklist�ಿλ��ʽ")
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'Worklist�ಿλ�ָ���','" & txtPatsSpliter.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Worklist�ಿλ�ָ���")
    
    If optResultFilter(1).value = True Then
        iMatch = 1
    Else
        iMatch = 0
    End If
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'Worklist��ѯ��������','" & iMatch & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Worklist��ѯ��������")
    
    With vfgList
        For i = 1 To .Rows - 1
            gstrSQL = "Zl_Ӱ��MWL�����_UPDATE(" & .TextMatrix(i, ColID) & ",'" & .TextMatrix(i, Col����ֵ) & "'," & IIf(.TextMatrix(i, Col�Ƿ����) = "", 0, 1) & "," & IIf(.TextMatrix(i, Colѡ��) = "", 0, 1) & ",'" & .TextMatrix(i, Colǿ�ƽ��ֵ) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����MWL�����")
        Next
    End With
    zlCommFun.StopFlash
   Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitvfgList
End Sub
Private Sub RefreshPara()
    Dim rsTemp As New ADODB.Recordset, i As Integer
    
    On Error GoTo err
    gstrSQL = "select ����ID,�������� ,����ֵ from Ӱ��DICOM������� where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlngSrvID)
    chkForceResult.value = False
    txtSearch.Text = 3
    cboMatchOther.ListIndex = 0
    optMultiParts(0).value = True
    txtPatsSpliter.Text = ""
    optResultFilter(0).value = True 'Ĭ��ʹ��ͼ��ɼ���ΪWorklist�Ĳ�ѯ��������
    
    Do Until rsTemp.EOF
        Select Case rsTemp!��������
            Case "WorkList���˷�ʽ"
                Call SeekIndexWithNo(cboMatchOther, Nvl(rsTemp!����ֵ, 0), True)
            Case "WorkListʹ��ǿ�ƽ��"
                chkForceResult.value = Nvl(rsTemp!����ֵ)
            Case "WorkList��������"
                txtSearch.Text = Nvl(rsTemp!����ֵ)
            Case "Worklist�ಿλ��ʽ"   '0-�ޣ�1-�ָ�����2-���¼��3-������
                If Nvl(rsTemp!����ֵ, 0) = 1 Then
                    optMultiParts(1).value = True
                    txtPatsSpliter.Enabled = True
                    txtPatsSpliter.BackColor = &H80000005
                ElseIf Nvl(rsTemp!����ֵ, 0) = 2 Then
                    optMultiParts(2).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                ElseIf Nvl(rsTemp!����ֵ, 0) = 3 Then
                    optMultiParts(3).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                Else
                    optMultiParts(0).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                End If
            Case "Worklist�ಿλ�ָ���"
                txtPatsSpliter.Text = Nvl(rsTemp!����ֵ)
            Case "Worklist��ѯ��������"
                If Nvl(rsTemp!����ֵ, 0) = 1 Then
                    optResultFilter(1).value = True
                Else
                    optResultFilter(0).value = True
                End If
        End Select
        rsTemp.MoveNext
    Loop
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReFreshReturnData()
'ˢ��Worklist�Ľ�����ݼ�

    Dim rsTemp As New ADODB.Recordset
    Dim rsQuery As New ADODB.Recordset  '����һ��rsTemp�����ݼ��������ж��Ƿ����ϼ�ID��
    
    On Error GoTo err
    
    InitvfgList
    gstrSQL = "select ID,����ID,���,Ԫ�غ�,�ϼ�ID,���ı���,Ӣ�ı���,����ֵ," & _
                    "�Ƿ�Ƕ������,�Ƿ����,ֵ����,ѡ��,Ԫ������,ǿ�ƽ��ֵ,Ĭ��ֵ,Ĭ��ѡ��,Ĭ��ǿ�ƽ��ֵ" & _
                    " from Ӱ��MWL����� WHERE ����ID=[1] Order by ���,Ԫ�غ�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlngSrvID)
    
    Set rsQuery = rsTemp.Clone
    
    rsTemp.Filter = "�ϼ�ID = NULL"
    Call AddMWLDataset(rsTemp, rsQuery, "")
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddMWLDataset(rsTemp As ADODB.Recordset, rsQuery As ADODB.Recordset, ByVal strPrefix As String)
    
    On Error GoTo err
    
    If rsTemp.EOF = False Then
        If Not IsNull(rsTemp!�ϼ�ID) Then strPrefix = strPrefix & ">"
    End If
    
    While rsTemp.EOF = False
        '��ӽ����
        With vfgList
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, ColID) = rsTemp!ID
            .TextMatrix(.Rows - 1, Col����ID) = rsTemp!����ID
            .TextMatrix(.Rows - 1, Col���) = strPrefix & rsTemp!��� & "," & rsTemp!Ԫ�غ�
            .TextMatrix(.Rows - 1, Col�ϼ�ID) = Nvl(rsTemp!�ϼ�ID)
            .TextMatrix(.Rows - 1, Col���ı���) = strPrefix & rsTemp!���ı���
            .TextMatrix(.Rows - 1, ColӢ�ı���) = strPrefix & rsTemp!Ӣ�ı���
            .TextMatrix(.Rows - 1, Col����ֵ) = Nvl(rsTemp!����ֵ)
            .TextMatrix(.Rows - 1, Col�Ƿ�Ƕ������) = IIf(Nvl(rsTemp!�Ƿ����, 0) = 1, "��", "")
            .TextMatrix(.Rows - 1, Col�Ƿ����) = IIf(rsTemp!�Ƿ���� = 1, "��", "")
            .TextMatrix(.Rows - 1, Colֵ����) = rsTemp!ֵ����
            .TextMatrix(.Rows - 1, Colѡ��) = IIf(rsTemp!ѡ�� = 1, "��", "")
            .TextMatrix(.Rows - 1, ColԪ������) = rsTemp!Ԫ������
            .TextMatrix(.Rows - 1, Colǿ�ƽ��ֵ) = Nvl(rsTemp!ǿ�ƽ��ֵ)
            .TextMatrix(.Rows - 1, ColĬ��ֵ) = Nvl(rsTemp!Ĭ��ֵ)
            .TextMatrix(.Rows - 1, ColĬ��ѡ��) = IIf(Nvl(rsTemp!Ĭ��ѡ��, 0) = 1, "��", "")
            .TextMatrix(.Rows - 1, ColĬ��ǿ�ƽ��ֵ) = Nvl(rsTemp!Ĭ��ǿ�ƽ��ֵ)
        End With
        
        '�����Ƿ����������ݼ����ϼ�ID=��ǰid,����У��������Щ���ݼ�
        rsQuery.Filter = "�ϼ�ID=" & rsTemp!ID
        If rsQuery.RecordCount > 0 Then
            '�鵽���ݼ�����Ҫ������ЩǶ�׵����ݼ�
            Dim rsClone As New ADODB.Recordset
            Set rsClone = rsTemp.Clone
            rsClone.Filter = "�ϼ�ID=" & rsTemp!ID
            Call AddMWLDataset(rsClone, rsQuery, strPrefix)
        End If
        rsTemp.MoveNext
    Wend
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitvfgList()
    
    On Error GoTo err
    
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 1
        .Cols = 16

        
        .ColWidth(ColID) = 0
        .ColWidth(Col����ID) = 0
        .ColWidth(Col���) = 1400
        .ColWidth(Col�ϼ�ID) = 0
        .ColWidth(Col���ı���) = 2000
        .ColWidth(ColӢ�ı���) = 3100
        .ColWidth(Col����ֵ) = 1700
        .ColWidth(Col�Ƿ�Ƕ������) = 0
        .ColWidth(Col�Ƿ����) = 600
        .ColWidth(Colֵ����) = 0
        .ColWidth(Colѡ��) = 600
        .ColWidth(ColԪ������) = 0
        .ColWidth(Colǿ�ƽ��ֵ) = 1200
        .ColWidth(ColĬ��ֵ) = 0
        .ColWidth(ColĬ��ѡ��) = 0
        .ColWidth(ColĬ��ǿ�ƽ��ֵ) = 0

        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col����ID) = "����ID"
        .TextMatrix(0, Col���) = "���"
        .TextMatrix(0, Col�ϼ�ID) = "�ϼ�ID"
        .TextMatrix(0, Col���ı���) = "���ı���"
        .TextMatrix(0, ColӢ�ı���) = "Ӣ�ı���"
        .TextMatrix(0, Col����ֵ) = "����ֵ"
        .TextMatrix(0, Col�Ƿ�Ƕ������) = "Ƕ��"
        .TextMatrix(0, Col�Ƿ����) = "����"
        .TextMatrix(0, Colֵ����) = "ֵ����"
        .TextMatrix(0, Colѡ��) = "ʹ��"
        .TextMatrix(0, ColԪ������) = "Ԫ������"
        .TextMatrix(0, Colǿ�ƽ��ֵ) = "ǿ�ƽ��"
        .TextMatrix(0, ColĬ��ֵ) = "Ĭ��ֵ"
        .TextMatrix(0, ColĬ��ѡ��) = "Ĭ��ѡ��"
        .TextMatrix(0, ColĬ��ǿ�ƽ��ֵ) = "Ĭ��ǿ�ƽ��ֵ"
        
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col����ID) = flexAlignLeftCenter
        .ColAlignment(Col���) = flexAlignLeftCenter
        .ColAlignment(Col�ϼ�ID) = flexAlignLeftCenter
        .ColAlignment(Col���ı���) = flexAlignLeftCenter
        .ColAlignment(ColӢ�ı���) = flexAlignLeftCenter
        .ColAlignment(Col����ֵ) = flexAlignLeftCenter
        .ColAlignment(Col�Ƿ�Ƕ������) = flexAlignLeftCenter
        .ColAlignment(Col�Ƿ����) = flexAlignLeftCenter
        .ColAlignment(Colֵ����) = flexAlignLeftCenter
        .ColAlignment(Colѡ��) = flexAlignLeftCenter
        .ColAlignment(ColԪ������) = flexAlignLeftCenter
        .ColAlignment(Colǿ�ƽ��ֵ) = flexAlignLeftCenter
        .ColAlignment(ColĬ��ֵ) = flexAlignLeftCenter
        .ColAlignment(ColĬ��ѡ��) = flexAlignLeftCenter
        .ColAlignment(ColĬ��ǿ�ƽ��ֵ) = flexAlignLeftCenter
        
        .Editable = flexEDKbdMouse
        .ComboSearch = flexCmbSearchNone
        .ColComboList(Col����ֵ) = mstrDBItem
    End With
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optMultiParts_Click(Index As Integer)
    If Index = 1 Then
        txtPatsSpliter.Enabled = True
        txtPatsSpliter.BackColor = &H80000005
    Else
        txtPatsSpliter.Enabled = False
        txtPatsSpliter.BackColor = &H8000000B
    End If
End Sub

Private Sub vfgList_Click()
    With vfgList
        If .Col = Col�Ƿ���� Or .Col = Colѡ�� Then
            .Editable = flexEDNone
        ElseIf .Col = Col��� Or .Col = Col���ı��� Or .Col = ColӢ�ı��� Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = Col�Ƿ���� Or .Col = Colѡ�� Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "��"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
End Sub
Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not (Col = Colǿ�ƽ��ֵ Or Col = Col����ֵ) Then
        KeyAscii = 0
    End If
End Sub

Private Sub vfgList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Error GoTo err
    
    '��֤����ֵ������
    If Col = Col����ֵ Then
        '���ַ����У����������'�ţ�[]����Ҫƥ�䣻[]�е����������ݿ��ֶ�
        Dim strTmp As String, strValue As String
        Dim intResult As Integer
        Dim strDBItems() As String
        Dim i As Integer
        Dim blnDBMatch As Boolean
        
        strTmp = vfgList.EditText
        strDBItems = Split(mstrDBItem, "|")
        
        If InStr(strTmp, "'") > 0 Then
            Cancel = True
            intResult = 1
        ElseIf InStr(strTmp, "[") <> 0 Then
            Do Until InStr(strTmp, "[") = 0
                If InStr(strTmp, "]") = 0 Or InStr(strTmp, "]") < InStr(strTmp, "[") Then
                    Cancel = True
                    intResult = 2
                    Exit Do
                End If
                blnDBMatch = False
                strValue = Mid(strTmp, InStr(strTmp, "["), InStr(strTmp, "]") - InStr(strTmp, "[") + 1)
                For i = 1 To UBound(strDBItems)
                    If strDBItems(i) = strValue Then
                        blnDBMatch = True
                        Exit For
                    End If
                Next i
                
                If blnDBMatch = False Then
                    Cancel = True
                    intResult = 3
                    Exit Do
                End If
                strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
            Loop
        End If
        
        If Cancel Then
            If intResult = 1 Then
                MsgBoxD Me, "����������Ϸ�,����ʹ�÷��š�'����Ϊ���ӷ���", vbInformation, gstrSysName
            ElseIf intResult = 2 Then
                MsgBoxD Me, "����������Ϸ�,��[���͡�]������Ŀ��ƥ�䡣", vbInformation, gstrSysName
            Else
                MsgBoxD Me, "����������Ϸ�,��[]���е����ݲ������ݿ���Ŀ��", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
