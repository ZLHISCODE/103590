VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmExecOnePart 
   Caption         =   "���ҽ�����ֲ�λִ��"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "frmExecOnePart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10680
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picOrder 
      Height          =   3735
      Left            =   5160
      ScaleHeight     =   3675
      ScaleWidth      =   4515
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
      Begin VSFlex8Ctl.VSFlexGrid vsfOrder 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4935
         _cx             =   8705
         _cy             =   5741
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
         BackColorBkg    =   -2147483636
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin VB.Frame frmButton 
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   9975
      Begin VB.CommandButton cmdExecPart 
         Caption         =   "�ֲ�λִ��"
         Height          =   400
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "ִ�в�λҽ��"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancelPart 
         Caption         =   "�ֲ�λȡ��"
         Height          =   400
         Left            =   4380
         TabIndex        =   5
         ToolTipText     =   "ȡ����λҽ����ִ��״̬"
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�˳�"
         Default         =   -1  'True
         Height          =   400
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Frame frmInfo 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.Label lblInfo 
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   10095
      End
      Begin VB.Label lblName 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   9735
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   1560
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmExecOnePart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjDockExpense As zlPublicExpense.clsDockExpense

Private mlngDeptID As Long
Private mlngSendNo As Long
Private mlngOrderID As Long

Private Enum Order_Column
    col_ҽ��ID = 0
    col_���ID = 1
    col_��� = 2
    col_ҽ������ = 3
    col_ִ��״̬ = 4
    col_ִ��״̬���� = 5
End Enum

Private Sub cmdCancelPart_Click()
    Dim strSql As String
    
    On Error GoTo err
    
    If vsfOrder.Rows < 1 Or vsfOrder.RowSel < 1 Then
        Call MsgBoxD(Me, "û��ѡ�еĲ�λҽ�������ֲܷ�λȡ��ִ�С�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strSql = "ZL_Ӱ����_CANCEL(" & Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_ҽ��ID)) & "," & mlngSendNo & ",1," & mlngDeptID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    'ˢ�´�������
    RefreshOrder
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdExecPart_Click()
    Dim strSql As String

    On Error GoTo err

    If vsfOrder.Rows < 1 Or vsfOrder.RowSel < 1 Then
        Call MsgBoxD(Me, "û��ѡ�еĲ�λҽ�������ֲܷ�λִ�С�", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    strSql = "Zl_Ӱ����_����ִ��(" & Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_ҽ��ID)) & "," & mlngSendNo & ",'" & UserInfo.��� & _
            "','" & UserInfo.���� & "'," & mlngDeptID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    'ˢ�´�������
    RefreshOrder

    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Public Sub zlShowMe(lngOrderID As Long, strName As String, strAge As String, strSex As String, strState As String, _
    frmParent As Object)
    
    On Error GoTo err
    
    '��ʾ������Ϣ
    lblName = "������" & strName
    lblInfo = "�Ա�" & strSex & "      ���䣺" & strAge & "      ���״̬��" & strState
    
    '��ʼ������
    Call InitForm
    
    mlngOrderID = lngOrderID
    
    'ˢ�´��ڣ�����ҽ��ID����ʾҽ���б�
    Call RefreshOrder
    
    Me.Show 1, frmParent
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitOrder()
    On Error GoTo err
    
    With vsfOrder
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        
        .TextMatrix(0, col_���) = "���"
        .TextMatrix(0, col_ҽ������) = "ҽ������"
        .TextMatrix(0, col_ִ��״̬����) = "ִ��״̬"
        
        .ColWidth(col_���) = 650
        .ColWidth(col_ҽ������) = 3500
        .ColWidth(col_ִ��״̬����) = 600
        
        '����ҽ��ID��
        .ColHidden(col_ҽ��ID) = True
        .ColHidden(col_���ID) = True
        .ColHidden(col_ִ��״̬) = True
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LoadOrder(lngOrderID As Long)
'����ҽ���б�

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSql = "select b.ҽ��id,a.ҽ������,a.�걾��λ ,a.��鷽��,b.ִ��״̬,a.���id,a.ִ�п���id,b.���ͺ�," & _
            " Decode(b.ִ��״̬, 0, 'δִ��', 1, '�����',  2, '�Ѿܾ�', '����ִ��') ִ��״̬���� " & _
            " from ����ҽ����¼ a ,����ҽ������ b " & _
            " where a.id = b.ҽ��id And (a.Id = [1] or a.���ID=[1]) order by ҽ��ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�ֲ�λִ��", lngOrderID)

    If rsTemp.EOF = True Then Exit Sub

    With vsfOrder
        If 1 + rsTemp.RecordCount <> .Rows Then
            .Rows = 1 + rsTemp.RecordCount
        End If

        mlngDeptID = rsTemp!ִ�п���ID
        mlngSendNo = rsTemp!���ͺ�

        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_ҽ��ID) = rsTemp!ҽ��ID
            .TextMatrix(i, col_���ID) = Nvl(rsTemp!���ID, 0)
            .TextMatrix(i, col_���) = IIf(Nvl(rsTemp!���ID, 0) = 0, "��ҽ��", "��" & i - 1 & "��")
            .TextMatrix(i, col_ҽ������) = IIf(Nvl(rsTemp!���ID, 0) = 0, rsTemp!ҽ������, rsTemp!ҽ������ & "��" & Nvl(rsTemp!�걾��λ) & "��" & Nvl(rsTemp!��鷽��) & "��")
            .TextMatrix(i, col_ִ��״̬) = rsTemp!ִ��״̬
            .TextMatrix(i, col_ִ��״̬����) = rsTemp!ִ��״̬����
            rsTemp.MoveNext
        Next i
        
        If .RowSel = 0 Then .RowSel = 1
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = frmInfo.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picOrder.hWnd
    ElseIf Item.ID = 3 Then
        Item.Handle = mobjDockExpense.zlGetForm.hWnd
    ElseIf Item.ID = 4 Then
        Item.Handle = frmButton.hWnd
    End If
End Sub

Private Sub dkpMain_Resize()
    cmdCancelPart.Left = frmButton.Width / 2 - cmdCancelPart.Width / 2
    cmdExecPart.Left = cmdCancelPart.Left - 1000 - cmdExecPart.Width
    cmdExit.Left = cmdCancelPart.Left + cmdCancelPart.Width + 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjDockExpense = Nothing
End Sub

Private Sub picOrder_Resize()
    vsfOrder.Left = 0
    vsfOrder.Top = 0
    vsfOrder.Width = picOrder.Width
    vsfOrder.Height = picOrder.Height
End Sub

Private Sub vsfOrder_SelChange()
    Dim lngOrderID As Long
    Dim blnIsPartOrder As Boolean
    
    On Error GoTo err
    
    '�����ð�ť��Ĭ��ֵ
    cmdExecPart.Enabled = False
    cmdCancelPart.Enabled = False
        
    If vsfOrder.Rows <= 1 Then Exit Sub
    
    lngOrderID = Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_ҽ��ID))
    If lngOrderID = 0 Then Exit Sub
    
    '���ð�ť�����ԣ�ֻ�в�λҽ�����ֲܷ�λִ�к�ȡ��
    blnIsPartOrder = Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_���ID)) <> 0
    If blnIsPartOrder = True Then
        '����ִ��״̬���ж��ĸ���ť����
        ' 0, 'δִ��', 1, '�����',  2, '�Ѿܾ�', 3,'����ִ��'
        If Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_ִ��״̬)) = 3 Then
            cmdCancelPart.Enabled = True
            cmdExecPart.Enabled = False
        ElseIf Val(vsfOrder.TextMatrix(vsfOrder.RowSel, col_ִ��״̬)) = 0 Then
            cmdExecPart.Enabled = True
            cmdCancelPart.Enabled = False
        End If
    End If
    
    'ˢ�·��ô���
    If Not mobjDockExpense Is Nothing Then
        Call mobjDockExpense.zlRefresh(mlngDeptID, lngOrderID & ":" & mlngSendNo & ":" & IIf(blnIsPartOrder, 1, 0))
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitFee()
    On Error GoTo err
    If mobjDockExpense Is Nothing Then
        Set mobjDockExpense = New zlPublicExpense.clsDockExpense
        Call mobjDockExpense.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub InitForm()
'------------------------------------------------
'���ܣ���ʼ������
'��������
'���أ���
'------------------------------------------------
    Dim Pane1 As Pane
    Dim Pane2 As Pane
    Dim Pane3 As Pane
    Dim pane4 As Pane
    
    On Error GoTo err
    
    '��ʼ��ҽ���б�
    Call InitOrder
    
    '��ʼ�������б�
    Call InitFee
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .TabPaintManager.BoldSelected = True
        .Options.DefaultPaneOptions = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        
        Set Pane2 = .CreatePane(2, 200, 200, DockLeftOf)
        Set Pane3 = .CreatePane(3, 200, 200, DockRightOf)
        Set Pane1 = .CreatePane(1, 400, 80, DockTopOf, Pane2 And Pane3)
        Set pane4 = .CreatePane(4, 400, 60, DockBottomOf, Nothing)
    
        Pane1.MaxTrackSize.Height = 80
        Pane1.MinTrackSize.Height = 80
        
        pane4.MaxTrackSize.Height = 60
        pane4.MinTrackSize.Height = 60
        
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshOrder()
'------------------------------------------------
'���ܣ�ˢ�´����е�ҽ���ͷ�������
'��������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    '����ҽ���б���ʾ����
    Call LoadOrder(mlngOrderID)
    
    'ˢ�·��ô���
    Call vsfOrder_SelChange
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub
