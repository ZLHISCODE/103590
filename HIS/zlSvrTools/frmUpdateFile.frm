VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmUpdateFile 
   BackColor       =   &H80000005&
   Caption         =   "վ�����п���"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUpdateFile.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2385
      TabIndex        =   20
      Top             =   5745
      Width           =   885
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   885
      TabIndex        =   19
      Top             =   5775
      Width           =   1500
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   1
      Left            =   15
      ScaleHeight     =   1695
      ScaleWidth      =   8475
      TabIndex        =   15
      Top             =   1590
      Width           =   8475
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   1515
         Left            =   15
         TabIndex        =   16
         Top             =   15
         Width           =   8325
         _cx             =   14684
         _cy             =   2672
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761024
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   12648384
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483630
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   2
         GridLines       =   1
         GridLinesFixed  =   12
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUpdateFile.frx":04F9
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   5040
      ScaleHeight     =   555
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   1035
      Width           =   3375
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "ϵͳ�ļ�"
         Height          =   240
         Index           =   5
         Left            =   1065
         TabIndex        =   14
         Top             =   255
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         Height          =   240
         Index           =   4
         Left            =   2130
         TabIndex        =   13
         Top             =   255
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "�����ļ�"
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   12
         Top             =   255
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "�����ļ�"
         Height          =   240
         Index           =   2
         Left            =   2130
         TabIndex        =   11
         Top             =   0
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "Ӧ�ò���"
         Height          =   240
         Index           =   1
         Left            =   1065
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chk���� 
         BackColor       =   &H80000005&
         Caption         =   "��������"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Value           =   1  'Checked
         Width           =   1050
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   45
      ScaleHeight     =   330
      ScaleWidth      =   3855
      TabIndex        =   6
      Top             =   1170
      Width           =   3855
      Begin VB.ComboBox cboSystem 
         Height          =   300
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   15
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����ϵͳ"
         Height          =   180
         Left            =   90
         TabIndex        =   17
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "�޸�(&E)"
      Height          =   360
      Left            =   7455
      TabIndex        =   5
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   360
      Left            =   6480
      TabIndex        =   4
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   360
      Left            =   5505
      TabIndex        =   3
      Top             =   5730
      Width           =   945
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":065A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   4395
      TabIndex        =   0
      Top             =   5745
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���Ʋ���"
      Height          =   180
      Left            =   90
      TabIndex        =   18
      Top             =   5835
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����ļ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   2
      Top             =   105
      Width           =   1440
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "����鿴�����ļ�,�ɶԵ����������������ӡ�ɾ�����޸ġ�"
      Height          =   345
      Left            =   1215
      TabIndex        =   1
      Top             =   750
      Width           =   7365
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   225
      Picture         =   "frmUpdateFile.frx":1124
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmUpdateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const StopColor = vbRed '����ʱ����ɫ
Const StartColor = &H80000008 '����ʱ����ɫ
Dim mintColumn As Integer '

Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private m_strCurTypeName        As String               '��ǰѡ�еķ�ʽ
Private m_strCurFileName        As String               '��ǰѡ�е�����
Private m_strCurVision          As String               '��ǰѡ�еİ汾
Private m_strCurEditDate        As String               '��ǰѡ�е��޸�����
Private m_strCurSysNum          As String               '��ǰѡ�е�ϵͳ
Private m_strCurSetupPath       As String               '��ǰѡ�еİ�װ·��
Private m_strCurSysOption       As String               '��ǰѡ�е�ϵͳ����
Private m_strCurFileExplanation As String               '��ǰѡ�е��ļ�˵��
Private m_strCurSellFile        As String               '��ǰѡ�е������ļ�
Private m_blnCurReg             As Boolean              '��ǰѡ�е��ļ��Ƿ�ע��
Private m_blnCurUpData          As Boolean              '��ǰѡ�е��ļ��Ƿ�ǿ�Ƹ���

Private m_lngCurRow             As Long
Dim mrsTemp      As New ADODB.Recordset

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = "�����ļ��嵥"
    
  
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(date, "yyyy��MM��dd��")
    Set objPrint.Body = Me.fgMain
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cboSystem_Click()
    Call refData
End Sub

Private Sub cmdAdd_Click()
    '����
    Call StandardAdd
End Sub

Private Sub cmdDel_Click()
    'ɾ��
    Call StandardDel
End Sub

Private Sub cmdEdit_Click()
    '�޸�
     Call StandardEdit
End Sub


Private Sub cmdFind_Click()
    txtFind_KeyPress 13
End Sub

Private Sub cmdRefresh_Click()
     Call refData 'ˢ��
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
    If KeyCode = vbKeyDelete Then
        If cmdDel.Enabled Then
            cmdDel_Click
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngWdt As Single
    
    err = 0
    On Error Resume Next
    lblNote.Width = ScaleWidth - lblNote.Left
    With cmdRefresh
        .Top = ScaleHeight - .Height - 50
    End With
    
    picPane(0).Move ScaleWidth - picPane(0).Width - 30, picPane(0).Top
    
    picPane(1).Move picPane(1).Left, picPane(1).Top, ScaleWidth - 30, cmdRefresh.Top - picPane(1).Top - 50
    

    With cmdAdd
        .Top = cmdRefresh.Top
        .Left = ScaleWidth - cmdAdd.Width * 3 - 30
    End With
    
    
    With cmdEdit
        .Top = cmdRefresh.Top
        .Left = cmdAdd.Left + cmdAdd.Width
    End With
    
    With cmdDel
        .Top = cmdRefresh.Top
        .Left = cmdEdit.Left + cmdEdit.Width
    End With
    
    Label2.Top = cmdRefresh.Top
    txtFind.Top = cmdRefresh.Top
    cmdFind.Top = cmdRefresh.Top
  
    
End Sub


'==============================================================================
'=���ܣ� ���ڳ�ʼ��
'==============================================================================
Private Sub Form_Load()
  On Error GoTo errH
    
    KeyPreview = True
    m_lngCurRow = -1
    '���Combo
    Call InitComBo

'    Call SetMenu
    Exit Sub
    
errH:
    MsgBox err.Description, vbInformation, "��ʾ"
End Sub


'==============================================================================
'=���ܣ� ����fgMain������ˢ��״̬��Ϣ
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    fgMain_SelChange
    Exit Sub
errH:
   
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� fgMain
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error GoTo errH
'    Select Case Button
'        Case 2          '�����˵�����
'            Call SendLMouseButton(fgMain.hwnd, X, Y)
'            mcbrPopupBarItem.ShowPopup
'    End Select
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub fgMain_RowColChange()
    On Error GoTo errH
    Call fgMain_SelChange
    Exit Sub
errH:

End Sub

'==============================================================================
'=���ܣ� ����ѡ�����б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub fgMain_SelChange()
    Dim lngID       As Long
    On Error GoTo errH
    
'    fgMain.WallPaper = imgBG_fg(1).Picture
    m_lngCurRow = fgMain.Row
    If m_lngCurRow = 0 Then Exit Sub
    m_strCurTypeName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 1)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 1))   '��ȡID
    m_strCurFileName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 2))     '�ļ���
    m_strCurVision = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 3)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 3))
    m_strCurEditDate = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 4))
    m_strCurSysNum = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 5))
    m_strCurSellFile = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 6))
    m_strCurSetupPath = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 7)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 7))
    m_strCurSysOption = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 10)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 10))
    m_strCurFileExplanation = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 11)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 11)) '�ļ�˵��
    m_blnCurReg = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 12)) = 0, False, fgMain.Cell(flexcpText, m_lngCurRow, 12)) '�Զ�ע��
    m_blnCurUpData = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 13)) = 0, False, fgMain.Cell(flexcpText, m_lngCurRow, 13)) 'ǿ�Ƹ���
    
    If m_strCurTypeName = "��������" Then
        cmdEdit.Enabled = True
        cmdDel.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
    End If
    
    Call SetMenu
    Exit Sub
errH:

End Sub

Private Sub fgMain_DblClick()
    If m_strCurTypeName = "��������" Then
        Call StandardEdit
    End If
End Sub

'==============================================================================
'=���ܣ� ���ϵͳ ComBo
'==============================================================================
Private Sub InitComBo()
    On Error GoTo errH
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngDefaultNum As Long
    Dim str���       As String
    With cboSystem
    .Clear
    strSQL = "select ���,����,����� from zlSystems order by ���"
    Call OpenRecordset(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        rs.MoveFirst
        .AddItem "[0]����ϵͳ"
        .ItemData(.NewIndex) = 0
        Do While Not rs.EOF
            str��� = rs("���").value \ 100
            .AddItem "[" & str��� & "]" & rs("����").value
            .ItemData(.NewIndex) = str���
            If Nvl(rs("�����").value, 0) = 0 Then
                lngDefaultNum = .ListCount - 1
            End If
            rs.MoveNext
        Loop
    End If
    .ListIndex = 0 'lngDefaultNum
    End With
    Exit Sub
errH:

End Sub

'==============================================================================
'=���ܣ� װ���Ӧ���������ֱ�׼
'==============================================================================
Public Sub DataLoad(Optional ByVal strFilter As String)

    Dim i            As Long
    Dim strSQL       As String
    Dim strSystemNum As String
    Dim strTypeID()  As String
    Dim strTemp      As String
    On Error GoTo errH
    
    With fgMain
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cols = 14
'        Exit Sub
        .Cell(flexcpText, 0, 0) = "���"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "�ļ�����"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "�ļ���"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "�汾��"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "�޸�����"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "����ϵͳ"
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        .ColWidth(5) = 1800
        .Cell(flexcpText, 0, 6) = "ҵ�񲿼�"
        .Cell(flexcpAlignment, 0, 6) = flexAlignCenterCenter
        .ColWidth(6) = 4800
        
        .Cell(flexcpText, 0, 7) = "��װ·��"
        .Cell(flexcpAlignment, 0, 7) = flexAlignCenterCenter
        .ColWidth(7) = 0
        
        .Cell(flexcpText, 0, 8) = "����ID"
        .Cell(flexcpAlignment, 0, 8) = flexAlignCenterCenter
        .ColWidth(8) = 0
        
        .Cell(flexcpText, 0, 9) = "��װ·��"
        .Cell(flexcpAlignment, 0, 9) = flexAlignCenterCenter
        .ColWidth(9) = 2000
         
        .Cell(flexcpText, 0, 10) = "ϵͳ����"
        .Cell(flexcpAlignment, 0, 10) = flexAlignCenterCenter
        .ColWidth(10) = 0
        .Cell(flexcpText, 0, 11) = "�ļ�˵��"
        .Cell(flexcpAlignment, 0, 11) = flexAlignCenterCenter
        .ColWidth(11) = 1000
        
        .Cell(flexcpText, 0, 12) = "�Զ�ע��"
        .Cell(flexcpAlignment, 0, 12) = flexAlignCenterCenter
        .ColWidth(12) = 0
        
        .Cell(flexcpText, 0, 13) = "ǿ�Ƹ���"
        .Cell(flexcpAlignment, 0, 13) = flexAlignCenterCenter
        .ColWidth(13) = 0
        
        If CheckTable = False Then
            Exit Sub
        End If
        
        If Len(strFilter) <> 0 Then
            If strFilter = "Clear" Then
                Exit Sub
            Else
                strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
                If strSystemNum = "" Then strSystemNum = "1"
                
                If strSystemNum = "0" Then
                     strSQL = "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                             "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.�ļ����� In (" & strFilter & ") order by lpad(a.���,5,'0')"
                             
                              Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                              GoTo zt
                End If
                
                If InStrRev(strFilter, "0") > 0 Then
                   strTypeID = Split(strFilter, ",")
                   For i = 0 To UBound(strTypeID)
                        If strTemp = "" Then
                            strTemp = strTypeID(i)
                        Else
                            strTemp = strTemp & "," & strTypeID(i)
                        End If
                   Next
                    strSQL = "Select B.���,B.����ID,B.�ļ�����,B.�ļ���,B.�汾��,B.�޸�����,B.����ϵͳ,B.ҵ�񲿼�,B.��װ·��,B.�ļ�˵��,B.�Զ�ע�� From ( " & vbNewLine & _
                                "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                                "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.�ļ����� In (" & strTemp & ") And (Instr(a.����ϵͳ, ','|| " & strSystemNum & " ||  ',' ) > 0 or a.����ϵͳ is null )" & vbNewLine & _
                                "Union" & vbNewLine & _
                                "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                                "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.�ļ����� =0" & vbNewLine & _
                                ") B Order by lpad(B.���,5,'0')"
                        
                    Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                Else
                    strSQL = "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                             "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.�ļ����� In (" & strFilter & ") And (Instr(a.����ϵͳ, ',' || " & strSystemNum & " || ',' ) > 0 or a.����ϵͳ is null ) order by lpad(a.���,5,'0')"

                    
             
                        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                End If
            End If
        Else
            strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
            If strSystemNum = "" Then strSystemNum = "100"
    
            strSQL = "Select B.���,B.����ID,B.�ļ�����,B.�ļ���,B.�汾��,B.�޸�����,B.����ϵͳ,B.ҵ�񲿼�,B.��װ·��,B.�ļ�˵��,B.�Զ�ע�� From ( " & vbNewLine & _
                        "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                         "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.�ļ����� In (1, 2, 3,4) And (Instr(a.����ϵͳ,  ',' ||  " & strSystemNum & " || ',') > 0 or a.����ϵͳ is null )" & vbNewLine & _
                         "Union" & vbNewLine & _
                         "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�汾��, a.�޸�����," & vbNewLine & _
                         "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.�ļ����� =0" & vbNewLine & _
                         ") B Order by lpad(B.���,5,'0')"
        
            Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
        End If
zt:
'    .AllowSelection = False '����
'    .Editable = flexEDKbdMouse
'    .AllowUserResizing = flexResizeBoth
'    .AllowUserFreezing = flexFreezeBoth
'    .BackColorFrozen = 14737632
'    .GridLines = flexGridFlatVert
    .ExtendLastCol = True
'    .ScrollTips = True
    
        .FocusRect = flexFocusSolid
        '��������
        .Rows = mrsTemp.RecordCount + 1
    
        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, 0) = Nvl(mrsTemp.Fields("���"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, 0) = flexAlignLeftCenter
            
            
            .Cell(flexcpText, i, 1) = Nvl(mrsTemp.Fields("�ļ�����"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
'            If NVL(mrsTemp.Fields("�ļ�����")) = "Ӧ�ò���" Then
'                .Cell(flexcpBackColor, i, 1) = &H80C0FF   '&H8080FF
'            End If
            .Cell(flexcpText, i, 2) = Nvl(mrsTemp.Fields("�ļ���"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftCenter
            
            strTemp = Nvl(mrsTemp.Fields("�汾��"))
            strTemp = GetFileVision(strTemp)
            
            .Cell(flexcpText, i, 3) = strTemp
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            
            If Nvl(mrsTemp.Fields("�޸�����")) <> "" Then
                strTemp = Format(Nvl(mrsTemp.Fields("�޸�����")), "yyyy-mm-dd hh:mm:ss")
            Else
                strTemp = ""
            End If
            
            .Cell(flexcpText, i, 4) = strTemp
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            
            strTemp = Nvl(mrsTemp.Fields("����ϵͳ"))
            
            If Len(strTemp) <> 0 Then
                strTemp = Right(strTemp, Len(strTemp) - 1)
                strTemp = Left(strTemp, Len(strTemp) - 1)
            End If
            
            If strTemp = "" Then
                strTemp = "����ϵͳ"
            End If
            
            .Cell(flexcpText, i, 5) = strTemp
            .Cell(flexcpAlignment, i, 5) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 6) = Nvl(mrsTemp.Fields("ҵ�񲿼�"))
            .Cell(flexcpAlignment, i, 6) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 7) = Nvl(mrsTemp.Fields("��װ·��"))
            .Cell(flexcpAlignment, i, 7) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 8) = Nvl(mrsTemp.Fields("����ID"))
            .Cell(flexcpAlignment, i, 8) = flexAlignLeftTop
            
            .Cell(flexcpText, i, 9) = Nvl(mrsTemp.Fields("��װ·��"))
            .Cell(flexcpAlignment, i, 9) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 10) = Nvl(mrsTemp.Fields("����ϵͳ")) 'NVL(mrsTemp.Fields("ϵͳ����"))
            .Cell(flexcpAlignment, i, 10) = flexAlignCenterCenter
            
            .Cell(flexcpText, i, 11) = Nvl(mrsTemp.Fields("�ļ�˵��"), "")
            .Cell(flexcpAlignment, i, 11) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 12) = Nvl(mrsTemp.Fields("�Զ�ע��"), "")
            .Cell(flexcpAlignment, i, 12) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 13) = ""
            .Cell(flexcpAlignment, i, 13) = flexAlignLeftCenter
            
            mrsTemp.MoveNext
            i = i + 1
        Loop

        '�Զ�����
        .WordWrap = True
        '�ϲ���Ԫ��
        .MergeCells = 0
        .MergeCol(.ColIndex("�ļ�����")) = True
        .MergeCol(.ColIndex("�ļ���")) = True
        '���ص�Ԫ��
        .ColWidth(.ColIndex("����ID")) = 0
        
        '�и�����
        .RowHeightMin = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("ҵ�񲿼�")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
         Call SetMenu
    End With
    Exit Sub
errH:

   
End Sub


'==============================================================================
'=���ܣ� ��ʾ��¼����Ϣ
'==============================================================================
Private Sub SetMenu()
 
    frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & fgMain.Rows - 1 & "�����ݡ�"

End Sub

'==============================================================================
'=���ܣ� �����Ƿ����±���߱��Ƿ����
'==============================================================================
Private Function CheckTable() As Boolean
    On Error GoTo errH
    Dim strSQL As String
    Dim i As Integer
    Dim blnUse As Boolean
    strSQL = "select * from zlFilesUpgrade where rownum =1"
    
    
    Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
    If mrsTemp.RecordCount >= 0 Then
        For i = 1 To mrsTemp.Fields.Count
            If mrsTemp.Fields.Item(i - 1).name = "����ϵͳ" Then
                blnUse = True
                Exit For
            End If
        Next
        
        If blnUse Then
            CheckTable = True
        Else
            MsgBox "��zlFilesUpgrade����,û���ҵ���Ӧ���ֶ�!" & vbCrLf & "�����ṹ�Ƿ�Ϊ����!", vbInformation
            CheckTable = False
        End If
    End If
    Exit Function
errH:

End Function




'��ȡ�汾��ֱ����ʾֵ
Private Function GetFileVision(ByVal strVision As String) As String
    Dim lng�汾�� As Variant
    Dim str�汾�� As String
    If Len(strVision) > 0 Then
        lng�汾�� = strVision
        str�汾�� = Int(lng�汾�� / 10 ^ 8)
        If Len(lng�汾��) > 9 Then
            lng�汾�� = Right(lng�汾��, 9) Mod (10 ^ 8)
        Else
            lng�汾�� = lng�汾�� Mod (10 ^ 8)
        End If
        
        str�汾�� = str�汾�� & "." & Int(lng�汾�� / 10 ^ 4)
        lng�汾�� = lng�汾�� Mod 10 ^ 4
        str�汾�� = str�汾�� & "." & lng�汾��
        GetFileVision = str�汾��
    End If
End Function

Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant
    
    err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Sub picPane_Resize(Index As Integer)
    Select Case Index
    Case 0
    Case 1
         fgMain.Move 15, 15, picPane(1).Width - 15 * 2, picPane(1).Height - 15 * 2
    End Select
End Sub


'==============================================================================
'=���ܣ� ˢ������
'==============================================================================
Private Sub refData()
    Dim strTemp As String
    On Error GoTo errH
    If chk����(0).value Then
        strTemp = "0,"
    End If
    
    If chk����(1).value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk����(2).value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk����(3).value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk����(4).value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk����(5).value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        Call DataLoad(strTemp)
    Else
        Call DataLoad("Clear")
    End If
    Exit Sub
errH:
End Sub

Private Sub chk����_Click(Index As Integer)
    Dim strTemp As String
    On Error GoTo errH
    If chk����(0).value Then
        strTemp = "0,"
    End If
    
    If chk����(1).value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk����(2).value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk����(3).value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk����(4).value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk����(5).value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        Call DataLoad(strTemp)
    Else
        Call DataLoad("Clear")
    End If
    Exit Sub
errH:

End Sub


'==============================================================================
'=�޸��ļ�
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 100
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    f.ShowForm "�޸�", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0"
    If f.Moded Then
        Call refData
        Dim lngRow As Long
        lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
        If lngRow <> -1 Then
              fgMain.Select lngRow, 2
              fgMain.ShowCell lngRow, 2
        End If
    End If
    Exit Sub
errH:
 
End Sub


'==============================================================================
'=�����ļ�
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 1
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    f.ShowForm "����", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0"
    If f.Moded Then
    
        Call refData
        fgMain.Row = fgMain.Rows
        fgMain.Select fgMain.Rows, 1, fgMain.Rows, 1
        fgMain.ro
    End If
    Exit Sub
errH:
  
End Sub

'==============================================================================
'=ɾ���ļ�
'==============================================================================
Private Sub StandardDel()
    Dim i         As Long
    Dim strName   As String
    Dim lngCurRow As Long
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strSys    As String
    Dim strSysNum As String
    Dim lngRow    As Long
    On Error GoTo errH
    
    If fgMain.SelectedRows = 0 Then Exit Sub
    
    If m_strCurTypeName <> "��������" Then
        Exit Sub
    End If
    
    If fgMain.SelectedRows = 1 Then
        If MsgBox("��ȷ��Ҫɾ��[" & Right(cboSystem.Text, Len(cboSystem.Text) - InStrRev(cboSystem.Text, "]", -1)) & "]" & vbCrLf & "�Ĳ���" & m_strCurFileName & "��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("��ȷ��Ҫɾ��ѡ���" & fgMain.SelectedRows & "��������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
'    gcnOracle.BeginTrans
    
    
    lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
    
    For i = 0 To fgMain.SelectedRows
        If fgMain.SelectedRow(i) Then
            lngCurRow = fgMain.SelectedRow(i)
            If lngCurRow <> -1 Then
                strName = IIf(Len(fgMain.Cell(flexcpText, lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, lngCurRow, 2))
                strName = UCase(strName)
            
                gstrSQL = "delete zlFilesUpgrade where upper(�ļ���)= upper('" & strName & "')"
                gcnOracle.Execute gstrSQL
'                End If
            End If

        End If
    Next
    
'    gcnOracle.CommitTrans
    
    ''Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Call refData
    Call SetMenu
    
    
    If lngRow <> -1 Then
        If lngRow >= 2 And fgMain.Rows > 2 Then
          fgMain.Select lngRow - 1, 2
          fgMain.ShowCell lngRow - 1, 2
        End If
    End If
    Exit Sub
errH:
'    gcnOracle.RollbackTrans

End Sub



'==============================================================================
'=���ٶ�λ
'==============================================================================
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long

    On Error GoTo errH
    
    lngRow = 0
    If txtFind.Locked Then Exit Sub
    If mstrFindKey = "����" Then mstrFindKey = "�ļ�����"
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
            If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To fgMain.Row
                If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then
            fgMain.Row = lngRow
            fgMain.ShowCell lngRow, 2
        End If
        
        
        'Call LocationObj(txtFind)
    End If
    If mstrFindKey = "�ļ�����" Then mstrFindKey = "����"

    Exit Sub
errH:
    mstrFindKey = "����"
    
End Sub


