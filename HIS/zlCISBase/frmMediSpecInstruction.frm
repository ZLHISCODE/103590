VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMediSpecInstruction 
   Caption         =   "��...��ʹ��˵���༭"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "frmMediSpecInstruction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   9390
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUp 
      Caption         =   "����"
      Height          =   350
      Left            =   210
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4815
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8220
      TabIndex        =   6
      Top             =   4815
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6900
      TabIndex        =   5
      Top             =   4815
      Width           =   1100
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   5355
      TabIndex        =   4
      Top             =   4815
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����"
      Height          =   350
      Left            =   1515
      TabIndex        =   3
      Top             =   4815
      Width           =   1100
   End
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   3315
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4365
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   420
      Width           =   75
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
      Height          =   4230
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   2925
      _cx             =   5159
      _cy             =   7461
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediSpecInstruction.frx":000C
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin RichTextLib.RichTextBox rtbDetails 
      Height          =   4230
      Left            =   3600
      TabIndex        =   0
      Top             =   435
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7461
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmMediSpecInstruction.frx":0108
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediSpecInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsʹ��˵�� As New Recordset
Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl
Private mstrDrugName As String  '��¼ҩƷ����
Private mstrDetails As String   '��¼ʹ��˵������
Private mblnOk As Boolean       '��¼ȷ����ť�Ƿ񱻵����
Private mblnCancel As Boolean   '��¼ȡ����ť�Ƿ񱻵����

Public Sub ShowMe(ByVal frmParent As Object, ByVal frmName As String)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Dim intRow As Integer
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandle
    
    gstrSql = "select ����,����,���� from ҩƷʹ��˵����Ŀ "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ҩƷʹ��˵����Ŀ")
    
    If rsTemp.EOF Then
        MsgBox "�㻹δ����ҩƷʹ��˵����Ŀ���뵽���ֵ�����ߡ����ã�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
        
    Call InitComandBars
    
    With vsfDetails
        '�������
        .Rows = rsTemp.RecordCount + 1
        .RowHeight(-1) = 270
        For intRow = 1 To rsTemp.RecordCount
            .TextMatrix(intRow, .ColIndex("��Ŀ")) = rsTemp!����
            .TextMatrix(intRow, .ColIndex("����")) = rsTemp!����
            rsTemp.MoveNext
        Next
        .Row = 1
    End With
    
    '��ʼ�����ݼ�
    With mrsʹ��˵��
        If .State = 1 Then .Close
        .Fields.Append "����", adVarChar, 5, adFldIsNullable
        .Fields.Append "��Ŀ", adVarChar, 20, adFldIsNullable
        .Fields.Append "����", adVarChar, 2000, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For intRow = 1 To vsfDetails.Rows - 1
            .AddNew
            !���� = vsfDetails.TextMatrix(intRow, vsfDetails.ColIndex("����"))
            !��Ŀ = vsfDetails.TextMatrix(intRow, vsfDetails.ColIndex("��Ŀ"))
            !���� = ""
            .Update
        Next
    End With
    
    Me.Caption = "��" & frmName & "��" & "ʹ��˵���༭"
    mstrDrugName = frmName
    
    Me.Show vbModal, frmParent
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 1
            Call CopyCol
    End Select
End Sub

Private Sub cmdCancel_Click()
    Dim blnChange As Boolean
    Dim intRow As Integer
    
    mblnCancel = True
    Call vsfDetails_LeaveCell
    
    With vsfDetails
        If .Rows > 1 Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) = "��" Then
                    blnChange = True
                    Exit For
                End If
            Next
        End If
        
        If blnChange Then
            If MsgBox("ʹ��˵�����޸��ˣ��Ƿ�ȷ���˳���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
                Set mrsʹ��˵�� = Nothing
                Unload Me
            Else
                mblnCancel = False
            End If
        Else
            Set mrsʹ��˵�� = Nothing
            Unload Me
        End If
    End With
End Sub

Private Sub cmdDown_Click()
    Dim str����  As String
    Dim str��Ŀ  As String
    Dim str���ݻ���  As String
    
    Call vsfDetails_LeaveCell
    
    With vsfDetails
        str���� = .TextMatrix(.Row + 1, .ColIndex("����"))
        str��Ŀ = .TextMatrix(.Row + 1, .ColIndex("��Ŀ"))
        str���ݻ��� = .TextMatrix(.Row + 1, .ColIndex("���ݻ���"))
        
        .TextMatrix(.Row + 1, .ColIndex("����")) = .TextMatrix(.Row, .ColIndex("����"))
        .TextMatrix(.Row + 1, .ColIndex("��Ŀ")) = .TextMatrix(.Row, .ColIndex("��Ŀ"))
        .TextMatrix(.Row + 1, .ColIndex("���ݻ���")) = .TextMatrix(.Row, .ColIndex("���ݻ���"))
        
        .TextMatrix(.Row, .ColIndex("����")) = str����
        .TextMatrix(.Row, .ColIndex("��Ŀ")) = str��Ŀ
        .TextMatrix(.Row, .ColIndex("���ݻ���")) = str���ݻ���
        
        mrsʹ��˵��.Filter = "����='" & .TextMatrix(.Row, .ColIndex("����")) & "'"
        rtbDetails.Text = mrsʹ��˵��!����
        
        .Row = .Row + 1
    End With
End Sub

Private Sub cmdOK_Click()
    mblnOk = True
    Call vsfDetails_LeaveCell
    Call GetDetails
    frmMediSpec.rtbDetails.Text = mstrDetails
    
    mstrDetails = ""
    Set mrsʹ��˵�� = Nothing
    Unload Me
End Sub

Private Sub cmdUp_Click()
    Dim str����  As String
    Dim str��Ŀ  As String
    Dim str���ݻ���  As String
    
    Call vsfDetails_LeaveCell
    
    With vsfDetails
        str���� = .TextMatrix(.Row - 1, .ColIndex("����"))
        str��Ŀ = .TextMatrix(.Row - 1, .ColIndex("��Ŀ"))
        str���ݻ��� = .TextMatrix(.Row - 1, .ColIndex("���ݻ���"))
        
        .TextMatrix(.Row - 1, .ColIndex("����")) = .TextMatrix(.Row, .ColIndex("����"))
        .TextMatrix(.Row - 1, .ColIndex("��Ŀ")) = .TextMatrix(.Row, .ColIndex("��Ŀ"))
        .TextMatrix(.Row - 1, .ColIndex("���ݻ���")) = .TextMatrix(.Row, .ColIndex("���ݻ���"))
        
        .TextMatrix(.Row, .ColIndex("����")) = str����
        .TextMatrix(.Row, .ColIndex("��Ŀ")) = str��Ŀ
        .TextMatrix(.Row, .ColIndex("���ݻ���")) = str���ݻ���
        
        mrsʹ��˵��.Filter = "����='" & .TextMatrix(.Row, .ColIndex("����")) & "'"
        rtbDetails.Text = mrsʹ��˵��!����
        
        .Row = .Row - 1
    End With
End Sub

Private Sub cmdView_Click()
    Call vsfDetails_LeaveCell
    Call GetDetails
    frmMediSpecInstructionView.rtbDetails.Text = mstrDetails
    Call frmMediSpecInstructionView.ShowMe(Me, mstrDrugName)
    Call vsfDetails_EnterCell
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call cmdCancel_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsfDetails.Move 60, 30, 1 / 3 * Me.ScaleWidth, Me.ScaleHeight - 30 - 300 - 200
    picLine.Move vsfDetails.Left + vsfDetails.Width, 30, 45, vsfDetails.Height
    rtbDetails.Move picLine.Left + picLine.Width, 30, Me.ScaleWidth - picLine.Left - picLine.Width - 60, vsfDetails.Height
    cmdUp.Move 200, vsfDetails.Top + vsfDetails.Height + 100
    cmdDown.Move 1500, vsfDetails.Top + vsfDetails.Height + 100
    cmdCancel.Move Me.ScaleWidth - 1300, cmdDown.Top
    cmdOK.Move Me.ScaleWidth - 2600, cmdDown.Top
    cmdView.Move Me.ScaleWidth - 4100, cmdDown.Top
End Sub

Private Sub CopyCol()
    'Ӧ���ڱ���
    Dim intRow As Integer
    
    With vsfDetails
        If .TextMatrix(.Row, .ColIndex("���ݻ���")) = "��" Then
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, .ColIndex("���ݻ���")) = "��"
            Next
        Else
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, .ColIndex("���ݻ���")) = ""
            Next
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnChange As Boolean
    Dim intRow As Integer
    
    If mblnOk Or mblnCancel Then Exit Sub
    
    With vsfDetails
        If .Rows > 1 Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) = "��" Then
                    blnChange = True
                    Exit For
                End If
            Next
        End If
        
        If blnChange Then
            If MsgBox("ʹ��˵�����޸��ˣ��Ƿ�ȷ���˳���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
                Set mrsʹ��˵�� = Nothing
                mblnCancel = False
                mblnOk = False
            Else
                Cancel = True
            End If
        Else
            Set mrsʹ��˵�� = Nothing
            mblnCancel = False
            mblnOk = False
        End If
    End With
End Sub

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfDetails.Width + x < 2000 Or vsfDetails.Width + x > Me.ScaleWidth - 4000 Then Exit Sub
        vsfDetails.Width = vsfDetails.Width + x
        picLine.Left = picLine.Left + x
        rtbDetails.Left = picLine.Left + picLine.Width
        rtbDetails.Width = rtbDetails.Width - x
    End If
End Sub

Private Sub rtbDetails_Change()
    With vsfDetails
        If Trim(rtbDetails.Text) <> "" Then
            .TextMatrix(.Row, 0) = "��"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
    End With
End Sub

Private Sub vsfDetails_DblClick()
    With vsfDetails
        If .Col = .ColIndex("���ݻ���") Then
            If .TextMatrix(.Row, .ColIndex("���ݻ���")) = "" Then
                .TextMatrix(.Row, .ColIndex("���ݻ���")) = "��"
            Else
                .TextMatrix(.Row, .ColIndex("���ݻ���")) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfDetails_EnterCell()
    With vsfDetails
        If .Rows = 1 Then Exit Sub
        
        cmdUp.Enabled = True: cmdDown.Enabled = True
        If .Row = 1 Then cmdUp.Enabled = False: cmdDown.Enabled = True
        If .Row = .Rows - 1 Then cmdUp.Enabled = True: cmdDown.Enabled = False
        If .Rows = 2 Then cmdUp.Enabled = False: cmdDown.Enabled = False
        
        If mrsʹ��˵��.State = 1 Then
            mrsʹ��˵��.Filter = "����='" & .TextMatrix(.Row, .ColIndex("����")) & "'"
            rtbDetails.Text = mrsʹ��˵��!����
        End If
    End With
End Sub

Private Sub vsfDetails_LeaveCell()
    With vsfDetails
        If .Rows = 1 Then Exit Sub
        
        If mrsʹ��˵��.State = 1 Then
            mrsʹ��˵��.Filter = "����='" & .TextMatrix(.Row, .ColIndex("����")) & "'"
            mrsʹ��˵��!���� = rtbDetails.Text
        End If
    End With
End Sub

Private Sub GetDetails()
    '��ȡʹ��˵������
    Dim intRow As Integer
    Dim strDetails As String
    
    With vsfDetails
        If .Rows = 1 Then Exit Sub
        mstrDetails = ""
        mstrDetails = frmMediSpec.rtbDetails.Text
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) = "��" Then
                mrsʹ��˵��.Filter = "����='" & .TextMatrix(intRow, .ColIndex("����")) & "'"
                
                If .TextMatrix(intRow, .ColIndex("���ݻ���")) = "��" Then
                    strDetails = strDetails & "��" & mrsʹ��˵��!��Ŀ & "��" & vbCrLf & mrsʹ��˵��!���� & vbCrLf
                Else
                    strDetails = strDetails & "��" & mrsʹ��˵��!��Ŀ & "��" & mrsʹ��˵��!���� & vbCrLf
                End If
            End If
        Next
        
        If strDetails <> "" Then
            mstrDetails = mstrDetails & vbCrLf & strDetails
        End If
    End With
End Sub

Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If vsfDetails.Rows = 1 Then Exit Sub
    If Button <> 2 Then Exit Sub
    
    If vsfDetails.Col = vsfDetails.ColIndex("���ݻ���") Then
        mobjPopup.ShowPopup
    End If
End Sub

Private Sub InitComandBars()
    '��ʼ��ComandBars�������˵�
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003
    
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
    End With
    
    '�Ҽ��˵�
    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopup.Controls
        Set mobjControl = .Add(xtpControlButton, 1, "Ӧ���ڱ���")
    End With
End Sub

