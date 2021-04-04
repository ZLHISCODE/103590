VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLeaveSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ��ҩƷѡ��"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   Icon            =   "frmLeaveSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Ctrl+A"
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   1350
      TabIndex        =   3
      ToolTipText     =   "Ctrl+R"
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7920
      TabIndex        =   2
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9030
      TabIndex        =   1
      Top             =   5520
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   5325
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   10440
      _cx             =   18415
      _cy             =   9393
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLeaveSelect.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
Attribute VB_Name = "frmLeaveSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As frmLeaveMediMana
Private mLeaveRecord As New ADODB.Recordset

Public Function LeaveSelect(ByVal frmMain As frmLeaveMediMana, ByVal strSQL As String)
    '
    Dim i As Integer
    On Error GoTo errHandle
    Set mfrmMain = frmMain
    Set mLeaveRecord = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not mLeaveRecord.EOF Then
        Call FillVsList
        Me.Show vbModal, frmMain
    Else
        MsgBox "û�пɹ�ѡ������ݣ�", vbInformation, gstrSysName
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdAll_Click()
    Dim lngRow As Long
    With vsList
        For lngRow = .FixedRows To .Rows - 1
            If Not (Val(.TextMatrix(.Rows - 1, .ColIndex("����"))) = 0 Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������"))) = "") Then
                vsList.Cell(flexcpChecked, lngRow, 0) = flexChecked
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    vsList.Cell(flexcpChecked, 1, 0, vsList.Rows - 1, 0) = flexUnchecked
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer, strItem As String
    Dim blnAdd As Boolean
    Dim lngRow As Long
    
    lngRow = mfrmMain.vsList.Row
    For i = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, i, 0) = flexChecked Then
        
               strItem = "ҽ��" & vbTab & vsList.TextMatrix(i, vsList.ColIndex("ҩƷ���������")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("���")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("��;")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("����")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("���㵥λ")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("����")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("���")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("ҩƷID")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("ҽ��ID")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("���ͺ�")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("������λ")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("����ϵ��")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("���ﵥλ")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("�����װ")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("����")) & vbTab & _
                        vsList.TextMatrix(i, vsList.ColIndex("�ɴ�����"))
                mfrmMain.vsList.AddItem strItem, mfrmMain.vsList.Rows - 1
                mfrmMain.vsList.Select mfrmMain.vsList.Row + 1, 0
            blnAdd = True
        End If
    Next
    
    If blnAdd = True Then
        Unload Me
    End If
End Sub

Private Sub FillVsList()
    Dim strHead As String
    Dim lngLast���ID As Long, cur�������� As Currency, cur��ִ�д�  As Currency, date���� As Date
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo hErr
    strHead = ",300,1;����ʱ��,1500,1;NO,900,1;ҩƷ���������,2500,1;���,1600,1;��;,550,1;����,750,7;���㵥λ,450,4;����,750,7;���,1000,7;" & _
              "ҩƷID,0,1;ҽ��ID,0,1;���ͺ�,0,1;������λ,0,1;����ϵ��,0,1;���ﵥλ,0,1;�����װ,0,1;����,0,1;�ɴ�����,0,1"
    Call SetVsFlexGridHead(strHead, vsList)
    With vsList
        '�ϲ���Ԫ��
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(.ColIndex("����ʱ��")) = True
        .MergeCol(.ColIndex("NO")) = True
        .AutoSize 1, .ColIndex("���")
    End With
    
    Do Until mLeaveRecord.EOF
        With vsList
            
            
            '���������
            cur�������� = 0
            date���� = zlDatabase.Currentdate
            strSQL = "Select Min(�Ǽ�ʱ��) as ���� From �ݴ�ҩƷ��¼ Where ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & mLeaveRecord.Fields("ҽ��ID")), Val("" & mLeaveRecord.Fields("���ͺ�")))
            Do Until rsTmp.EOF
                date���� = IIf(IsNull(rsTmp!����), zlDatabase.Currentdate, rsTmp!����)
                rsTmp.MoveNext
            Loop
            
            strSQL = "Select Sum(Nvl(A.��������, 0)) As ��������" & vbNewLine & _
                    "From  ����ҽ��ִ�� A" & vbNewLine & _
                    "Where A.ҽ��id = [1] And A.���ͺ� = [2] And A.ִ��ʱ��  < [3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & mLeaveRecord.Fields("���ID")), Val("" & mLeaveRecord.Fields("���ͺ�")), date����)
            Do Until rsTmp.EOF
                cur��ִ�д� = Val("" & rsTmp!��������)
                rsTmp.MoveNext
            Loop
'            cur�������� = (cur��ִ�д� * Val(mLeaveRecord.Fields("��������"))) / Val("" & mLeaveRecord.Fields("����ϵ��"))
            If zlCommFun.NVL(mLeaveRecord.Fields("����ɷ����"), 0) = 0 Then
                cur�������� = (cur��ִ�д� * Val(mLeaveRecord.Fields("��������"))) / Val("" & mLeaveRecord.Fields("����ϵ��"))
            Else
                '���ﲻ�ɷ��㣬Abs����������ȡ��
                cur�������� = cur��ִ�д� * Abs(Int(0 - Val(mLeaveRecord.Fields("��������")) / Val("" & mLeaveRecord.Fields("����ϵ��"))))
            End If
            
            If Val("" & mLeaveRecord.Fields("�ɴ�����")) - cur�������� - Val(mLeaveRecord.Fields("�Ѵ�����")) > 0 Then
                If lngLast���ID <> 0 And lngLast���ID <> Val(mLeaveRecord.Fields("���ID")) Then
                    .AddItem ""
                    .RowHidden(.Rows - 2) = True
                End If
                lngLast���ID = Val(mLeaveRecord.Fields("���ID"))
                .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(mLeaveRecord.Fields("����ʱ��"), "yy-MM-dd hh:mm")
                .TextMatrix(.Rows - 1, .ColIndex("NO")) = mLeaveRecord.Fields("NO")
                .TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������")) = "[" & mLeaveRecord.Fields("����") & "]" & mLeaveRecord.Fields("����")
                .TextMatrix(.Rows - 1, .ColIndex("���")) = mLeaveRecord.Fields("���")
                Select Case mLeaveRecord.Fields("��;")
                    Case 1
                        .TextMatrix(.Rows - 1, .ColIndex("��;")) = "��Һ"
                    Case 2
                        .TextMatrix(.Rows - 1, .ColIndex("��;")) = "ע��"
                    Case 3
                        .TextMatrix(.Rows - 1, .ColIndex("��;")) = "Ƥ��"
                    Case Else
                        .TextMatrix(.Rows - 1, .ColIndex("��;")) = "����"
                End Select

                
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Val("" & mLeaveRecord.Fields("�ɴ�����")) - cur�������� - Val(mLeaveRecord.Fields("�Ѵ�����"))
                .TextMatrix(.Rows - 1, .ColIndex("���㵥λ")) = "" & mLeaveRecord.Fields("���㵥λ")
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(Val("" & mLeaveRecord.Fields("�ּ�")), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("���")) = Format((Val("" & mLeaveRecord.Fields("�ɴ�����")) - cur�������� - Val("" & mLeaveRecord.Fields("�Ѵ�����"))) * Val("" & mLeaveRecord.Fields("�ּ�")), "0.00")
                
                .TextMatrix(.Rows - 1, .ColIndex("ҩƷID")) = "" & mLeaveRecord.Fields("�շ�ϸĿID")
                .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = "" & mLeaveRecord.Fields("ҽ��ID")
                .TextMatrix(.Rows - 1, .ColIndex("���ͺ�")) = "" & mLeaveRecord.Fields("���ͺ�")
                .TextMatrix(.Rows - 1, .ColIndex("������λ")) = "" & mLeaveRecord.Fields("������λ")
                .TextMatrix(.Rows - 1, .ColIndex("����ϵ��")) = "" & mLeaveRecord.Fields("����ϵ��")
                .TextMatrix(.Rows - 1, .ColIndex("���ﵥλ")) = "" & mLeaveRecord.Fields("���ﵥλ")
                .TextMatrix(.Rows - 1, .ColIndex("�����װ")) = "" & mLeaveRecord.Fields("�����װ")
                .TextMatrix(.Rows - 1, .ColIndex("����")) = Val("" & mLeaveRecord.Fields("����") + 0)
                .TextMatrix(.Rows - 1, .ColIndex("�ɴ�����")) = Val("" & mLeaveRecord.Fields("�ɴ�����")) - cur�������� - Val("" & mLeaveRecord.Fields("�Ѵ�����"))
                .AddItem ""
            End If
        End With
        mLeaveRecord.MoveNext
    Loop
    If vsList.Rows > 2 Then
        vsList.RemoveItem (vsList.Rows - 1)
    End If
    vsList.Cell(flexcpChecked, 1, 0, vsList.Rows - 1, 0) = flexUnchecked 'ȫ����Ϊδѡ
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_Click()
    If vsList.MouseCol = 0 Then
        With vsList
            If Val(.TextMatrix(.Rows - 1, .ColIndex("����"))) = 0 Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������"))) = "" Then Exit Sub
        End With
        vsList.Cell(flexcpChecked, vsList.Row, 0) = IIf(vsList.Cell(flexcpChecked, vsList.Row, 0) = flexUnchecked, flexChecked, flexUnchecked)
    End If
End Sub



