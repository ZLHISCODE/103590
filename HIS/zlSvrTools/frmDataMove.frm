VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataMove 
   BackColor       =   &H80000005&
   Caption         =   "���ݹ鵵ת��"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmDataMove.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDay 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2970
      MaxLength       =   3
      TabIndex        =   9
      Top             =   3600
      Width           =   645
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ת��(&I)��"
      Height          =   350
      Index           =   1
      Left            =   4620
      TabIndex        =   13
      Top             =   4035
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��ԭ����(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   945
      TabIndex        =   12
      Top             =   4035
      Width           =   1170
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "��������(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2115
      TabIndex        =   11
      Top             =   4035
      Width           =   1170
   End
   Begin MSComCtl2.UpDown udDay 
      Height          =   300
      Left            =   3615
      TabIndex        =   10
      Top             =   3600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   150
      BuddyControl    =   "txtDay"
      BuddyDispid     =   196617
      OrigLeft        =   3435
      OrigTop         =   3600
      OrigRight       =   3675
      OrigBottom      =   3915
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdGroup 
      Height          =   2565
      Left            =   975
      TabIndex        =   5
      Top             =   975
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "ת��(&O)��"
      Height          =   350
      Index           =   0
      Left            =   3465
      TabIndex        =   4
      Top             =   4035
      Width           =   1170
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   615
      Width           =   3570
   End
   Begin VB.Line LinCondition 
      X1              =   1575
      X2              =   1575
      Y1              =   3540
      Y2              =   3960
   End
   Begin VB.Shape shpCondition 
      Height          =   450
      Left            =   975
      Top             =   3525
      Width           =   4305
   End
   Begin VB.Label lblRelation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����            ��Ϊ��ʷ����"
      Height          =   180
      Left            =   2565
      TabIndex        =   8
      Top             =   3660
      Width           =   2520
   End
   Begin VB.Label lblColumn 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�շ�ʱ��"
      Height          =   180
      Left            =   1815
      TabIndex        =   7
      Top             =   3660
      Width           =   720
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   1095
      TabIndex        =   6
      Top             =   3675
      Width           =   360
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   1050
      Left            =   960
      TabIndex        =   3
      Top             =   4470
      Width           =   4320
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݹ鵵ת��"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmDataMove.frx":04F9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmDataMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim intCount As Integer

Private Sub cmbSystem_Click()
    
    With rsTemp
        cmbSystem.Tag = GetOwnerName(Val(cmbSystem.ItemData(cmbSystem.ListIndex)), gcnOracle)
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_zlDatamove", Val(cmbSystem.ItemData(cmbSystem.ListIndex)))
        If Not .EOF And Not .BOF Then
            udDay.Enabled = True
            cmdExecute(0).Enabled = True
            cmdExecute(1).Enabled = True
        Else
            udDay.Enabled = False
            cmdExecute(0).Enabled = False
            cmdExecute(1).Enabled = False
        End If
    End With
    Set hgdGroup.Recordset = rsTemp
    With hgdGroup
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .ColWidth(2) = 4600
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        If .Rows = 1 Then
            .Rows = 2
            .FixedRows = 1
        End If
        .Row = 1
        .Col = 0
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        Call hgdGroup_RowColChange
    End With
End Sub

Private Sub cmdCancel_Click()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    Call hgdGroup_RowColChange
End Sub

Private Sub cmdExecute_Click(Index As Integer)
    '--------------------------------------------
    '���ܣ�������ת�Ƶ���ʷ���ݱ���
    '--------------------------------------------
    Dim lngSystem As Long
    
    If MsgBox("���ת�����ݽ϶࣬������Ҫ�ϳ�ʱ�䡣" & vbCr & vbCr & "����ִ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
    lngSystem = cmbSystem.ItemData(cmbSystem.ListIndex) \ 100
    If Index = 0 Then
        frmMDIMain.stbThis.Panels(2).Text = "���ڽ���" & hgdGroup.TextMatrix(hgdGroup.Row, 1) & "(ת��)��"
        strSQL = cmbSystem.Tag & ".zl" & lngSystem & _
                "_DataMoveOut" & hgdGroup.TextMatrix(hgdGroup.Row, 0)
        strSQL = strSQL & "(" & udDay.Value & ")"
    Else
        frmMDIMain.stbThis.Panels(2).Text = "���ڽ���" & hgdGroup.TextMatrix(hgdGroup.Row, 1) & "(ת��)��"
        strSQL = cmbSystem.Tag & ".zl" & lngSystem & _
                "_DataMoveIn" & hgdGroup.TextMatrix(hgdGroup.Row, 0)
        strSQL = strSQL & "(" & udDay.Value & ")"
    End If
    
    MousePointer = 11
    gcnOracle.Execute strSQL, , adCmdStoredProc
    MousePointer = 0
    
    frmMDIMain.stbThis.Panels(2).Text = ""
    MsgBox "����ת����ϣ�", vbExclamation, gstrSysName

End Sub

Private Sub cmdSave_Click()
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    
    If IsNumeric(txtDay.Text) = False Then
        MsgBox "��������ȷ��������", vbInformation, gstrSysName
        txtDay.SetFocus
        Exit Sub
    End If
    
    If Val(txtDay.Text) < udDay.Min Or Val(txtDay.Text) > udDay.Max Then
        MsgBox "�������ܳ���1000����Ϊ������", vbInformation, gstrSysName
        txtDay.SetFocus
        Exit Sub
    End If
    hgdGroup.TextMatrix(hgdGroup.Tag, 4) = Val(txtDay.Text)
    strSQL = "update zlDataMove" & _
            " set ת������=" & Val(txtDay.Text) & _
            " where ϵͳ=" & cmbSystem.ItemData(cmbSystem.ListIndex) & " and ���=" & hgdGroup.TextMatrix(hgdGroup.Tag, 0)
    gcnOracle.Execute strSQL
End Sub

Private Sub dtpStart_Change()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Form_Load()
    lblMain.Caption = "Ϊ����ϵͳǰ̨�����Ч���У����Խ�����ʹ�õ���ʷ����ת�Ʊ��浽��ʷ���ݱ��У���ͬʱҲ������ϵͳ�������ؽ������ȶ���ϵͳά������Ŀ��ٽ��У���֤����ϵͳ���е������ԡ�" & _
        vbCrLf & vbCrLf & "�����Ҫ�����Զ�ѭ�����ִ������ת�ƹ鵵���������ú�̨��ҵ��������й����ù��ߡ�"
    On Error GoTo ErrHandle
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If
    
    With rsTemp
        Do While Not .EOF
            cmbSystem.AddItem !���� & " v" & !�汾�� & "��" & !��� & "��"
            cmbSystem.ItemData(cmbSystem.NewIndex) = !���
            .MoveNext
        Loop
        If cmbSystem.ListCount = 0 Then
            udDay.Enabled = False
            cmdExecute(0).Enabled = False
            cmdExecute(1).Enabled = False
            cmdSave.Enabled = False
            cmdCancel.Enabled = False
        End If
        If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
        If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    End With
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    With hgdGroup
        If ScaleWidth - .Left - 200 > cmbSystem.Left + cmbSystem.Width - .Left Then
            .Width = ScaleWidth - .Left - 200
        Else
            .Width = cmbSystem.Left + cmbSystem.Width - .Left
        End If
    End With
    shpCondition.Width = hgdGroup.Width
    cmdExecute(1).Left = hgdGroup.Left + hgdGroup.Width - cmdExecute(1).Width
    cmdExecute(0).Left = cmdExecute(1).Left - cmdExecute(0).Width
    With lblMain
        .Top = cmdExecute(1).Top + cmdExecute(1).Height + 200
        .Height = ScaleHeight - .Top - 100
        .Left = lblSys.Left
        .Width = ScaleWidth - .Left - imgMain.Left
    End With
    
End Sub

Private Sub hgdGroup_RowColChange()
    If Val(hgdGroup.TextMatrix(hgdGroup.Row, 0)) = 0 Then Exit Sub
    If cmdSave.Enabled = True Then
        If hgdGroup.TextMatrix(hgdGroup.Tag, 4) <> udDay.Value Then
            If MsgBox("��" & hgdGroup.TextMatrix(hgdGroup.Tag, 1) & "��������øı��δ���棬�Ƿ񱣴棿", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                Call cmdSave_Click
            Else
                Call cmdCancel_Click
            End If
        End If
    End If
    With hgdGroup
        .Tag = .Row
        lblColumn.Caption = .TextMatrix(.Row, 3)
        udDay.Value = .TextMatrix(.Row, 4)
    End With
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub txtDay_Change()
    Call udDay_Change
End Sub

Private Sub udDay_Change()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub updSpace_Change()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub

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
    
    objPrint.Title.Text = "���ݹ鵵ת��"
    
    objRow.Add "Ӧ��ϵͳ��" & cmbSystem.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    Set objPrint.Body = hgdGroup
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

