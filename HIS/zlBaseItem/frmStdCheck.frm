VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStdCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��׼�˲�"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   Icon            =   "frmStdCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6660
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd 
      Caption         =   "��ϸ"
      Height          =   350
      Index           =   2
      Left            =   5460
      TabIndex        =   10
      Top             =   2325
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   16
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "���(&K)"
      Height          =   350
      Left            =   3960
      TabIndex        =   14
      Top             =   3945
      Width           =   1100
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "��ӡ(&P)..."
      Enabled         =   0   'False
      Height          =   350
      Left            =   165
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��ϸ"
      Height          =   350
      Index           =   3
      Left            =   5460
      TabIndex        =   13
      Top             =   3075
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��ϸ"
      Height          =   350
      Index           =   1
      Left            =   5460
      TabIndex        =   7
      Top             =   1620
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmd 
      Caption         =   "��ϸ"
      Height          =   350
      Index           =   0
      Left            =   5460
      TabIndex        =   4
      Top             =   915
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "�˲�ҽԺ������׼ҽ�۵���Ŀ"
      Height          =   300
      Index           =   3
      Left            =   420
      TabIndex        =   11
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "�˲�ҽԺͣ�ö���׼ҽ��δע������Ŀ"
      Height          =   300
      Index           =   2
      Left            =   420
      TabIndex        =   8
      Top             =   2370
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "�˲�ҽԺ���õ���׼ҽ���Ѿ�ע������Ŀ"
      Height          =   300
      Index           =   1
      Left            =   420
      TabIndex        =   5
      Top             =   1665
      Value           =   1  'Checked
      Width           =   3570
   End
   Begin VB.CheckBox chkCaption 
      Caption         =   "�˲�ҽԺδ��ȷ��Ӧ��׼ҽ�۵���Ŀ"
      Height          =   300
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   945
      Value           =   1  'Checked
      Width           =   3180
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   15
      TabIndex        =   15
      Top             =   3780
      Width           =   6675
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   6675
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2100
      Index           =   1
      Left            =   660
      TabIndex        =   19
      Top             =   1980
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2100
      Index           =   0
      Left            =   660
      TabIndex        =   18
      Top             =   1275
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2280
      Index           =   2
      Left            =   660
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3030
      Index           =   3
      Left            =   660
      TabIndex        =   21
      Top             =   30
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   5345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   3
      Left            =   4110
      TabIndex        =   12
      Top             =   3165
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   2
      Left            =   4110
      TabIndex        =   9
      Top             =   2430
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   1
      Left            =   4110
      TabIndex        =   6
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   0
      Left            =   4110
      TabIndex        =   3
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ����Ҫ�˲������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   420
      TabIndex        =   0
      Top             =   210
      Width           =   2550
   End
   Begin VB.Menu mnuPop 
      Caption         =   "������ӡ"
      Visible         =   0   'False
      Begin VB.Menu mnuPopExcel 
         Caption         =   "�����(&E)xcel"
      End
      Begin VB.Menu mnuPopPreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuPopPrint 
         Caption         =   "�������ӡ��(&P)"
      End
   End
End
Attribute VB_Name = "frmStdCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mstrCol As String = "����,1200,0,2;����,1500,0,0;��λ,1000,0,0"
Private Const mstrCol1 As String = "����,1200,0,2;����,1500,0,0;��λ,1000,0,0;�ּ�,1000,1,0;���,1000,1,0;���,1000,1,0"
Private mintColumn As Integer, mintColumn1 As Integer, mintColumn2 As Integer, mintColumn3 As Integer

Private Sub chkCaption_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim i As Long
    For i = Me.lvw.LBound To Me.lvw.UBound
        If i = Index Then
            Me.lvw(Index).Visible = Not Me.lvw(Index).Visible
            If Me.lvw(Index).Visible Then
                Me.lvw(Index).ZOrder
                Me.lvw(Index).SetFocus
                Me.cmdReport.Enabled = True
                Me.cmdReport.Tag = i
            Else
                Me.cmdReport.Enabled = False
            End If
        Else
            Me.lvw(i).Visible = False
        End If
    Next
End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCheck_Click()
On Error GoTo errHandle
    Dim i As Long
    Dim strSQL As String
    Dim strTmp As String
    Dim ObjItem As ListItem
    Dim rsTmp As New ADODB.Recordset
    Dim blnHave As Boolean
    
    Me.cmdClose.Enabled = False
    Me.cmdCheck.Enabled = False
    zlCommFun.ShowFlash "��ʼ�������..."
    '�ȳ�ʼ��
    For i = Me.lvw.LBound To Me.lvw.UBound
        Me.lvw(i).ListItems.Clear
        Me.cmd(i).Visible = False
        If Me.chkCaption(i).value = 1 And blnHave = False Then
            blnHave = True
        End If
        Me.lvw(i).Sorted = False
        Me.lvw(i).Visible = False
        Me.lblCaption(i).Caption = ""
    Next
    Me.cmdReport.Enabled = False
    If blnHave = False Then
        MsgBox "��ѡ��Ҫ�˲�����ݣ�", vbInformation, gstrSysName
        Me.chkCaption(0).SetFocus
        Me.cmdClose.Enabled = True
        Me.cmdCheck.Enabled = True
        Exit Sub
    End If
    '�������ý��к˲�
    strSQL = ""
    For i = Me.chkCaption.LBound To Me.chkCaption.UBound
        If Me.chkCaption(i).value = 1 Then
            Select Case True
                Case i = 0   'δ��ȷ��Ӧ��׼ҽ����Ŀ��
                    strTmp = " SELECT 1 ��������, 'δ��ȷ��Ӧ' ˵��,A.ID,A.����,A.����,A.���㵥λ ��λ,B.����޼�,B.����޼�,0 �ּ� " & vbCrLf & _
                           " FROM �շ���ĿĿ¼ A, ��׼ҽ�۹淶 B  " & vbCrLf & _
                           " WHERE A.��� <> '5' AND A.��� <> '6' AND A.��� <> '7' " & vbCrLf & _
                           " AND A.��ʶ����=B.��Ŀ����(+) AND B.��Ŀ���� IS NULL "
                Case i = 1   'ҽԺ���õ���׼���Ѿ�ע����
                    strTmp = "SELECT 2 ��������, 'ҽԺ���ñ�׼��ע��' ˵��,A.ID,A.����,A.����,A.���㵥λ ��λ,B.����޼�,B.����޼�,0 �ּ� " & vbCrLf & _
                            "  FROM �շ���ĿĿ¼ A, ��׼ҽ�۹淶 B " & vbCrLf & _
                            "  WHERE A.��� <> '5' AND A.��� <> '6' AND A.��� <> '7' AND A.��ʶ����=B.��Ŀ����  " & vbCrLf & _
                            "    AND NVL(A.����ʱ��,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD') AND LTRIM(RTRIM(NVL(B.ע����־,'0')))='1' "
                Case i = 2   'ҽԺͣ�õ���׼ҽ��δע����
                    strTmp = "SELECT 3 ��������, 'ҽԺ��ͣ�ñ�׼������' ˵��,A.ID,A.����,A.����,A.���㵥λ ��λ,B.����޼�,B.����޼�,0 �ּ�  " & vbCrLf & _
                            "  FROM �շ���ĿĿ¼ A, ��׼ҽ�۹淶 B " & vbCrLf & _
                            "  WHERE A.��� <> '5' AND A.��� <> '6' AND A.��� <> '7' AND A.��ʶ����=B.��Ŀ����  " & vbCrLf & _
                            "    AND NVL(A.����ʱ��,TO_DATE('3000-01-01','YYYY-MM-DD'))<>TO_DATE('3000-01-01','YYYY-MM-DD') AND NOT (LTRIM(RTRIM(NVL(B.ע����־,'0')))='1') "
                Case i = 3   'ҽԺ�۸����׼ҽ�۲����ϵ�
                    strTmp = " SELECT 4 ��������, '�������ͼ۲���' ˵��,A.ID,A.����,A.����,A.���㵥λ ��λ,B.����޼�,B.����޼�,0 �ּ�  " & vbCrLf & _
                            "   FROM �շ���ĿĿ¼ A, ��׼ҽ�۹淶 B " & vbCrLf & _
                            "   WHERE A.��� <> '5' AND A.��� <> '6' AND A.��� <> '7' AND A.��ʶ����=B.��Ŀ����  " & vbCrLf & _
                            "     AND (A.����޼�<>B.����޼� OR A.����޼�<>B.����޼�) " & vbCrLf & _
                            " UNION ALL " & vbCrLf & _
                            " SELECT 5 ��������, '��ǰ�۸񲻷�' ˵��,C.ID,C.����,C.����,C.���㵥λ ��λ,B.����޼�,B.����޼�,SUM(A.�ּ�) �ּ� " & vbCrLf & _
                            "   FROM �շѼ�Ŀ A,��׼ҽ�۹淶 B,  �շ���ĿĿ¼ C  " & vbCrLf & _
                            "  WHERE A.�շ�ϸĿID = C.ID  AND NVL(C.�Ƿ���,0)=0  AND  C.��ʶ����=B.��Ŀ����  " & vbCrLf & _
                            "    AND A.ִ������<=SYSDATE AND (A.��ֹ����>=SYSDATE OR A.��ֹ���� IS NULL)   " & vbCrLf & _
                            " GROUP BY C.ID,C.����,C.����,C.���㵥λ ,B.����޼�,B.����޼�,A.�۸�ȼ� " & vbCrLf & _
                            " HAVING NOT (SUM(A.�ּ�) >=B.����޼� AND SUM(A.�ּ�)<=B.����޼�) "
            End Select
            strSQL = strSQL & " UNION ALL " & vbCrLf & strTmp
        End If
    Next
    strSQL = Mid(strSQL, 11)
    strSQL = "select * from (" & strSQL & ")  order by ��������,����"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        If Me.chkCaption(0).value = 1 Then
            'δ��ȷ��Ӧ��׼ҽ����Ŀ��
            rsTmp.Filter = "��������=1"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "���ڴ����˲�ҽԺδ��ȷ��Ӧ��׼ҽ�۵���Ŀ��..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(0).ListItems.Add(, , zlCommFun.Nvl(rsTmp!����))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!����)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!��λ)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(1).value = 1 Then
            'ҽԺ���õ���׼���Ѿ�ע����
            rsTmp.Filter = "��������=2"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "���ڴ����˲�ҽԺ���õ���׼ҽ���Ѿ�ע������Ŀ��..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(1).ListItems.Add(, , zlCommFun.Nvl(rsTmp!����))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!����)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!��λ)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(2).value = 1 Then
            'ҽԺͣ�õ���׼ҽ��δע����
            rsTmp.Filter = "��������=3"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "���ڴ����˲�ҽԺͣ�ö���׼ҽ��δע������Ŀ��..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(2).ListItems.Add(, , zlCommFun.Nvl(rsTmp!����))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!����)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!��λ)
                    rsTmp.MoveNext
                Next
            End If
        End If
        If Me.chkCaption(3).value = 1 Then
            '�������������޼�
            rsTmp.Filter = "��������=4"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "���ڴ����˲�ҽԺ������׼ҽ�۵����������޼۵���Ŀ��..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    strSQL = "ZL_�շ�ϸĿ��׼�޼�_UPDATE(" & rsTmp!ID & "," & rsTmp!����޼� & "," & rsTmp!����޼� & ")"
                    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                    rsTmp.MoveNext
                Next
            End If
            '��ʾ������ҽ�۵���Ŀ
            rsTmp.Filter = "��������=5"
            If rsTmp.RecordCount > 0 Then
                zlCommFun.ShowFlash "���ڴ����˲�ҽԺ������׼ҽ�۵���Ŀ��..."
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = Me.lvw(3).ListItems.Add(, , zlCommFun.Nvl(rsTmp!����))
                    ObjItem.SubItems(1) = zlCommFun.Nvl(rsTmp!����)
                    ObjItem.SubItems(2) = zlCommFun.Nvl(rsTmp!��λ)
                    ObjItem.SubItems(3) = CStr(Format(zlCommFun.Nvl(rsTmp!�ּ�, 0), "0.00"))
                    ObjItem.SubItems(4) = CStr(Format(zlCommFun.Nvl(rsTmp!����޼�, 0), "0.00"))
                    ObjItem.SubItems(5) = CStr(Format(zlCommFun.Nvl(rsTmp!����޼�, 0), "0.00"))
                    rsTmp.MoveNext
                Next
            End If
        End If
    End If
    '������ʾ״̬
    For i = Me.lvw.LBound To Me.lvw.UBound
        If Me.lvw(i).ListItems.Count > 0 Then
            Me.cmd(i).Visible = True
            Me.lblCaption(i).Caption = "(" & Me.lvw(i).ListItems.Count & " ��)"
        Else
            If Me.chkCaption(i).value = 1 Then
                Me.lblCaption(i).Caption = "(��)"
            Else
                Me.lblCaption(i).Caption = ""
            End If
        End If
    Next
    Me.cmdClose.Enabled = True
    Me.cmdCheck.Enabled = True
    zlCommFun.ShowFlash
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.cmdClose.Enabled = True
    Me.cmdCheck.Enabled = True
    zlCommFun.ShowFlash
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReport_Click()
    PopupMenu mnuPop
End Sub

Private Sub Form_Load()
    Dim i As Long
    For i = Me.lvw.LBound To Me.lvw.UBound
        If i = Me.lvw.UBound Then
            zlControl.LvwSelectColumns Me.lvw(i), mstrCol1, True
        Else
            zlControl.LvwSelectColumns Me.lvw(i), mstrCol, True
        End If
    Next
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvw(Index).Sorted = True
    If Choose(Index + 1, mintColumn, mintColumn1, mintColumn2, mintColumn3) = ColumnHeader.Index - 1 Then
        lvw(Index).SortOrder = IIF(lvw(Index).SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Select Case Index
            Case 0
                mintColumn = ColumnHeader.Index - 1
            Case 1
                mintColumn1 = ColumnHeader.Index - 1
            Case 2
                mintColumn2 = ColumnHeader.Index - 1
            Case 3
                mintColumn3 = ColumnHeader.Index - 1
        End Select
        lvw(Index).SortKey = ColumnHeader.Index - 1
        lvw(Index).SortOrder = lvwAscending
    End If
End Sub

Private Sub subPrint(ByVal Index As Long, bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    objPrint.Title.Text = "��׼�۸�˲�"
    Set objPrint.Body.objData = Me.lvw(Index)
    Select Case Index
        Case 0
            objPrint.UnderAppItems.Add "�˲�ҽԺδ��ȷ��Ӧ��׼ҽ�۵���Ŀ"
        Case 1
            objPrint.UnderAppItems.Add "�˲�ҽԺ���õ���׼ҽ���Ѿ�ע������Ŀ"
        Case 2
            objPrint.UnderAppItems.Add "�˲�ҽԺͣ�ö���׼ҽ��δע������Ŀ"
        Case 3
            objPrint.UnderAppItems.Add "�˲�ҽԺ������׼ҽ�۵���Ŀ"
    End Select
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Private Sub mnuPopExcel_Click()
    '�����Excel
    Call subPrint(cmdReport.Tag, 3)
End Sub

Private Sub mnuPopPreview_Click()
    '��ӡԤ��
    Call subPrint(cmdReport.Tag, 2)
End Sub

Private Sub mnuPopPrint_Click()
    '�������ӡ��
    Call subPrint(cmdReport.Tag, 1)
End Sub
