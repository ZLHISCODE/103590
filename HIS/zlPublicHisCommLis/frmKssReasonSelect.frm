VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmKssReasonSelect 
   BorderStyle     =   0  'None
   Caption         =   "������ҩ����"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2970
      Left            =   0
      ScaleHeight     =   2940
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   0
         TabIndex        =   5
         Top             =   2520
         Width           =   6255
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Top             =   2595
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   2595
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   300
         Left            =   5280
         TabIndex        =   1
         Top             =   2595
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsgMain 
         Height          =   2535
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6255
         _cx             =   11033
         _cy             =   4471
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssReasonSelect.frx":0000
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
End
Attribute VB_Name = "frmKssReasonSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrName As String   '���ص���ҩ��������
Private mstrFind As String
Private mlngleft As Long
Private mlngTop As Long
Private mintType As Integer
Private Enum COL������ҩ����
    col���� = 0
    col���� = 1
    col���� = 2
End Enum

Public Function ShowMe(frmParent As Object, ByVal strFind As String, ByRef blnCancle As Boolean, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal intType As Integer) As String
'���أ���ҩ��������
'������strFind -Ϊ����������У��������strFind���Ҽ��룬���룬����
'      intType 1-������ҩ���ɣ�2-��������
    mstrFind = strFind
    mlngleft = lngLeft
    mlngTop = lngTop
    mintType = intType
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    blnCancle = Not mblnOk
    If mblnOk Then
        ShowMe = mstrName
    Else
        ShowMe = ""
    End If
End Function

Private Sub cmdDelete_Click()
          Dim strSQL As String
          
1         On Error GoTo cmdDelete_Click_Error

2         If vsgMain.Row < 1 Or vsgMain.Row = vsgMain.Rows - 1 Then Exit Sub
          
3         If mintType = 1 Then
4             strSQL = "zl_������ҩ����_Update(1,'" & vsgMain.TextMatrix(vsgMain.Row, col����) & "')"
5         Else
6             strSQL = "zl_��������_Insert(Null,Null,'" & gUserInfo.Name & "','" & vsgMain.TextMatrix(vsgMain.Row, col����) & "')"
7         End If

8         ComExecuteProc Sel_His_DB, strSQL, Me.Caption
9         vsgMain.RemoveItem vsgMain.Row


10        Exit Sub
cmdDelete_Click_Error:
11        Call WriteErrLog("zlPublicHisCommLis", "frmKssReasonSelect", "ִ��(cmdDelete_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
12        Err.Clear

End Sub

Private Sub cmdOK_Click()
    Call vsgMain_DblClick
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
     vsgMain.SetFocus
End Sub

Private Sub Form_Load()
          Dim strTmp As String, strSQL As String
          Dim rsTmp As Recordset, i As Long
          
1         On Error GoTo Form_Load_Error

2         mstrName = ""
3         mblnOk = False
4         If mstrFind <> "" Then
5             If IsNumeric(mstrFind) Then
6                 strTmp = " Where (����=LPAD([1]," & IIf(mintType = 1, "4", "5") & ",'0') Or ���� Like [2]) "
7             Else
8                 strTmp = " Where (���� Like [2] Or ���� Like [2]) "
9             End If
10        End If
11        If mintType = 1 Then
12            strSQL = "Select ����,����,���� From ������ҩ����" & strTmp & " order by to_number(����)"
13        Else
14            strSQL = "Select ����,����,���� From ��������" & strTmp & IIf(strTmp = "", " Where ", " And ") & " (��Ա=[3] or ��Ա is null) order by to_number(����)"
15        End If
       
16        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, mstrFind, "%" & UCase(mstrFind) & "%", gUserInfo.Name)
          
17        vsgMain.Rows = 1: vsgMain.AddItem ""
18        Me.Left = mlngleft
19        Me.Top = mlngTop
20        If Not rsTmp.EOF Then
21            If rsTmp.RecordCount = 1 Then
                  'ֻ��һ����¼ֱ�ӷ���
22                mblnOk = True
23                mstrName = rsTmp!���� & ""
24                Unload Me
25            Else
26                With vsgMain
27                    For i = 1 To rsTmp.RecordCount
28                        .TextMatrix(i, col����) = NVL(rsTmp!����)
29                        .TextMatrix(i, col����) = NVL(rsTmp!����)
30                        .TextMatrix(i, col����) = NVL(rsTmp!����)
31                        rsTmp.MoveNext
32                        .AddItem ""
33                    Next
34                    vsgMain.Cell(flexcpBackColor, vsgMain.Rows - 1, col����) = &HFFEADA
35                    vsgMain.Row = 1
36                End With
37            End If
38        Else
39            Unload Me
40            mblnOk = True
41        End If


42        Exit Sub
Form_Load_Error:
43        Call WriteErrLog("zlPublicHisCommLis", "frmKssReasonSelect", "ִ��(Form_Load)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
44        Err.Clear

End Sub


Private Sub vsgMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = vsgMain.Rows - 1 And NewCol = col���� Then
        vsgMain.FocusRect = flexFocusHeavy
        vsgMain.Editable = flexEDKbdMouse
    Else
        vsgMain.FocusRect = flexFocusNone
        vsgMain.Editable = flexEDNone
    End If
End Sub

Private Sub vsgMain_DblClick()
    If vsgMain.Row < 1 Or vsgMain.Row = vsgMain.Rows - 1 Then Exit Sub
    mblnOk = True
    mstrName = vsgMain.TextMatrix(vsgMain.Row, col����)
    Unload Me
End Sub

Private Sub vsgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call vsgMain_DblClick
End Sub

Private Sub vsgMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
          Dim strSQL As String, rsTmp As Recordset
          Dim strSpellCode As String
          
1         On Error GoTo vsgMain_ValidateEdit_Error

2         If Row = vsgMain.Rows - 1 And Col = col���� Then
3             If vsgMain.EditText = "" Then Exit Sub
4             If mintType = 1 Then
5                 If ActualLen(vsgMain.EditText) > 1000 Then
6                     MsgBox "�������ݲ������� 500 �����ֻ� 1000 ���ַ���", vbInformation, gSysInfo.ShortName
7                     Cancel = True: Exit Sub
8                 End If
9                 strSQL = "Select 1 From ������ҩ���� Where ����=[1]"
10                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, vsgMain.EditText)
                  '����Ѿ����ˣ���ʾ�û��Ƿ������
11                If rsTmp.RecordCount > 0 Then
12                    MsgBox "�Ѿ�������ͬ����ҩ���ɡ�", vbInformation, Me.Caption
13                    Cancel = True: Exit Sub
14                End If
15                strSQL = "Select LPad(To_Char(Max(To_Number(����)) + 1), 4, '0') as ���� From ������ҩ����"
16                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
17                If rsTmp.RecordCount < 1 Then Exit Sub
18                strSpellCode = Mid(SpellCode(vsgMain.EditText), 1, 10)
19                strSQL = "zl_������ҩ����_Update(0,'" & rsTmp!���� & "" & "','" & vsgMain.EditText & "','" & strSpellCode & "')"
20                ComExecuteProc Sel_His_DB, strSQL, Me.Caption
                  
21            Else
22                If ActualLen(vsgMain.EditText) > 100 Then
23                    MsgBox "�������ݲ������� 50 �����ֻ� 100 ���ַ���", vbInformation, gSysInfo.ShortName
24                    Cancel = True: Exit Sub
25                End If
26                strSQL = "Select 1 From �������� Where ����=[1] And (��Ա=[2] Or ��Ա is null)"
27                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, Replace(vsgMain.EditText, "'", "''"), gUserInfo.Name)
28                If rsTmp.RecordCount > 0 Then
29                    MsgBox "�����������Ѿ��ڳ��������С�", vbInformation, Me.Caption
30                    Cancel = True: Exit Sub
31                    Exit Sub
32                End If
                  
                  
33                strSpellCode = zlGetSymbol(vsgMain.EditText, CByte(0))
34                strSQL = "zl_��������_Insert('" & Replace(vsgMain.EditText, "'", "''") & "','" & strSpellCode & "','" & gUserInfo.Name & "')"
35                Call ComExecuteProc(Sel_His_DB, strSQL, Me.Caption)
                  '���ϱ���
36                strSQL = "Select ���� From �������� Where ����=[1] And (��Ա=[2] Or ��Ա is null)"
37                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, Replace(vsgMain.EditText, "'", "''"), gUserInfo.Name)
38            End If
39            vsgMain.Editable = flexEDNone
40            If rsTmp.RecordCount > 0 Then
41                vsgMain.TextMatrix(Row, col����) = rsTmp!����
42                vsgMain.TextMatrix(Row, col����) = strSpellCode
43            End If
44            vsgMain.Cell(flexcpBackColor, Row, col����) = &H80000005
45            vsgMain.AddItem ""
46            vsgMain.Cell(flexcpBackColor, vsgMain.Rows - 1, col����) = &HFFEADA
47        End If


48        Exit Sub
vsgMain_ValidateEdit_Error:
49        Call WriteErrLog("zlPublicHisCommLis", "frmKssReasonSelect", "ִ��(vsgMain_ValidateEdit)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
50        Err.Clear

End Sub
