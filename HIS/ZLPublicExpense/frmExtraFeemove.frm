VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmExtraFeemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҽ������ת��"
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6228
   Icon            =   "frmExtraFeemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExtraFeemove.frx":058A
   ScaleHeight     =   3504
   ScaleWidth      =   6228
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   760
      Left            =   0
      ScaleHeight     =   756
      ScaleWidth      =   6228
      TabIndex        =   4
      Top             =   0
      Width           =   6225
      Begin VB.Image imgInfo 
         Height          =   576
         Left            =   120
         Picture         =   "frmExtraFeemove.frx":0B14
         Top             =   0
         Width           =   576
      End
      Begin VB.Label lblNote 
         BackColor       =   &H80000005&
         Caption         =   "��ǰѡ��ķ��õ���:XXXXYYYY,��ת�ƹ���������ָ����ҽ����,��ѡ��һ�д�������ҽ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   612
      ScaleWidth      =   6228
      TabIndex        =   0
      Top             =   2892
      Width           =   6225
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   3950
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5070
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Bindings        =   "frmExtraFeemove.frx":2656
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   795
      Width           =   6195
      _cx             =   10927
      _cy             =   3625
      Appearance      =   2
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExtraFeemove.frx":266A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
Attribute VB_Name = "frmExtraFeemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngҽ��ID As Long      '���ѹ�����ҽ��ID(��ID,���ѿ����ǹ������Ը�IDΪ���ID�Ĳ�λ�򷽷���ҽ��ID�ϵ�)
Private mstrNO As String        '���ѵ��ݺ�
Private mint��¼���� As Integer '���Ѽ�¼����
Private mint�������� As Integer '1=�������(�����������)��2-סԺ����
Private mblnOK As Boolean

Private Const col_NO = 0
Private Const col_ҽ������ = 1
Private Const col_����ʱ�� = 2


Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳId As Long, ByVal str�Һŵ� As String, _
    ByVal lngҽ��ID As Long, ByVal str������� As String, ByVal lngִ�в���id As Long, _
    ByVal strNO As String, ByVal int��¼���� As Integer, ByVal int�������� As Integer)
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    '1.���ιҺŻ򱾴�סԺ��ͬһ���˵�ҽ��
    '2.�뵱ǰ���ѹ���ҽ����ͬ����ִ�п��ҵ�ҽ��
    '3.�����ǰ�ǹ�������鷽����λ�ϵģ�ת�ƺ��������ҽ��������¼��
    strSQL = "Select b.No, b.ҽ��id, b.���ͺ�, To_Char(b.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��, a.ҽ������" & vbNewLine & _
            "From ����ҽ����¼ A, ����ҽ������ B" & vbNewLine & _
            "Where a.Id = b.ҽ��id And a.���id Is Null And a.������� = [5] And b.ִ�в���id = [6]" & vbNewLine & _
            "      And a.id <> [4] And a.����id = [1]" & _
            IIf(str�Һŵ� <> "", " And a.�Һŵ� = [2]", " And a.��ҳID = [3]") & " Order by NO"
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "����ת��", lng����ID, str�Һŵ�, lng��ҳId, lngҽ��ID, str�������, lngִ�в���id)
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò����ڱ�����û����ͬ����ҽ�������ܽ��и���ת�ơ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mlngҽ��ID = lngҽ��ID
    mstrNO = strNO
    mint��¼���� = int��¼����
    mint�������� = int��������
    
    
    lblNote.Caption = Replace(lblNote.Caption, "XXXXYYYY", strNO)
    Call LoadList(rsTmp)
    
    mblnOK = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    
    ShowMe = mblnOK
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Sub LoadList(ByRef rsList As ADODB.Recordset)
'���ܣ�����ҽ�������嵥
    Dim i As Long
    
    With vsList
        .Rows = .FixedRows
        .Rows = .FixedRows + rsList.RecordCount
        
        For i = 1 To rsList.RecordCount
            .TextMatrix(i, col_NO) = rsList!NO
                        
            .TextMatrix(i, col_ҽ������) = rsList!ҽ������
            .Cell(flexcpData, i, col_ҽ������) = Val(rsList!ҽ��ID)
            
            .TextMatrix(i, col_����ʱ��) = rsList!����ʱ��
            .Cell(flexcpData, i, col_����ʱ��) = Val(rsList!���ͺ�)
            
            rsList.MoveNext
        Next
        If .Rows = .FixedRows + 1 Then
            .Row = .Rows - 1
        Else
            .Row = 0 'ȱʡ��ѡ���κ�һ��
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strSQL As String
    With vsList
        If .Row <= .FixedRows - 1 Then
            MsgBox "��ѡ��һ�д�������ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷ��Ҫ��ҽ������" & mstrNO & "������" & vbCrLf & "ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """(" & _
                    .TextMatrix(.Row, col_NO) & ")����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        lngҽ��ID = Val(.Cell(flexcpData, .Row, col_ҽ������))
        lng���ͺ� = Val(.Cell(flexcpData, .Row, col_����ʱ��))
        
        strSQL = "Zl_����ҽ������_Move(" & mint��¼���� & ",'" & mstrNO & "'," & mint�������� & "," & lngҽ��ID & "," & lng���ͺ� & ")"
        Call gobjDatabase.ExecuteProcedure(strSQL, "����ת��")
    End With
    
    mblnOK = True
    Unload Me
End Sub
