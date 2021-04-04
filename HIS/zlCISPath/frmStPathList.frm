VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStPathList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��׼·��ѡ��"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   Icon            =   "frmStPathList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8580
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   8580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3555
      Width           =   8580
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6960
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5520
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   240
         X2              =   10240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsStPathList 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8580
      _cx             =   15134
      _cy             =   6324
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   360
      RowHeightMax    =   360
      ColWidthMin     =   200
      ColWidthMax     =   5000
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStPathList.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmStPathList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent     As Object '������
Private mlStPathID     As Long 'ѡ��ı�׼·��ID
Private mblnOK         As Boolean
Private mrsStPath      As ADODB.Recordset '������ϻ�ȡ�Ĳ��˿��õı�׼·���б�

Private Enum Cols
    COL���� = 0
    COL���� = 1
    COL���� = 2
    COL�汾˵�� = 3
End Enum

Public Function ShowMe(frmParent As Object, ByVal str�������� As String, Optional ByVal intMode As Integer) As Boolean
'���ܣ����ݼ��������ȡ��ر�׼·��
'������frmParent  ������
'       str�������� :��ʽ����������1,��������2,��������3...
'       ˵�����������ṩ���ֲο���׼·����ģʽ
'       1���Լ����������ο���ر�׼·�� str��������<>""
'       2���ο����б�׼·�� str��������=""
'       intMode 0-סԺ��1-����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, str���� As String, strSub���� As String
    Dim i As Long
    Dim arrtmp As Variant, strSub���� As String '���з���ļ���������н�ȡ
    Dim strTables As String
    
    On Error GoTo errH:
    Set mfrmParent = frmParent
    If intMode = 1 Then         '����
        strTables = "��׼����·��Ŀ¼ A, ��׼����·������ B"
    Else                        'סԺ
        strTables = "��׼·��Ŀ¼ A, ��׼·������ B"
    End If
    
    If Trim(str��������) = "" Then
        Call frmStandardPathRef.ShowMe(mfrmParent, 0, , intMode)
        Exit Function
    Else
        arrtmp = Split(str��������, ",")
        For i = LBound(arrtmp) To UBound(arrtmp)
            If i <> LBound(arrtmp) Then
                If InStr(Trim(arrtmp(i)), ".") > 0 Then
                    strSub���� = strSub���� & " Or InStr(b.��������,'" & Left(Trim(arrtmp(i)), InStr(Trim(arrtmp(i)), ".") - 1) & "') > 0"
                Else
                    strSub���� = strSub���� & " Or InStr(b.��������,'" & Trim(arrtmp(i)) & "') > 0"
                End If
                str���� = str���� & " Or InStr(b.��������,'" & Trim(arrtmp(i)) & "') > 0"
            Else
                If InStr(Trim(arrtmp(i)), ".") > 0 Then
                    strSub���� = " ( InStr(b.��������,'" & Left(Trim(arrtmp(i)), InStr(Trim(arrtmp(i)), ".") - 1) & "') > 0"
                Else
                    strSub���� = " ( InStr(b.��������,'" & Trim(arrtmp(i)) & "') > 0"
                End If
                str���� = " ( InStr(b.��������,'" & Trim(arrtmp(i)) & "') > 0"
            End If
        Next
        str���� = str���� & " )"
        strSub���� = strSub���� & " )"
        strSql = " Select a.Id, a.��������, a.����, a.·������, a.�汾˵��, b.��������" & vbNewLine & _
                 " From " & strTables & vbNewLine & _
                 " Where  a.Id = b.��׼·��id  And " & str����
        Set mrsStPath = zlDatabase.OpenSQLRecord(strSql, gstrSysName)
    End If
    
    If mrsStPath.RecordCount = 0 Then 'û�з�����Ҫ��ϵı�׼·��,�ͽ�ȡ����������࣬��Ҫ�ӷ������ƥ�����
        strSql = " Select a.Id, a.��������, a.����, a.·������, a.�汾˵��, b.��������" & vbNewLine & _
                 " From " & strTables & vbNewLine & _
                 " Where  a.Id = b.��׼·��id  And " & strSub����
        Set mrsStPath = zlDatabase.OpenSQLRecord(strSql, gstrSysName)
        If mrsStPath.RecordCount = 0 Then 'û�з�����Ҫ��ϵı�׼·��
            Call frmStandardPathRef.ShowMe(mfrmParent, 0, , intMode)
            Exit Function
        ElseIf mrsStPath.RecordCount = 1 Then '����һ����׼·��������Ҫ���
            Call frmStandardPathRef.ShowMe(mfrmParent, mrsStPath!ID, , intMode)
            Exit Function
        Else '���ж�����׼·��������Ҫ���
            Me.Show 1, frmParent
            Call frmStandardPathRef.ShowMe(mfrmParent, mlStPathID, , intMode)
            ShowMe = mblnOK
            Exit Function
        End If
    ElseIf mrsStPath.RecordCount = 1 Then '����һ����׼·��������Ҫ���
        Call frmStandardPathRef.ShowMe(mfrmParent, mrsStPath!ID, , intMode)
        Exit Function
    Else '���ж�����׼·��������Ҫ���
        Me.Show 1, frmParent
        Call frmStandardPathRef.ShowMe(mfrmParent, mlStPathID, , intMode)
        ShowMe = mblnOK
        Exit Function
    End If
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
'���ܣ�ȡ��ʱ�Ƴ�����
    mblnOK = False
    mlStPathID = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()
'���ܣ�ȷ��ʱ��ȡѡ�б�׼·��
    If mlStPathID = 0 Then
        MsgBox "�㻹δѡ���׼·������ѡ���׼·��", vbOKOnly, gstrSysName
    Else
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
'���ܣ���׼·�����ݼ��ص��ؼ���
    Dim i As Long
    With vsStPathList
        .Rows = .FixedRows
        mrsStPath.MoveFirst
        For i = 1 To mrsStPath.RecordCount
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL����) = mrsStPath!��������
            .TextMatrix(.Rows - 1, COL����) = mrsStPath!���� & ""
            .TextMatrix(.Rows - 1, COL����) = mrsStPath!·������ & ""
            .TextMatrix(.Rows - 1, COL�汾˵��) = mrsStPath!�汾˵�� & ""
            .RowData(.Rows - 1) = mrsStPath!ID & ""
            mrsStPath.MoveNext
        Next
    End With
End Sub

Private Sub Form_Resize()
'���ܣ����������пؼ�λ��
    vsStPathList.Width = Me.ScaleWidth - vsStPathList.Left
    picBottom.Top = Me.ScaleHeight - picBottom.Height
    vsStPathList.Height = picBottom.Top - vsStPathList.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
'���ܣ�ȷ���˳�ʱ�����ұ�׼·���Ѿ�ѡ��ʱ����ֹ�˳���������ֹ�˳�
    If mblnOK And mlStPathID = 0 Then
        Cancel = True
    Else
        Set mrsStPath = Nothing
        Set mfrmParent = Nothing
    End If
End Sub

Private Sub vsStPathList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'���ܣ�ѡ���׼·��
    mlStPathID = Val(vsStPathList.RowData(NewRow))
End Sub

Private Sub vsStPathList_DblClick()
'���ܣ�˫��ĳ����׼·��ʱ��ȷ��ѡ��
    mlStPathID = Val(vsStPathList.RowData(vsStPathList.Row))
    Call cmdOK_Click
End Sub


