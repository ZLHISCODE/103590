VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ٴ�·��ѡ��"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
   Icon            =   "frmPathImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6690
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6690
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Width           =   6690
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5520
         TabIndex        =   5
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4200
         TabIndex        =   4
         Top             =   195
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
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
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   6495
      _cx             =   11456
      _cy             =   2090
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathImport.frx":169B2
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
   Begin VSFlex8Ctl.VSFlexGrid vsDiag 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6495
      _cx             =   11456
      _cy             =   2355
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   320
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathImport.frx":16A4A
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
   Begin VB.Label lblDiag 
      Caption         =   "�ò����ж���ϲ�֢�򲢷�֢����ѡ��һ����"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   400
      Width           =   4575
   End
   Begin VB.Label lblPait 
      Caption         =   "��ǰ���ˣ���С��,��,46��"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblPath 
      Caption         =   "����±���ѡ��һ�������ڸò��˵��ٴ�·��"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6495
   End
End
Attribute VB_Name = "frmPathImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun As Long '0-��Ҫ·������,1-��ת·��(�׶�����ʱ)��=2  �ϲ�·������,=3�ϲ�·��ȡ�����룬=4�ϲ�·�����
Private mlngPathID As Long  '·����תʱ������ѡ���·����ID
Private mlngPathVersion As Long
Private mlngCurPathID As Long '·����תʱ����ǰ·��ID
Private mlngDiagnosisType As Long '�������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
Private mlngDiagnosisSorce As Long '�����Դ1-������2-��Ժ�Ǽǣ�3-��ҳ����;4-����

Private mPati As TYPE_Pati
Private mrsPath As ADODB.Recordset
Private mrsPati As ADODB.Recordset
Private mrsMerge As ADODB.Recordset    '�˶����Ǵ�����ڴ��ַ�Ͷ�����Unload�в���ж�ء�
Private mfrmParent As Object
Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng���ID As Long
Private mrsDiag As ADODB.Recordset
Private mlng��Ҫ·����¼ID As Long
Private mblnTmp As Boolean
Private mblnChoose As Boolean
Private mblnPathSend As Boolean
Private mbln��� As Boolean
Private mt_pp As TYPE_PATH_Pati
Private mlngHwnd As Long

Public Function ShowMe(frmParent As Object, t_pati As TYPE_Pati, ByVal lngFun As Long, ByRef t_pp As TYPE_PATH_Pati, _
    Optional ByVal lngCurPathID As Long, Optional ByRef lngPathID As Long, Optional ByRef lngPathVersion As Long, _
     Optional ByVal blnAuto As Boolean, Optional lng��Ҫ·����¼ID As Long, Optional rsMerge As Recordset, _
    Optional ByVal blnChoose As Boolean, Optional ByVal lngHwnd As Long, _
    Optional ByRef str���� As String, Optional ByRef lngDiagnosisType As Long, Optional ByRef lngDiagnosisSorce As Long, _
    Optional ByRef lng����ID As Long, Optional ByRef lng���ID As Long) As Boolean
'������
'       lngCurPathID As Long '·����תʱ����ǰ·��ID
'       lngPathID,lngPathVersion=·����תʱ������֮ǰѡ���·��ID�Ͱ汾������ѡ���·��ID�Ͱ汾
'       lngFun=0  ��Ҫ��ϵ��룬=1  �ϲ�·������,=2�ϲ�·��ȡ�����룬=3�ϲ�·�����, =4�ϲ�·�����
'       lngFun=1ʱ��blnAuto=true ������Ҫ·�����Զ����ü���Ƿ��пɵ���ĺϲ�·����û��ֱ���˳�������ʾ
'       rsMerge=�ϲ�·��ȡ�������ʱ������ж���ϲ�·�����򵯳�ѡ���rsmerge��Ϊ���кϲ�·���ļ�¼,blnChoose=true ,��ѡ
'       lngHwnd=�°没�����븸��������Ĭ��Ϊ0,�°没������ʾ���벻�ɹ���ԭ��
    Dim str������� As String
    Dim rsTmp As ADODB.Recordset, rsNext As ADODB.Recordset
    Dim str����IDs As String
    Dim str���IDs As String
    Dim blnTmp As Boolean
    Dim bln��ҽ As Boolean
    Dim i As Long
    
    mPati = t_pati
    mlng��Ҫ·����¼ID = lng��Ҫ·����¼ID
    Set mfrmParent = frmParent
    mlngFun = lngFun
    
    mblnOK = False
    mlngCurPathID = lngCurPathID
    mlngPathID = lngPathID
    mlngPathVersion = lngPathVersion
    Set mrsMerge = rsMerge
    mblnChoose = blnChoose
    mlngHwnd = lngHwnd
    
    mlngDiagnosisType = 0
    mlngDiagnosisSorce = 0
    mblnPathSend = CheckPathSend(mPati.����ID, mPati.��ҳID)
    If lngHwnd <> 0 Then blnAuto = True
    
    '����·��
    If rsMerge Is Nothing And lngFun <> 3 And lngFun <> 4 Then
        Set rsTmp = Get����ID(mPati.����ID, mPati.��ҳID, IIf(lngFun = 2, 1, IIf(mblnPathSend And lngCurPathID = 0, 2, 0)), mPati.����ID, bln��ҽ)
        If rsTmp.RecordCount > 0 Then
            mlng����ID = Val("" & rsTmp!����id)
            mlng���ID = Val("" & rsTmp!���id)
            str������� = "" & rsTmp!�������
            mlngDiagnosisType = Val("" & rsTmp!�������)
            mlngDiagnosisSorce = Val("" & rsTmp!��¼��Դ)
        End If
        If mlng����ID = 0 And mlng���ID = 0 Then
            If Not blnAuto Then
                If mlngFun = 0 Then
                    MsgBox "�ò���û����д�κ���ϣ�������д����ִ�е��롣", vbInformation, gstrSysName
                ElseIf mlngFun = 2 Then
                    MsgBox "�ò��˳���������⣬��δ��д������ϻ򲢷�֢��������д���ڵ���ϲ�·����", vbInformation, gstrSysName
                ElseIf mlngFun = 1 Then
                    MsgBox "�ò��˵���ϼ�¼��ɾ�����޷�ִ��·����ת��", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        End If
        
        If mlngFun = 0 Or mlngFun = 1 Then
            '��һ�ε�����Ч·��ʱ����ת·��ʱ������һ���ֲ�ѯ
            If mblnPathSend = False Or mlngFun = 1 Then
                '�������Ҫ·������ݵ�һ������ID�ж�
                If mblnPathSend = False And bln��ҽ Then
                    '����ҽ������Ժ���û��ƥ�䵽��ҽ·��ʱĬ���ж���ҽ��Ժ���
                    rsTmp.Filter = "������� =12 OR ������� =2 "
                    For i = 1 To rsTmp.RecordCount
                        mlng����ID = Val("" & rsTmp!����id)
                        mlng���ID = Val("" & rsTmp!���id)
                        str������� = "" & rsTmp!�������
                        mlngDiagnosisType = Val("" & rsTmp!�������)
                        mlngDiagnosisSorce = Val("" & rsTmp!��¼��Դ)
                        
                        Set mrsPath = GetPathTable(mlng����ID, mlng���ID, mPati.����ID, lngCurPathID)
                        If mrsPath.RecordCount > 0 Then Exit For
                        rsTmp.MoveNext
                    Next
                    If mrsPath Is Nothing Then
                        If lngHwnd = 0 Then MsgBox "��ǰ����û���ʺ��ڸò�����Ҫ���[" & str������� & "]���ٴ�·����", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                     Set mrsPath = GetPathTable(mlng����ID, mlng���ID, mPati.����ID, lngCurPathID)
                End If
                
                If mblnPathSend = False And mrsPath.RecordCount = 0 Then
                    Set rsNext = Get����ID(mPati.����ID, mPati.��ҳID, 3, mPati.����ID)
                    If rsNext.RecordCount > 0 Then
                        If MsgBox("��ǰ����û���ʺ��ڸò�����Ҫ���[" & str������� & "]���ٴ�·����" & vbCrLf & _
                                "�������ʺ��ڸò��˴�Ҫ���[" & rsNext!������� & "]���ٴ�·����" & vbCrLf & _
                                "�Ƿ����Ҫ��ϵ��ٴ�·����", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                            mlng����ID = Val("" & rsNext!����id)
                            mlng���ID = Val("" & rsNext!���id)
                            str������� = "" & rsNext!�������
                            mlngDiagnosisType = Val("" & rsNext!�������)
                            mlngDiagnosisSorce = Val("" & rsNext!��¼��Դ)
                            Set mrsPath = GetPathTable(mlng����ID, mlng���ID, mPati.����ID, lngCurPathID)
                        Else
                            Exit Function
                        End If
                    End If
                End If
            Else
                rsTmp.Filter = ""
                Do While Not rsTmp.EOF
                    If Val(rsTmp!����id & "") <> 0 Then
                        str����IDs = str����IDs & "," & rsTmp!����id
                    End If
                    If Val(rsTmp!���id & "") <> 0 Then
                        str���IDs = str���IDs & "," & rsTmp!���id
                    End If
                    rsTmp.MoveNext
                Loop
                If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
                Set mrsDiag = rsTmp
                str����IDs = Mid(str����IDs, 2)
                str���IDs = Mid(str���IDs, 2)
                Set mrsPath = GetPathTable(0, 0, mPati.����ID, 0, str����IDs, 0, str���IDs, mPati.����ID, mPati.��ҳID)
            End If
        Else
            '����Ǻϲ�·�����������зǵ��벡�ֵ�������ϻ򲢷�֢
            Do While Not rsTmp.EOF
                If Val(rsTmp!����id & "") <> 0 Then
                    str����IDs = str����IDs & "," & rsTmp!����id
                End If
                If Val(rsTmp!���id & "") <> 0 Then
                    str���IDs = str���IDs & "," & rsTmp!���id
                End If
                rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Set mrsDiag = rsTmp
            str����IDs = Mid(str����IDs, 2)
            str���IDs = Mid(str���IDs, 2)
            Set mrsPath = GetPathTable(0, 0, mPati.����ID, 0, str����IDs, mlng��Ҫ·����¼ID, str���IDs)
        End If
        
        If mrsPath.RecordCount = 0 Then
            If Not blnAuto Then
                If mlngFun = 0 Then
                    MsgBox "��ǰ����û��" & IIf(mlngFun = 0, "", "����") & "�ʺ��ڸò�����Ҫ���[" & str������� & "]���ٴ�·����", vbInformation, gstrSysName
                Else
                    MsgBox "��ǰ����û���ʺ��ڸò��˵��ٴ��ϲ�·����", vbInformation, gstrSysName
                End If
            End If
            Exit Function
        Else
            Set mrsPati = GetPatiInfo(mPati.����ID, mPati.��ҳID)
            If mrsPati.RecordCount = 0 Then
                MsgBox "��ȡ���˵�ǰסԺ��Ϣʧ�ܡ�", vbInformation, gstrSysName
                Exit Function
            End If
            If mlngFun = 2 And blnAuto And mrsPath.RecordCount = 1 Then
                If MsgBox("��ǰ���˴��ڿɵ���ĺϲ�·��:""" & mrsPath!���� & """���Ƿ�Ҫ���룿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    Else
        Set mrsPati = GetPatiInfo(mPati.����ID, mPati.��ҳID)
        If mrsPati.RecordCount = 0 Then
            MsgBox "��ȡ���˵�ǰסԺ��Ϣʧ�ܡ�", vbInformation, gstrSysName
            Exit Function
        End If
        'ֻ��δִ�й���
        If mblnChoose Then
            If mlngFun = 3 Then
                rsMerge.Filter = "�Ƿ�ִ��<>1"
                If rsMerge.RecordCount = 0 Then
                    MsgBox "�ò����кϲ�·�����Ѿ���������Ŀ����ȡ�����ɺϲ�·������Ŀ����ȡ�����롣", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf mlngFun = 4 Then
                blnTmp = False
                Do While Not rsMerge.EOF
                    If Val(mrsMerge!��ʾ & "") = 1 Then blnTmp = True
                    
                    rsMerge.MoveNext
                Loop
                If Not blnTmp Then
                    MsgBox "�ò���û�дﵽ��׼סԺ�յĺϲ�·����������ǰ��ɣ���ѡ����һ�׶���ǰ��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    On Error Resume Next
    If lngHwnd <> 0 Then
        Me.Show 1
    Else
        Me.Show 1, frmParent
    End If
    
    lngPathID = mlngPathID
    lngPathVersion = mlngPathVersion
    t_pp.����·��ID = mt_pp.����·��ID
    t_pp.����·��״̬ = mt_pp.����·��״̬
    t_pp.��ǰ�׶�ID = mt_pp.��ǰ�׶�ID
    t_pp.��ǰ�׶η�֧ID = mt_pp.��ǰ�׶η�֧ID
    t_pp.��ǰ���� = mt_pp.��ǰ����
    t_pp.��ǰ���� = mt_pp.��ǰ����
    t_pp.�ϲ�·������ = mt_pp.�ϲ�·������
    t_pp.�׶θ�ID = mt_pp.�׶θ�ID
    t_pp.����·������ = mt_pp.����·������
    t_pp.·��ID = mt_pp.·��ID
    t_pp.δ����ԭ�� = mt_pp.δ����ԭ��
    t_pp.ԭ·��ID = mt_pp.ԭ·��ID
    t_pp.�汾�� = mt_pp.�汾��
    lngDiagnosisType = mlngDiagnosisType
    lngDiagnosisSorce = mlngDiagnosisSorce
    lng����ID = mlng����ID
    lng���ID = mlng���ID
    
    ShowMe = mblnOK
    If mblnOK And Not rsMerge Is Nothing And (mlngFun = 3 Or mlngFun = 4) Then
         Set rsMerge = mrsMerge
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Set�������(ByVal lng·��ID As Long)
'���ܣ�����ǲ��˵���סԺ�ڶ��ε���·������ҪĬ����һ��������ϣ�����ж�����Ӧ����Ĭ��˳������Ժ����Ժ��
    Dim strSql As String, rsTmp As Recordset
    Dim str����IDs As String, str���IDs As String
    
    On Error GoTo errH
    strSql = "Select a.����ID,a.���id From �ٴ�·������ A Where a.·��ID=[1] And a.����=0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Set�������", lng·��ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If Val(rsTmp!����id & "") <> 0 Then
                str����IDs = str����IDs & "," & rsTmp!����id
            End If
            If Val(rsTmp!���id & "") <> 0 Then
                str���IDs = str���IDs & "," & rsTmp!���id
            End If
            rsTmp.MoveNext
        Loop
        If mrsDiag.RecordCount > 0 Then
            str����IDs = Mid(str����IDs, 2)
            str���IDs = Mid(str���IDs, 2)
            mrsDiag.MoveFirst
            Do While Not mrsDiag.EOF
                If InStr("," & str����IDs & ",", "," & mrsDiag!����id & ",") > 0 And Val(mrsDiag!����id & "") <> 0 _
                    Or InStr("," & str���IDs & ",", "," & mrsDiag!���id & ",") > 0 And Val(mrsDiag!���id & "") <> 0 Then
                    mlng����ID = Val("" & mrsDiag!����id)
                    mlng���ID = Val("" & mrsDiag!���id)
                    mlngDiagnosisType = Val("" & mrsDiag!�������)
                    mlngDiagnosisSorce = Val("" & mrsDiag!��¼��Դ)
                    mrsDiag.MoveFirst
                    Exit Sub
                End If
                mrsDiag.MoveNext
            Loop
            mrsDiag.MoveFirst
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim t_pp As TYPE_PATH_Pati, arrtmp As Variant
    Dim lngסԺ���� As Long, lng��׼סԺ�� As Long
    Dim rsTmp As ADODB.Recordset, strδ������� As String, strδ�������� As String
    Dim lngB As Long, lngE As Long, strUnit As String, strTmp As String, DatCur As Date, lngValue As Long
    Dim i As Long, strFilter As String
    Dim bln����ж� As Boolean
    Dim dt��Ժʱ�� As Date
    Dim dtDate As Date
    
    If mrsMerge Is Nothing And mlngFun <> 3 And mlngFun <> 4 Then
        If vsPath.Row <= 0 Then
            MsgBox "��ѡ��һ�������ڸò��˵��ٴ�·��.", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        arrtmp = Split(vsPath.RowData(vsPath.Row), ":")
        t_pp.·��ID = arrtmp(0)
        t_pp.�汾�� = arrtmp(1)
        mrsPath.Filter = "ID=" & t_pp.·��ID
        If mlngFun = 0 And mblnPathSend Then
            Call Set�������(t_pp.·��ID)
        End If

        Set rsTmp = GetUnImportReson


        '���˲������ͣ�����·����Ҳ������Ҫ��
        If Not IsNull(mrsPati!��������) Then
            If mrsPath!�������� <> "��" And mrsPati!�������� <> mrsPath!�������� Then
                MsgBox "��·��Ҫ��Ĳ�������[" & mrsPath!�������� & "]���ʺ��ڸò��˵Ĳ�������[" & mrsPati!�������� & "]", vbInformation, gstrSysName

                strδ�������� = "�������Ͳ�����"
                GoTo UnImport
            End If
        End If

        If Not IsNull(mrsPati!��ǰ����) Then
            If mrsPath!���ò��� <> "ͨ��" And mrsPath!���ò��� <> mrsPati!��ǰ���� Then
                MsgBox "��·��[" & mrsPath!���ò��� & "]���ʺ��ڸò��˲���[" & mrsPati!��ǰ���� & "]", vbInformation, gstrSysName

                strδ�������� = "���鲻����"
                GoTo UnImport
            End If
        End If
        If Val(mrsPath!�����Ա�) <> 0 Then
            If Val(mrsPath!�����Ա�) <> IIf(mrsPati!�Ա� = "��", 1, IIf(mrsPati!�Ա� = "Ů", 2, 0)) Then
                MsgBox "��·�����ʺ��ڸò����Ա�[" & mrsPati!�Ա� & "]", vbInformation, gstrSysName

                strδ�������� = "�Ա��ʺ�"
                GoTo UnImport
            End If
        End If

        If Not IsNull(mrsPath!��������) And Not IsNull(mrsPati!����) Then
            lngValue = 0
            lngB = Split(mrsPath!��������, "-")(0)
            strTmp = Split(mrsPath!��������, "-")(1)
            lngE = Mid(strTmp, 1, Len(strTmp) - 1)
            strUnit = Mid(strTmp, Len(strTmp))

            strTmp = mrsPati!����           '���⣺2��3�µ�
            If strUnit = Mid(strTmp, Len(strTmp)) And IsNumeric(Mid(strTmp, 1, Len(strTmp) - 1)) Or IsNumeric(strTmp) Then
                '��ͬ���䵥λ�����Ƚ�
                lngValue = Val(strTmp)
            ElseIf Not IsNull(mrsPati!��������) Then
                DatCur = zlDatabase.Currentdate
                lngValue = DateDiff(IIf(strUnit = "��", "yyyy", IIf(strUnit = "��", "m", "d")), CDate(mrsPati!��������), DatCur)
                If lngValue = 0 Then lngValue = 1
            End If
            If lngValue <> 0 Then
                If lngValue < lngB Or lngValue > lngE Then
                    MsgBox "��·�����ʺ��ڸò�������[" & mrsPati!���� & "]", vbInformation, gstrSysName

                    strδ�������� = "���䲻�ʺ�"
                    GoTo UnImport
                End If
            End If
        End If


        'סԺ�ղ��ܴ���·���ı�׼סԺ�պ�ȷ������(���û��������ȷ��������������)
        dt��Ժʱ�� = GetPatiInDate(mPati, lngסԺ����)
        dtDate = zlDatabase.Currentdate

        If InStr(mrsPath!��׼סԺ��, "-") > 0 Then
            lng��׼סԺ�� = Split(mrsPath!��׼סԺ��, "-")(1)
        Else
            lng��׼סԺ�� = Val(mrsPath!��׼סԺ��)
        End If
        '����Ǻϲ�·�������ǵ���סԺ���Ѿ����ɹ�·����Ŀ�ģ������סԺ����������׼סԺ��
        '104002:������ȷ������,סԺ��������ȷ��������ֹ����·��;ȷ������δ���û�Ϊ0ʱ,��סԺ�������ڱ�׼סԺ��ʱ��ֹ����·��
        If mlngFun = 0 Or mlngFun = 1 Then
            If Not CheckPathSend(mPati.����ID, mPati.��ҳID) Then
                If mrsPath!ȷ������ <> 0 Then
                    If dtDate > Format(DateAdd("d", Val(mrsPath!ȷ������), dt��Ժʱ��), "yyyy-MM-DD HH:mm:ss") Then
                        MsgBox "�ò�������Ժ" & lngסԺ���� & "�죬�����˹涨��ȷ������(" & mrsPath!ȷ������ & "��)��", vbInformation, gstrSysName
                        strδ�������� = "����ȷ������"
                        GoTo UnImport
                    End If
                Else
                    If lngסԺ���� > lng��׼סԺ�� Then
                        MsgBox "�ò�������Ժ" & lngסԺ���� & "�죬�����˸�·���ı�׼סԺ��(" & lng��׼סԺ�� & "��)��", vbInformation, gstrSysName
                        strδ�������� = "������׼סԺ��"
                        GoTo UnImport
                    End If
                End If
            End If
        End If

        If mlngFun = 0 Or mlngFun = 2 Then
            Me.Hide
            bln����ж� = True
            '�ٴ�·������ǰ������ҿ�
            If CreatePlugInOK(p�ٴ�·��Ӧ��) Then
                On Error Resume Next
                bln����ж� = gobjPlugIn.PathImportBefore(glngSys, p�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, t_pp.·��ID, t_pp.�汾��, , mlngDiagnosisSorce, mlng����ID, mlng���ID)
                '����ӿڲ����ڣ���Ӱ��ԭ���߼�
                If Not bln����ж� And Err.Number <> 0 Then bln����ж� = True
                Call zlPlugInErrH(Err, "PathImportBefore")
                Err.Clear: On Error GoTo 0
                If Not bln����ж� Then
                    mbln��� = True
                    mblnOK = True
                    Unload Me
                    Exit Sub
                End If
            End If
            
            If mlngHwnd = 0 Then
                mblnOK = frmEvaluate.ShowMe(mfrmParent, 0, 1, mPati, t_pp, mrsPath!����, mlngDiagnosisType, mlngDiagnosisSorce, mlng����ID, mlng���ID, IIf(mlngFun = 0, 0, 1), mlng��Ҫ·����¼ID)
            Else
                mblnOK = True
                mt_pp.����·��ID = t_pp.����·��ID
                mt_pp.����·��״̬ = t_pp.����·��״̬
                mt_pp.��ǰ�׶�ID = t_pp.��ǰ�׶�ID
                mt_pp.��ǰ�׶η�֧ID = t_pp.��ǰ�׶η�֧ID
                mt_pp.��ǰ���� = t_pp.��ǰ����
                mt_pp.��ǰ���� = t_pp.��ǰ����
                mt_pp.�ϲ�·������ = t_pp.�ϲ�·������
                mt_pp.�׶θ�ID = t_pp.�׶θ�ID
                mt_pp.����·������ = t_pp.����·������
                mt_pp.·��ID = t_pp.·��ID
                mt_pp.δ����ԭ�� = t_pp.δ����ԭ��
                mt_pp.ԭ·��ID = t_pp.ԭ·��ID
                mt_pp.�汾�� = t_pp.�汾��
            End If
            '�ٴ�·������ǰ������ҿ�
            If CreatePlugInOK(p�ٴ�·��Ӧ��) Then
                On Error Resume Next
                Call gobjPlugIn.PathImportAfter(glngSys, p�ٴ�·��Ӧ��, mPati.����ID, mPati.��ҳID, t_pp.·��ID, t_pp.�汾��)
                Call zlPlugInErrH(Err, "PathImportAfter")
                Err.Clear: On Error GoTo 0
            End If
            
            If cmdOK.Tag <> "Unload" Then Unload Me
        Else
            mlngPathID = t_pp.·��ID
            mlngPathVersion = t_pp.�汾��
            mblnOK = True
            Unload Me
        End If
    Else
        With vsPath
            On Error Resume Next
            If mblnChoose Then
                If mlngFun = 3 Then
                    For i = .FixedRows To .Rows - 1
                        If .Cell(flexcpChecked, i, 0) = 1 Then
                            strFilter = strFilter & " Or ID=" & .RowData(i)
                        End If
                    Next
                    If strFilter = "" Then
                        MsgBox "������ѡ��һ����Ҫȡ���ĺϲ�·����", vbInformation, gstrSysName
                        Exit Sub
                    Else
                        mrsMerge.Filter = Mid(strFilter, 5)
                        mblnOK = True
                        Me.Hide
                    End If
                Else
                    For i = .FixedRows To .Rows - 1
                        mrsMerge.Filter = "ID=" & .RowData(i)
                        If .Cell(flexcpChecked, i, 0) = 1 Then
                            mrsMerge!ѡ�� = 1
                        Else
                            mrsMerge!ѡ�� = 0
                        End If
                    Next
                    mrsMerge.Update
                    mrsMerge.Filter = 0
                    mrsMerge.MoveFirst
                    mblnOK = True
                    Unload Me
                End If
            Else
                If .Row > 0 Then
                    mrsMerge.Filter = "ID=" & .RowData(.Row)
                    mblnOK = True
                    Me.Hide
                Else
                    MsgBox "��ѡ��һ���ϲ�·����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End With
    End If
    Exit Sub

UnImport:
    '��Ҫ��ϲű���δ����ԭ��
    If mlngFun = 0 Then
        rsTmp.Filter = "����='" & strδ�������� & "'"
        If rsTmp.RecordCount = 0 Then
            strδ������� = ""
        Else
            strδ������� = rsTmp!����
        End If

        Call SaveUnImport(mPati, t_pp, strδ�������, strδ��������)
        mblnOK = True
        If cmdOK.Tag <> "Unload" Then
            Unload Me
        End If
    End If
End Sub

Private Sub SaveUnImport(mPati As TYPE_Pati, mPP As TYPE_PATH_Pati, strVariationCode As String, strVariationTitle As String)
'���ܣ�����δ����ԭ��
'������strVariationCode=δ����ԭ�����,strVariationTitle=δ����ԭ������
    Dim strSql As String, strID As String, DateInPath As Date
    Dim str���ϵ��� As String

    strID = zlDatabase.GetNextId("�����ٴ�·��")
    If CheckPathSend(mPati.����ID, mPati.��ҳID) Then
        DateInPath = zlDatabase.Currentdate
    Else
        DateInPath = GetPatiInDate(mPati)
    End If
    str���ϵ��� = "0"
    
    
    strSql = "Zl_����·������_Insert(" & mPati.����ID & "," & mPati.��ҳID & "," & mPati.����ID & "," & _
            mPP.·��ID & "," & mPP.�汾�� & "," & strID & ",'" & UserInfo.���� & "','" & strVariationTitle & "'," & _
            str���ϵ��� & ",To_Date('" & Format(DateInPath, "yyyy-MM-DD HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & _
            strVariationCode & "'," & mlngDiagnosisType & "," & mlngDiagnosisSorce & "," & IIf(mlng����ID = 0, "NULL", mlng����ID) & "," & IIf(mlng���ID = 0, "NULL", mlng���ID) & ",Null,1)"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSql, "δ����ԭ��")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetUnImportReson() As ADODB.Recordset
'���ܣ���ȡ�̶��ļ���δ����ԭ��
    Dim strSql As String
 
    strSql = "Select ����, ���� From ���쳣��ԭ�� Where ���� = 0 And ĩ�� = 1"
    On Error GoTo errH
    Set GetUnImportReson = zlDatabase.OpenSQLRecord(strSql, "δ����ԭ��")

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If cmdOK.Tag = "Unload" Then
        cmdOK.Tag = ""
        Unload Me
    Else
        If vsPath.Rows = vsPath.FixedRows + 1 Then vsPath.Row = vsPath.Rows - 1: vsPath.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strFilter As String
    Dim blnVisble As Boolean
    
    lblPait.Caption = "��ǰ���ˣ�" & mrsPati!���� & "," & mrsPati!�Ա� & "," & mrsPati!����
    
    If mrsMerge Is Nothing And mlngFun <> 3 And mlngFun <> 4 Then
        If mlngFun = 2 Then
            With vsDiag
                '�Ȳ����ж��ٸ��ϲ�֢�򲢷�֢��·����Ӧ
                Do While Not mrsPath.EOF
                    If Val(mrsPath!����id & "") <> 0 Then
                        If InStr(strFilter & "", " Or ����ID=" & mrsPath!����id) = 0 Then
                            strFilter = strFilter & " Or ����ID=" & mrsPath!����id
                        End If
                    ElseIf Val(mrsPath!���id & "") <> 0 Then
                        If InStr(strFilter & "", " Or ���ID=" & mrsPath!���id) = 0 Then
                            strFilter = strFilter & " Or ���ID=" & mrsPath!���id
                        End If
                    End If
                    mrsPath.MoveNext
                Loop
                strFilter = Mid(strFilter, 5)
                If strFilter <> "" Then
                    mrsDiag.Filter = strFilter
                End If
                
                If mrsDiag.RecordCount > 0 Then
                    If mrsDiag.RecordCount = 1 Then
                        blnVisble = False
                    Else
                        blnVisble = True
                    End If
                    .Rows = .FixedRows
                    Do While Not mrsDiag.EOF
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 0) = IIf(Val(mrsDiag!������� & "") = 10, "����֢", "�ϲ�֢")
                        .TextMatrix(.Rows - 1, 1) = Decode(Val(mrsDiag!��¼��Դ & ""), 1, "����", 2, "��Ժ�Ǽ�", 3, "��ҳ����", "")
                        .TextMatrix(.Rows - 1, 2) = Decode(Val(mrsDiag!������� & ""), 10, "����֢", 1, "��ҽ�������", 2, "��ҽ��Ժ���", 3, "��ҽ��Ժ���", 11, "��ҽ�������", 12, "��ҽ��Ժ���", 13, "��ҽ��Ժ���")
                        .TextMatrix(.Rows - 1, 3) = mrsDiag!������� & ""
                        .RowData(.Rows - 1) = Val(mrsDiag!����id & "")
                        .TextMatrix(.Rows - 1, 4) = Val(mrsDiag!���id & "")
                        mrsDiag.MoveNext
                    Loop
                    .Row = .FixedRows
                    'ֻ��һ����Ϻ�һ��·����
                    If .Rows = .FixedRows + 1 And vsPath.Rows = vsPath.FixedRows + 1 Then
                        vsPath.Row = vsPath.Rows - 1
                        cmdOK.Tag = "Unload"
                        cmdOK_Click
                    End If
                End If
            End With
        Else
            blnVisble = False
            If mrsPath.RecordCount > 0 Then mrsPath.MoveFirst
        
            With vsPath
                .Rows = .FixedRows
                For i = 1 To mrsPath.RecordCount
                    'ȱʡ��ѡ���κ�һ��
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = mrsPath!����
                    .TextMatrix(.Rows - 1, 2) = mrsPath!����
                    .TextMatrix(.Rows - 1, 3) = "" & mrsPath!˵��
                    .RowData(.Rows - 1) = mrsPath!ID & ":" & mrsPath!���°汾
                    If mlngFun = 1 Then
                        If mrsPath!ID = mlngPathID Then .Row = i
                    End If
                    
                    mrsPath.MoveNext
                Next
                
                If mlngFun = 0 Then
                    cmdCancel.Visible = False
                    cmdOK.Left = cmdCancel.Left
                    
                    'ֻ��һ��·����
                    If .Rows = .FixedRows + 1 Then
                        .Row = .Rows - 1
                        cmdOK.Tag = "Unload"
                        cmdOK_Click
                    End If
                End If
            End With
        End If
    Else
        With vsPath
            .Rows = .FixedRows
            If mlngFun = 3 Then
                For i = 1 To mrsMerge.RecordCount
                    'ȱʡ��ѡ���κ�һ��
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = mrsMerge!����
                    .TextMatrix(.Rows - 1, 2) = mrsMerge!����
                    .TextMatrix(.Rows - 1, 3) = "" & mrsMerge!˵��
                    .RowData(.Rows - 1) = mrsMerge!ID & ""
                    mrsMerge.MoveNext
                Next
            ElseIf mlngFun = 4 Then
                If mrsMerge.RecordCount > 0 Then mrsMerge.MoveFirst
                For i = 1 To mrsMerge.RecordCount
                    If Val(mrsMerge!��ʾ & "") = 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, 1) = mrsMerge!����
                        .TextMatrix(.Rows - 1, 2) = mrsMerge!����
                        .TextMatrix(.Rows - 1, 3) = "" & mrsMerge!˵��
                        .TextMatrix(.Rows - 1, 4) = Val(mrsMerge!�����޸� & "")
                        .RowData(.Rows - 1) = mrsMerge!ID & ""
                        If Val(mrsMerge!ѡ�� & "") = 1 Then
                            .Cell(flexcpChecked, .Rows - 1, 0) = 1
                        End If
                    End If
                    mrsMerge.MoveNext
                Next
            End If
            lblPath.Caption = "������кϲ�·����ѡ��:"
            If mblnChoose Then
                vsPath.Editable = flexEDKbdMouse
                .ColHidden(0) = False
                .ColWidth(2) = .ColWidth(2) - .ColWidth(0)
            End If
            'ֻ��һ��·����,��ɲ��Զ���ѡ
            If .Rows = .FixedRows + 1 And mlngFun = 3 Then
                .Row = .Rows - 1
                If mblnChoose Then
                    .Cell(flexcpChecked, .Row, 0) = 1
                End If
                cmdOK.Tag = "Unload"
                cmdOK_Click
            End If
        End With
    End If
    If mbln��� Then mbln��� = False: Exit Sub
    If Not blnVisble Then
        'ֻ��һ���ϲ�֢�����ǵ�����Ҫ��ϻ���ת����������б�
        lblDiag.Visible = False
        vsDiag.Visible = False
        lblPath.Top = lblDiag.Top
        vsPath.Top = vsDiag.Top
        vsPath.Height = vsPath.Height + 800
        Me.Height = vsPath.Top + vsPath.Height + picBottom.Height + 450
    End If
            
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '������Ҫ���ʱ������ȡ����ť��ֻ�ܵ�ȷ��
    If mlngFun = 0 And mblnOK = False And mrsMerge Is Nothing Then
        Cancel = 1
    Else
        Set mrsPath = Nothing
        Set mrsPati = Nothing
        Set mfrmParent = Nothing
        Set mrsDiag = Nothing
        mlng����ID = 0
        mlng���ID = 0
    End If
End Sub

Private Sub vsDiag_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsPath
        If mlngFun = 2 And NewRow >= .FixedRows Then
            .Rows = .FixedRows
            mrsPath.Filter = "����ID=" & Val(vsDiag.RowData(NewRow) & "") & " Or ���ID=" & Val(vsDiag.TextMatrix(NewRow, 4))
            If mrsPath.RecordCount > 0 Then mrsPath.MoveFirst: mlng����ID = Val(vsDiag.RowData(NewRow) & ""): mlng���ID = Val(vsDiag.TextMatrix(NewRow, 4))
            For i = 1 To mrsPath.RecordCount
                'ȱʡ��ѡ���κ�һ��
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = mrsPath!����
                .TextMatrix(.Rows - 1, 2) = mrsPath!����
                .TextMatrix(.Rows - 1, 3) = "" & mrsPath!˵��
                .RowData(.Rows - 1) = mrsPath!ID & ":" & mrsPath!���°汾
                
                mrsPath.MoveNext
            Next
        End If
    End With
End Sub

Private Sub vsPath_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If mblnTmp Then
        If Row > 0 And Not mrsMerge Is Nothing And Col = 0 And mblnChoose And (mlngFun = 3 Or mlngFun = 4) Then
            If vsPath.Cell(flexcpChecked, Row, 0) = 1 Then
                vsPath.Cell(flexcpChecked, Row, 0) = 0
            Else
                vsPath.Cell(flexcpChecked, Row, 0) = 1
            End If
        End If
        mblnTmp = False
    End If
End Sub

Private Sub vsPath_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsPath_DblClick()
    Call cmdOK_Click
End Sub

Private Sub vsPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If vsPath.Row > 0 And Not mrsMerge Is Nothing And mblnChoose And (mlngFun = 3 Or mlngFun = 4) Then
            If vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 1 Then
                If vsPath.TextMatrix(vsPath.Row, 4) <> "1" Then
                    vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 0
                End If
            Else
                vsPath.Cell(flexcpChecked, vsPath.Row, 0) = 1
            End If
            mblnTmp = True
            '����깴ѡһ�Σ��ڰ��ո�����������AfterEdit�����ѡ�������а��ո���û�����⡣
        End If
    End If
End Sub



Private Sub vsPath_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsPath.TextMatrix(Row, 4) = "1" Then
        MsgBox "�úϲ�·���Ѿ��ﵽ��׼סԺ��,���費��������ѡ����һ�׶��Ӻ�", vbInformation, "�ϲ�·������"
        Cancel = True
        Exit Sub
    End If
End Sub


