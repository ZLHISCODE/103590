VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ʷ���ݵ��빤��"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3645
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�� ��(&E)"
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   990
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      _cx             =   5318
      _cy             =   4895
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      Editable        =   2
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
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ϵͳ�ļ������"
      Height          =   195
      Left            =   1155
      TabIndex        =   5
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ӧ��"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�뽫��ϵͳ�ļ������ "
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InitVfgData()
'��ʼ��VfgData��ʽ
    Dim i As Integer
    
    With vfgData
        .Width = 3150
        .Height = 2700
    
        .ColWidth(0) = 300
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .Cols = 3
        
        .TextMatrix(0, 1) = "ԭϵͳ�������"
        .TextMatrix(0, 2) = "��ϵͳ�������"
        
        .FixedCols = 2
    
    End With
    
    For i = 1 To vfgData.Rows - 1
        vfgData.TextMatrix(i, 0) = i
    Next

End Sub

Private Sub LoadOldPathologyData()
'�����ϲ���������
    Dim strSql As String
    Dim rsOldPathData As ADODB.Recordset
    Dim i As Integer
    
    strSql = "select ���� from Ӱ�������"
    
    Set rsOldPathData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If Not rsOldPathData Is Nothing Then
        For i = 1 To rsOldPathData.RecordCount
            'ѭ���������ݼ�
             vfgData.TextMatrix(i, 1) = rsOldPathData("����").Value
             
             If Not rsOldPathData.EOF Then
                rsOldPathData.MoveNext
             End If
        Next
        
        '��̬���ÿؼ�����
        vfgData.Rows = rsOldPathData.RecordCount + 1
        '������ϵͳ�̶���鷽��
        vfgData.ColComboList(2) = "����|����|ϸ��|����|ʬ��|����ʯ��|"
        
        If rsOldPathData.RecordCount > 0 Then
          '������ʾ��Ϣ
            With lblPrompt
                .Font.Bold = True
                .Caption = "���ݿ����ӳɹ�,��������Ѽ���!"
            End With
            
            '�̶��о�����ʾ
            vfgData.ColAlignment(1) = flexAlignCenterCenter
            
            Exit Sub
        End If
    
        '������ʾ��Ϣ
        With lblPrompt
            .Font.Bold = True
            .Caption = "���ݿ����ӳɹ�,��������ͼ�¼Ϊ��!"
        End With
    
        
    End If
End Sub

Private Sub cmdCancel_Click()
'�ر����ݿ�� ж�ش���
On Error GoTo errHandle

    If OraDataClose Then
        Unload Me
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim strDecode As String
    Dim intCheckType As Integer
    Dim strSql As String
    Dim rsPathologyData As ADODB.Recordset
    Dim intPathDataCount As Integer
    Dim i As Integer
    
    For i = 1 To vfgData.Rows - 1
        '�ж��Ƿ������ѡ��
        If vfgData.TextMatrix(i, 2) = "" Then
            MsgBox "��ѡ���Ӧ�ļ�����ͣ�"
            Exit Sub
        End If
        
        '�ж���ϵͳ������Ͷ�Ӧ�ı��
        Select Case vfgData.TextMatrix(i, 2)
            Case "����"
                intCheckType = 0
            Case "����"
                intCheckType = 1
            Case "ϸ��"
                intCheckType = 2
            Case "����"
                intCheckType = 3
            Case "ʬ��"
                intCheckType = 4
            Case "����ʯ��"
                intCheckType = 5
        End Select

        strDecode = strDecode + ",'" & vfgData.TextMatrix(i, 1) & "'" & ",'" & intCheckType & "'"
    Next
    
    strDecode = "decode(���������" & strDecode & ")"
    
    
    If MsgBox("��ȷ����ʼ�������ݵ��������(�˲���������,������!)", vbOKCancel + vbDefaultButton2) = vbOK Then

        '��ִ�й����н���ȷ�Ϻ��˳���ť
        cmdOK.Enabled = False
        cmdCancel.Enabled = False

        'ִ��ǰ�õ�����¼���¼��
        strSql = "select count(*) as ��¼�� from ��������Ϣ"
        Set rsPathologyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        intPathDataCount = Val(rsPathologyData("��¼��").Value)
        
        '��ʼ����
        gcnOracle.BeginTrans
    On Error GoTo errTrans
        
        '��ʾ��ͬ����״̬��Ϣ
        With lblPrompt
            .Font.Bold = True
            .Caption = "��ʼ���벡��걾��Ϣ����,���Ժ�....."
        End With
        
        'ִ�е��벡������SQL���
        Call gcnOracle.Execute("insert into ����걾��Ϣ(�걾ID, ҽ��ID, �걾����,�걾����,����,��������) " & _
                        " select ����걾��Ϣ_�걾ID.Nextval,a.ҽ��ID,a.�걾��λ,0,a.����,b.����ʱ�� " & _
                        " from Ӱ����걾 a, Ӱ��걾����ȡ�� b where a.ҽ��id=b.ҽ��id " & _
                        " and not exists(Select 1 From ����걾��Ϣ where ҽ��ID=a.ҽ��ID and �걾����=a.�걾��λ and ����=a.���� and ��������=b.����ʱ��)")
                        
        With lblPrompt
            .Font.Bold = True
            .Caption = "��ʼ���벡������Ϣ����,���Ժ�....."
        End With
            
        Call gcnOracle.Execute("insert into ��������Ϣ(����ҽ��ID,�����,ҽ��ID,�������,�޼�����,ʣ��λ��) " & _
                               " select ��������Ϣ_����ҽ��ID.Nextval,�����,ҽ��ID," & strDecode & ",�޼�����,ʣ��걾λ�� " & _
                               " from Ӱ��걾����ȡ�� where ҽ��ID not in(select ҽ��ID from ��������Ϣ)")
        With lblPrompt
            .Font.Bold = True
            .Caption = "��ʼ���벡���ͼ���Ϣ����,���Ժ�....."
        End With
                               
        Call gcnOracle.Execute("insert into �����ͼ���Ϣ(ID,ҽ��ID,�ͼ쵥λ,�ͼ����,�ͼ���,�ͼ�����,�Ǽ���,����״̬,����ԭ��,��ע) " & _
                               " select �����ͼ���Ϣ_id.nextval,ҽ��ID,'��Ժ','', 'δ¼��',����ʱ��,decode(���ռ�ʦ,null,'δ¼��',���ռ�ʦ),decode(�������,'1','1','0'),����ԭ��,��ע " & _
                               " from Ӱ��걾����ȡ�� a where not exists(Select 1 From �����ͼ���Ϣ where ҽ��ID=a.ҽ��id and �ͼ�����=a.����ʱ�� and ����ԭ��=a.����ԭ��)")
        
        Call gcnOracle.Execute("update ����걾��Ϣ a set �ͼ�ID=(select  ID from �����ͼ���Ϣ where ҽ��ID=a.ҽ��ID and rownum=1)")
        
        '�ύ����
        gcnOracle.CommitTrans
        GoTo transOk
errTrans:
   gcnOracle.RollbackTrans
transOk:
        'ִ�еõ������¼��SQL
        strSql = "select count(*) as ��¼�� from ��������Ϣ"
        Set rsPathologyData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        '���㵼��ǰ�͵����ļ�¼��
        intPathDataCount = Val(rsPathologyData("��¼��").Value) - intPathDataCount

        With lblPrompt
            .Font.Bold = True
            .Caption = "��ʷ������ȫ������,������" & intPathDataCount & "����¼"
            
        End With

        '��ִ����ɺ������˳���ť
        cmdOK.Enabled = True
        cmdCancel.Enabled = True

        Exit Sub
        
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'���س�ʼ������
    Call InitVfgData
    Call LoadOldPathologyData
End Sub

