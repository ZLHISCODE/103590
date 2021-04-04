VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPathTrackView 
   AutoRedraw      =   -1  'True
   Caption         =   "�����ٴ�·��"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10860
   Icon            =   "frmPathTrackView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10860
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeSuiteControls.TabControl tbcPath 
      Height          =   3090
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   5475
      _Version        =   589884
      _ExtentX        =   9657
      _ExtentY        =   5450
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPathTrackView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmPath As Object
Private mbyt���� As Byte                    '0-סԺ�ٴ�·������;1-�����ٴ�·������

Public Sub ShowMe(frmParent As Object, vPati As TYPE_Pati, ByVal blnMoved As Boolean, Optional ByVal byt���� As Byte = 0)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    mbyt���� = byt����
    If mbyt���� = 0 Then
        strSql = "Select NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.סԺ��,B.��Ժ���� as ����," & _
            " C.���� as ����,B.��Ժ����,B.��Ժ����" & _
            " From ������Ϣ A,������ҳ B,���ű� C" & _
            " Where A.����ID=B.����ID And B.��Ժ����ID=C.ID And A.����ID=[1] And B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vPati.����ID, vPati.��ҳID)
        Me.tbcPath.Item(0).Caption = "������" & rsTmp!���� & "���Ա�" & Nvl(rsTmp!�Ա�) & "�����䣺" & Nvl(rsTmp!����) & _
            "�����ң�" & rsTmp!���� & "��סԺ�ţ�" & Nvl(rsTmp!סԺ��) & "�����ţ�" & Nvl(rsTmp!����) & _
            "����" & vPati.��ҳID & "��סԺ��" & Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm") & _
            IIf(Not IsNull(rsTmp!��Ժ����), "-" & Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm"), "")
        
        With vPati
            Call mfrmPath.zlRefresh(.����ID, .��ҳID, .����ID, .����ID, .����״̬, blnMoved)
        End With
    Else
        strSql = "Select NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.�����,C.���� as ���� " & _
            " From ������Ϣ A,���˹Һż�¼ B,���ű� C" & _
            " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=[1] And B.ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vPati.����ID, vPati.�Һ�ID)
        Me.tbcPath.Item(0).Caption = "������" & rsTmp!���� & "���Ա�" & Nvl(rsTmp!�Ա�) & "�����䣺" & Nvl(rsTmp!����) & _
            "�����ң�" & rsTmp!���� & "������ţ�" & Nvl(rsTmp!�����)
        
        With vPati
            Call mfrmPath.zlRefresh(.����ID, .�Һ�ID, .�Һ�NO, .����ID, .����״̬, blnMoved)
        End With
    End If
    Me.Show , frmParent
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mbyt���� = 0 Then
        Set mfrmPath = New frmPathTable
    Else
        Set mfrmPath = New frmPathTableOut
    End If
    'TabControl
    '-----------------------------------------------------
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "�����ٴ�·��", mfrmPath.Hwnd, 0
    End With
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    Me.tbcPath.Left = 0
    Me.tbcPath.Top = 0
    Me.tbcPath.Width = Me.ScaleWidth
    Me.tbcPath.Height = Me.ScaleHeight
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Unload mfrmPath
    Set mfrmPath = Nothing
End Sub
