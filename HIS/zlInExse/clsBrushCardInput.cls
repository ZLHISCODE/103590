VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrushCardInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*************************************************************************************************
'����:ˢ������
'����:���˺�
'����:2016-08-12 15:49:37
'˵��:
'    1. ����˳��:�ȵ���"InitCompents"
'    2. ����:ReadCardNoed����ȡ�����Ŀ���)
'    3.������˵��:
'       1).��ʼ����:�ı������2000����ʱ��ʶ���һ���ַ�ʱ��Ϊ��ʼˢ��
'       2).��ֹ����:�ض��ַ��ͼ������0.04����ʱ��
'       3).������Ч�ַ�(ȱʡ):��ĸ������
'       4).�����ַ�(ȱʡ): ?��:��;��
'*************************************************************************************************
Private WithEvents mobjBrush As BillEdit  'ˢ���ı������
Attribute mobjBrush.VB_VarHelpID = -1
Private mstrInputText As String   'ˢ���������ַ���(ȫ��������һЩ�����ַ�)
Private mstrCardNO As String       'ˢ����������Ч����
Private msngBeginTime As Single '��ʼˢ��ʱ��
Private msngPreInputCharTime As Single '�ϴ������ַ�ʱ��
Private mstrPreChar As String   '�ϸ��ַ�
Private mobjMain As Object
Private mobjCard As Card
Private mintInputChar As Integer '�ַ���������:0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�
Private mstrSpecialChar As String '�ض��ַ�:Ascii��1��Ascii��2....
Private mblnEndReceiveChar As Boolean '���������ַ�
Private WithEvents mobjTime As Timer
Attribute mobjTime.VB_VarHelpID = -1

'�¼�����
Public Event ReadCardNoed(ByVal strCardNo As String, ByVal blnBrushCard As Boolean)
'��������ֵ�ľֲ�����
Private mvarOnlyLegalCardNo As Boolean '�ֲ�����:������Ϸ��Ŀ����ַ�(�����ֹ����뻹��ˢ������)

Public Property Let OnlyLegalCardNo(ByVal vData As Boolean)
Attribute OnlyLegalCardNo.VB_Description = "������Ϸ��Ŀ����ַ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.OnlyLegalCardNo = 5
    mvarOnlyLegalCardNo = vData
End Property

Public Property Get OnlyLegalCardNo() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.OnlyLegalCardNo
    OnlyLegalCardNo = mvarOnlyLegalCardNo
End Property

Public Sub InitCompents(ByRef objMain As Object, ByRef objBrush As Object, objCard As Card)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '���:objMain-���õ�������
    '    txtBrushObj-ˢ���ı���
    '    objCard-��������
    '����:���˺�
    '����:2016-08-12 16:01:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��̬����ʱ��ؼ�
    Set mobjTime = objMain.Controls.Add("VB.Timer", "mtimEnd")
    With mobjTime
        .Enabled = False
        .Interval = 500
    End With
   
    Set mobjCard = objCard
    Set mobjMain = objMain
    Set mobjBrush = objBrush
    '��������(0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�)|Ascii��1��Ascii��2....
'    mintInputChar = Val(Split(mobjCard.������Ч�ַ� & "|", "|")(0))
'    mstrSpecialChar = Trim(Split(mobjCard.������Ч�ַ� & "|", "|")(1))
    mblnEndReceiveChar = False
End Sub
 

Private Sub mobjBrush_EditChange(strInput As String)
    If Trim(mobjBrush.Text) <> "" Then Exit Sub
    mstrInputText = "": mstrCardNO = "": msngBeginTime = 0
    mblnEndReceiveChar = False
    mobjTime.Enabled = False
End Sub

Private Sub mobjBrush_EditKeyPress(KeyAscii As Integer)
    Dim sngNow As Single, blnCard As Boolean
    Dim blnAdd As Boolean
    Dim intKeyAscii As Integer
    
    
    sngNow = Timer: blnCard = False: mobjTime.Enabled = False
    
    intKeyAscii = KeyAscii
    If Format((sngNow - msngPreInputCharTime), "0.000") > 2 Or KeyAscii = 8 Or mstrInputText = "" Or msngBeginTime = 0 Then
        mstrInputText = "": mstrCardNO = "": msngBeginTime = 0
        mblnEndReceiveChar = False
    End If
    
    sngNow = 0
    msngPreInputCharTime = Timer
    
    If mstrInputText = "" Then
        msngBeginTime = Timer
       ' Debug.Print "msngBeginTime:" & msngBeginTime & "-->" & mstrInputText
    ElseIf KeyAscii > 32 Or KeyAscii = 13 Then
        sngNow = Timer
        blnCard = Format((sngNow - msngBeginTime) / IIf(mstrInputText = "", 1, Len(mstrInputText)), "0.000") < 0.04  'һ����0.014����
        'Debug.Print "blnCard=" & blnCard & " ---->" & KeyAscii & " -->" & "sngNow:" & sngNow & " - " & msngBeginTime & "=" & (sngNow - msngBeginTime) & "  -->" & mstrInputText
    End If
    
    If Not blnCard Then  '��ǰ����ˢ�������ֹ�����
        If KeyAscii = 13 Or KeyAscii = 8 Then
            mstrCardNO = "": mstrInputText = "": msngBeginTime = 0
        Else
'            If InStr(1, ":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
            If KeyAscii <> 8 And KeyAscii <> 0 Then
                mstrInputText = Chr(KeyAscii): mstrPreChar = Chr(KeyAscii)
            End If
            mstrCardNO = ""
        End If
        '�����������ԣ��Ƿ�ﵽ���ų���ʱ���ֹ�����ﵽ����ʱ��Ҳ�Զ��ң���ҪǸ����ǰ�ģ������������ﵽ���ȣ��ͻ��Զ����ң�
        If (Len(mobjBrush.Text) = IIf(mobjCard.���ų��� = 0, 100, mobjCard.���ų���) - 1 And IsNumeric(mobjBrush.Text)) Or KeyAscii = 13 Then
            '�ﵽ���ȣ�Ҳȥ��,�������Ա��Ҫ��һ�»س�
            mstrCardNO = mobjBrush.Text & Chr(KeyAscii)
            RaiseEvent ReadCardNoed(mstrCardNO, False)
            mstrCardNO = "": mstrInputText = ""
        End If
        
        'ֻ���պϷ��Ŀ�����Ϣ
'        If mvarOnlyLegalCardNo Then If CheckInputCharValied(KeyAscii) = False Then KeyAscii = 0
        Exit Sub
    End If
    
    'ˢ������
     If mstrPreChar <> "" Then  '��һ����Ч�����ַ�
        If CheckInputCharValied(Asc(mstrPreChar)) = False Then mstrPreChar = ""
        mstrCardNO = mstrPreChar: mstrPreChar = ""
    End If
    If KeyAscii = 13 Then    '��ֱ�ӽ�����־�ģ���ֱ��ȷ��
'        KeyAscii = 0
        mobjTime.Enabled = False
        Call mobjTime_Timer
        Exit Sub
    End If
    If KeyAscii <> 8 And KeyAscii > 32 Then
        mstrInputText = mstrInputText & Chr(KeyAscii)
    End If
    
    mobjTime.Enabled = True
    'ֻ���պϷ��Ŀ�����Ϣ
'    If CheckInputCharValied(KeyAscii) = False Then
'        intKeyAscii = 0
'        If mvarOnlyLegalCardNo Then KeyAscii = 0
'        Exit Sub
'    End If
    If Len(mstrCardNO) > IIf(mobjCard.���ų��� = 0, 100, mobjCard.���ų���) - 1 Then
        intKeyAscii = 0: Exit Sub '���￨�ų���ʱ�����ڽ����ַ�
    End If
    If KeyAscii <> 0 Then mstrCardNO = mstrCardNO & Chr(KeyAscii) '�Ϸ�����
End Sub
Private Function CheckInputCharValied(ByRef KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ַ��Ƿ�Ϸ�
    '����:�Ϸ�������true,���򷵻�False
    '����:���˺�
    '����:2016-08-17 16:42:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    CheckInputCharValied = False
    If KeyAscii = 8 Or KeyAscii = 13 Then CheckInputCharValied = True: Exit Function
    
    '1.�����ַ��ж�
    If InStr(1, ":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    
    '2.���ŵĽ�����������:213442=222 ���������Ϊ:=,��ȡ213442
'    If InStr(1, "," & mobjCard.���Ž����� & ",", "," & KeyAscii & ",") > 0 Then mblnEndReceiveChar = True
    If mblnEndReceiveChar Then Exit Function
    
    '3.��Ч�ַ��ж�
'    If InStr(1, "," & mobjCard.������Ч�ַ� & ",", "," & KeyAscii & ",") > 0 Then Exit Function


    '4.��Ч�ַ��ж�: mintInputChar:0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�
    If mintInputChar = 1 And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then Exit Function   '����
    If mintInputChar = 2 And (Not ((KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")))) Then Exit Function        '��ĸ
   
    If mintInputChar = 3 Then '���ֻ���ĸ
       If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) _
            Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z"))) Then Exit Function
    End If

    '6.ֻ����ָ���ַ�
    If mstrSpecialChar <> "" And (InStr(1, "," & mstrSpecialChar & ",", "," & KeyAscii & ",") = 0) Then Exit Function
    
    CheckInputCharValied = True
End Function
Private Sub mobjTime_Timer()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ͨ��ʱ�����ж��Ƿ�ˢ������
    '����:���˺�
    '����:2016-08-17 16:55:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mobjTime.Enabled = False
    
'    If zlCommFun.ActualLen(mstrCardNO) < mobjCard.������С���� Then Exit Sub
    
    '��Ч�ַ��ж�: mintInputChar:0-�����ַ�,1-����,2-��ĸ;3-���ֻ���ĸ;4-ָ���ַ�
    If mintInputChar = 1 And IsNumeric(mstrCardNO) = False Then Exit Sub      '����
    If mintInputChar = 2 And zlCommFun.IsCharAlpha(mstrCardNO) = False Then Exit Sub       '��ĸ
    If mintInputChar = 3 And zlCommFun.IsNumOrChar(mstrCardNO) = False Then Exit Sub         '���ֺ���ĸ
    
    
    If Not mobjBrush Is Nothing Then
        mobjBrush.Text = mstrCardNO
        mobjBrush.SelStart = Len(mstrCardNO)
    End If
    mobjTime.Enabled = False:   mblnEndReceiveChar = False
    If mstrCardNO = "" Then mstrInputText = "": Exit Sub
    RaiseEvent ReadCardNoed(mstrCardNO, True)
    mstrCardNO = "": mstrInputText = ""
End Sub
 

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mobjTime Is Nothing Then Unload mobjTime
    Set mobjBrush = Nothing
    Set mobjTime = Nothing
End Sub