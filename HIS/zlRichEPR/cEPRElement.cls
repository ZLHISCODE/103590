VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarKey As Long                 '������
Private mvarID As Double                  'ID
Private mvar�ļ�ID As Long              '�ļ�ID
Private mvar��ID As Double               '��ID
Private mvar������� As Double            '���
Private mvar�������� As Boolean         '��������
Private mvar�����ı� As String          '���ֵ�������ı�
Private mvar�Ƿ��� As Boolean         '�Ƿ���
Private mvar����Ҫ��ID As Long          '����Ҫ��ID
Private mvarǩ��Ҫ�� As Boolean         '���õ�ǩ��Ҫ��
Private mvar�滻�� As Integer           '�Ƿ����滻��
Private mvarҪ������ As String          'Ҫ������
Private mvarҪ������ As Integer         'Ҫ������       0:��ֵ 1:���� 2:����
Private mvarҪ�س��� As Long            'Ҫ�س���
Private mvarҪ��С�� As Long            'Ҫ��С��λ��
Private mvarҪ�ص�λ As String          'Ҫ�ص�λ
Private mvarҪ�ر�ʾ As Integer         'Ҫ�ر�ʾ       0-�ı� 1-���� 2-��ѡ 3-��ѡ
Private mvar������̬ As Integer         '������̬       0-����;1-չ��(��ʾ��ѡ,��ѡ������Ҫ�ص�������̬��չ������ѡ����Ŀ)
Private mvarҪ��ֵ�� As String          'Ҫ��ֵ��       ��ֵ��Ŀ�����������Сֵ�����ֵ������������ѡ��Ŀ��ѡ����Ŀ�б�
Private mvar��ʼ�� As Long              '��ʼ�棬ר���ڡ����Ӳ�����¼��
Private mvar��ֹ�� As Long              '��ֹ�棬ר���ڡ����Ӳ�����¼��
Private mvar�Զ�ת�ı� As Boolean       '�����滻Ҫ�ض��ԣ��Ƿ�������滻���Զ�ת��Ϊ�ı�
Private mvar����   As Byte            '�Ƿ����Ҫ�ء�1-���� 0-�Ǳ���
Private mvar��̬�� As Byte              '�Ƿ������Զ��� 1-���� 0-������
Private mvar��ʾ As String              '��Ҫ�ر༭��������ʾ

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Double)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Double
    ��ID = mvar��ID
End Property

Public Property Let �������(ByVal vData As Double)
    mvar������� = vData
End Property

Public Property Get �������() As Double
    ������� = mvar�������
End Property

Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property

Public Property Let ǩ��Ҫ��(ByVal vData As Boolean)
    mvarǩ��Ҫ�� = vData
End Property

Public Property Get ǩ��Ҫ��() As Boolean
    ǩ��Ҫ�� = mvarǩ��Ҫ��
End Property

Public Property Let �����ı�(ByVal vData As String)
    '�ı���Ч�Լ�⣡
    mvar�����ı� = vData
End Property

Public Property Get �����ı�() As String
    �����ı� = mvar�����ı�
End Property

Public Property Let �Ƿ���(ByVal vData As Boolean)
    mvar�Ƿ��� = vData
End Property

Public Property Get �Ƿ���() As Boolean
    �Ƿ��� = mvar�Ƿ���
End Property

Public Property Let ����Ҫ��ID(ByVal vData As Long)
    mvar����Ҫ��ID = vData
End Property

Public Property Get ����Ҫ��ID() As Long
    ����Ҫ��ID = mvar����Ҫ��ID
End Property

Public Property Let �滻��(ByVal vData As Integer)
    mvar�滻�� = vData
End Property

Public Property Get �滻��() As Integer
    �滻�� = mvar�滻��
End Property

Public Property Let Ҫ������(ByVal vData As String)
    mvarҪ������ = vData
End Property

Public Property Get Ҫ������() As String
    Ҫ������ = mvarҪ������
End Property

Public Property Let Ҫ������(ByVal vData As Integer)
    mvarҪ������ = vData
End Property

Public Property Get Ҫ������() As Integer
    Ҫ������ = mvarҪ������
End Property

Public Property Let Ҫ�س���(ByVal vData As Long)
    mvarҪ�س��� = vData
End Property

Public Property Get Ҫ�س���() As Long
    Ҫ�س��� = mvarҪ�س���
End Property

Public Property Let Ҫ��С��(ByVal vData As Long)
    mvarҪ��С�� = vData
End Property

Public Property Get Ҫ��С��() As Long
    Ҫ��С�� = mvarҪ��С��
End Property

Public Property Let Ҫ�ص�λ(ByVal vData As String)
    mvarҪ�ص�λ = vData
End Property

Public Property Get Ҫ�ص�λ() As String
    Ҫ�ص�λ = mvarҪ�ص�λ
End Property

Public Property Let Ҫ�ر�ʾ(ByVal vData As Integer)
    mvarҪ�ر�ʾ = vData
End Property

Public Property Get Ҫ�ر�ʾ() As Integer
    Ҫ�ر�ʾ = mvarҪ�ر�ʾ
End Property

Public Property Let ������̬(ByVal vData As Integer)
    mvar������̬ = vData
End Property

Public Property Get ������̬() As Integer
    ������̬ = mvar������̬
End Property

Public Property Let Ҫ��ֵ��(ByVal vData As String)
    mvarҪ��ֵ�� = vData
End Property

Public Property Get Ҫ��ֵ��() As String
    Ҫ��ֵ�� = mvarҪ��ֵ��
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = IIf(mvar��ʼ�� = 0, 1, mvar��ʼ��)
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
    ��ֹ�� = IIf(mvar��ֹ�� = 0, 0, mvar��ֹ��)
End Property

Public Property Let �Զ�ת�ı�(ByVal vData As Boolean)
    mvar�Զ�ת�ı� = vData
End Property

Public Property Get �Զ�ת�ı�() As Boolean
    �Զ�ת�ı� = mvar�Զ�ת�ı�
End Property
Public Property Let ����(ByVal vData As Byte)
    mvar���� = vData
End Property
Public Property Get ����() As Byte
    ���� = mvar����
End Property
Public Property Let ��̬��(ByVal vData As Byte)
    mvar��̬�� = vData
End Property
Public Property Get ��̬��() As Byte
    ��̬�� = mvar��̬��
End Property
Public Property Let ��ʾ(ByVal vDate As String)
    mvar��ʾ = vDate
End Property
Public Property Get ��ʾ() As String
    ��ʾ = mvar��ʾ
End Property
'################################################################################################################
'## ���ܣ�  �Ƿ��Զ�ת�ı�
'################################################################################################################
Public Property Get ��������() As String
    �������� = IIf(mvar�Զ�ת�ı�, 1, 0) & "|" & mvar���� & "|" & mvar��̬�� & "|" & mvar��ʾ
End Property

'################################################################################################################
'## ���ܣ�  �Ƿ��Զ�ת�ı�
'################################################################################################################
Public Property Let ��������(ByRef vData As String)
    Select Case UBound(Split(vData, "|"))
        Case Is <= 0
            mvar�Զ�ת�ı� = IIf(Val(vData) = 0, False, True)
            mvar���� = 0
            mvar��̬�� = 0
            mvar��ʾ = ""
        Case 1
            mvar�Զ�ת�ı� = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar���� = Split(vData, "|")(1)
            mvar��̬�� = 0
            mvar��ʾ = ""
        Case 2
            mvar�Զ�ת�ı� = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar���� = Split(vData, "|")(1)
            mvar��̬�� = Split(vData, "|")(2)
            mvar��ʾ = ""
        Case 3
            mvar�Զ�ת�ı� = IIf(Val(Split(vData, "|")(0)) = 0, False, True)
            mvar���� = Split(vData, "|")(1)
            mvar��̬�� = Split(vData, "|")(2)
            mvar��ʾ = Split(vData, "|")(3)
    End Select
End Property

'################################################################################################################
'## ���ܣ�  ��������Ҫ�ض���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRElement
    '���ڿ������ݣ�����һ���������
    Dim cNode As New cEPRElement
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.�ļ�ID = Me.�ļ�ID
    cNode.��ʼ�� = Me.��ʼ��
    cNode.��ֹ�� = Me.��ֹ��
    cNode.��ID = Me.��ID
    cNode.������� = Me.�������
    cNode.�������� = Me.��������
    cNode.�����ı� = Me.�����ı�
    cNode.�Ƿ��� = Me.�Ƿ���
    cNode.����Ҫ��ID = Me.����Ҫ��ID
    cNode.�滻�� = Me.�滻��
    cNode.Ҫ������ = Me.Ҫ������
    cNode.Ҫ������ = Me.Ҫ������
    cNode.Ҫ�س��� = Me.Ҫ�س���
    cNode.Ҫ��С�� = Me.Ҫ��С��
    cNode.Ҫ�ص�λ = Me.Ҫ�ص�λ
    cNode.Ҫ�ر�ʾ = Me.Ҫ�ر�ʾ
    cNode.������̬ = Me.������̬
    cNode.Ҫ��ֵ�� = Me.Ҫ��ֵ��
    cNode.�������� = Me.��������
    cNode.���� = Me.����
    cNode.��̬�� = Me.��̬��
    cNode.��ʾ = Me.��ʾ
    Set Clone = cNode
End Function
Public Function FillElementMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'����:�����ݼ�¼������Ա
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!������, 0) <> 0 Then Me.Key = NVL(rs!������, 0)
    Me.ID = rs!ID
    Me.�ļ�ID = NVL(rs!�ļ�ID, 0)
    Me.��ID = NVL(rs!��ID, 0)
    Me.������� = NVL(rs!�������, 0)
    Me.�������� = IIf(NVL(rs!��������, 0) = 0, False, True)
    Me.�����ı� = NVL(rs!�����ı�)
    Me.�Ƿ��� = IIf(NVL(rs!�Ƿ���, 0) = 0, False, True)
    Me.����Ҫ��ID = NVL(rs!����Ҫ��ID, 0)
    Me.�滻�� = NVL(rs!�滻��, 0)
    Me.Ҫ������ = NVL(rs!Ҫ������)
    Me.Ҫ������ = NVL(rs!Ҫ������, 0)
    Me.Ҫ�س��� = NVL(rs!Ҫ�س���, 0)
    Me.Ҫ��С�� = NVL(rs!Ҫ��С��, 0)
    Me.Ҫ�ص�λ = NVL(rs!Ҫ�ص�λ)
    Me.Ҫ�ر�ʾ = NVL(rs!Ҫ�ر�ʾ, 0)
    Me.������̬ = NVL(rs!������̬, 0)
    Me.Ҫ��ֵ�� = NVL(rs!Ҫ��ֵ��)
    Me.�������� = NVL(rs!��������)
        
    If InStr(strTableName, "���Ӳ�������") > 0 Then '�����ļ��ṹ�Ͳ�����������û���������ֶ�,���������������
        Me.��ʼ�� = NVL(rs!��ʼ��, 1)
        Me.��ֹ�� = NVL(rs!��ֹ��, 0)
    End If
    
    FillElementMember = True
End Function
'################################################################################################################
'## ���ܣ�  �����ݿ�ֱ�ӻ�ȡ����Ҫ������
'##
'## ������  eEdtType    :��ǰ�༭��ʽ��
'##         lngID       :���ݿ��еĶ���ID��
'##         blnWithKey  :�Ƿ����Keyֵһ��ָ�
'##         strTableName:���������ʡ�ԣ�����ݵ�ǰ��eEdtType���жϱ���
'################################################################################################################
Public Function GetElementFromDB(ByVal eEdtType As EditTypeEnum, ByVal lngID As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean
    
    Dim rs As New ADODB.Recordset
    If strTableName = "" Then
        Select Case eEdtType
        Case cprET_�����ļ�����
            strTableName = "�����ļ��ṹ"
        Case cprET_ȫ��ʾ���༭
            strTableName = "������������"
        Case cprET_�������༭, cprET_���������
            strTableName = "���Ӳ�������"
        End Select
    End If
    
    gstrSQL = "Select ID, �ļ�id," & IIf(InStr(strTableName, "���Ӳ�������") > 0, "��ʼ��, ��ֹ��,", "") & " ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, ����Ҫ��id," & vbNewLine & _
                "       �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ�� " & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRElement", lngID)
    If Not rs.EOF Then
        If blnWithKey And NVL(rs("������"), 0) <> 0 Then Me.Key = NVL(rs("������"), 0)
        Me.ID = rs("ID")
        Me.�ļ�ID = NVL(rs("�ļ�ID"), 0)
        Me.��ID = NVL(rs("��ID"), 0)
        Me.������� = NVL(rs("�������"), 0)
        Me.�������� = IIf(NVL(rs("��������"), 0) = 0, False, True)
        Me.�����ı� = NVL(rs("�����ı�"))
        Me.�Ƿ��� = IIf(NVL(rs("�Ƿ���"), 0) = 0, False, True)
        Me.����Ҫ��ID = NVL(rs("����Ҫ��ID"), 0)
        Me.�滻�� = NVL(rs("�滻��"), 0)
        Me.Ҫ������ = NVL(rs("Ҫ������"))
        Me.Ҫ������ = NVL(rs("Ҫ������"), 0)
        Me.Ҫ�س��� = NVL(rs("Ҫ�س���"), 0)
        Me.Ҫ��С�� = NVL(rs("Ҫ��С��"), 0)
        Me.Ҫ�ص�λ = NVL(rs("Ҫ�ص�λ"))
        Me.Ҫ�ر�ʾ = NVL(rs("Ҫ�ر�ʾ"), 0)
        Me.������̬ = NVL(rs("������̬"), 0)
        Me.Ҫ��ֵ�� = NVL(rs("Ҫ��ֵ��"))
        Me.�������� = NVL(rs("��������"))
        
        If eEdtType = cprET_��������� Or eEdtType = cprET_�������༭ Then
            Me.��ʼ�� = NVL(rs("��ʼ��"), 1)
            Me.��ֹ�� = NVL(rs("��ֹ��"), 0)
        End If
        GetElementFromDB = True
    Else
        GetElementFromDB = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  �������ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_�����ļ�����
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("�����ļ��ṹ")
    Case cprET_ȫ��ʾ���༭
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("������������")
    Case cprET_�������༭, cprET_���������
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  �������Ҫ�ص�ԭʼ�����Ƿ���ڣ�����XML����ʱ����֤��
'################################################################################################################
Public Function CheckValid() As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select ID From ����������Ŀ Where ID = [1] And ������ = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRElement", Me.ID, Me.Ҫ������)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs!ID > 0)
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ�ɹ���ʾ���ı�����
'##
'## ˵����  ע�ⲻ�������ı��ĳ��ȣ����������ı�Ϊ��ʱҪ��ʾ[Ҫ������]
'################################################################################################################
Public Function GetValidTextLength() As Long
    Dim lngLen As Long, lngLen2 As Long
    If Me.�滻�� = 0 Then
        lngLen = Len(IIf(Me.�����ı� = "", "[" & Me.Ҫ������ & "]", Me.�����ı�))
    Else
        lngLen = Len(IIf(Me.�����ı� = "", "{" & Me.Ҫ������ & "}", Me.�����ı�))
    End If
    lngLen2 = Len(Me.Ҫ�ص�λ)
    GetValidTextLength = lngLen + lngLen2
End Function


Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1, Optional ByVal blnFirstIns As Boolean = False, Optional ByVal blnFromPhrase As Boolean = False, Optional ByVal blnForceInsert As Boolean = False) As Boolean
    '******************************************************************************************************************
    ' ���ܣ�  ��������Ҫ�ص�Editor��
    '
    ' ������  edtThis         :��ǰ�ı༭���ؼ�
    '         lngStartPos     :��ǰλ��
    '         blnFirstIns     :�Ƿ��ǵ�һ�β��루Ĭ�ϵ�һ�β������ĩβ�Ӷ��ţ�
    '         blnFromPhrase   :����ǴӴʾ�ʾ���е��������Ҫ��ΪTrue����ʾ��δ���б༭�����»��ߣ���
    '         blnForceInsert  :�Ƿ�ǿ�Ʋ��룬���޸�Ҫ��ʱ���ǰ�󶼱�������ʱ��Ҫǿ�Ʋ���
    '******************************************************************************************************************
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then InsertIntoEditor = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR _
        And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then Exit Function

    Dim p As Long, lngKey As Long, blnForce As Boolean, lngLen As Long, lngLen2 As Long, strTmp As String
    p = lngStartPos
    blnForce = edtThis.ForceEdit

    If blnFirstIns Then Me.�Ƿ��� = False
    With edtThis
        .Freeze
        .ForceEdit = True
        lngKey = Me.Key
        
        strTmp = Format(lngKey, "00000000") & "," & IIf(Me.��������, 1, 0) & ",0)"

        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        If Me.������̬ = 0 Then
            If Me.�滻�� = 0 Then
                lngLen = Len(IIf(Me.�����ı� = "", "[" & Me.Ҫ������ & "]", Me.�����ı�))
            Else
                lngLen = Len(IIf(Me.�����ı� = "", "{" & Me.Ҫ������ & "}", Me.�����ı�))
            End If
            lngLen2 = Len(Me.Ҫ�ص�λ)
            If Me.�滻�� = 0 Then
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.�����ı� = "", "[" & Me.Ҫ������ & "]", Me.�����ı�) & _
                    Me.Ҫ�ص�λ & "EE(" & strTmp & IIf(Me.�Ƿ���, vbCrLf, "")
            ElseIf Me.�滻�� = 2 Then
                '�ֵ���Ŀ
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.�����ı� = "", "<" & Me.Ҫ������ & ">", Me.�����ı�) & _
                    Me.Ҫ�ص�λ & "EE(" & strTmp & IIf(Me.�Ƿ���, vbCrLf, "")
            Else
                .Range(lngStartPos, lngStartPos).Text = "ES(" & strTmp & _
                    IIf(Me.�����ı� = "", "{" & Me.Ҫ������ & "}", Me.�����ı�) & _
                    Me.Ҫ�ص�λ & "EE(" & strTmp & IIf(Me.�Ƿ���, vbCrLf, "")
            End If
        Else
            If blnFirstIns Then
                Dim strContent As String, i As Long, T As Variant
                T = Split(Me.Ҫ��ֵ��, ";")
                For i = 0 To UBound(T)
                    strContent = strContent & IIf(Me.Ҫ�ر�ʾ = 3, "��", "��") & T(i) & IIf(i = UBound(T), "", "  ")   '������
                Next
                lngLen = Len(strContent)
                lngLen2 = Len(Me.Ҫ�ص�λ)
                Me.�����ı� = strContent
                .Range(lngStartPos, lngStartPos) = "ES(" & strTmp & _
                    strContent & _
                    Me.Ҫ�ص�λ & "EE(" & strTmp & IIf(Me.�Ƿ���, vbCrLf, "")
            Else
                lngLen = Len(Me.�����ı�)
                lngLen2 = Len(Me.Ҫ�ص�λ)
                .Range(lngStartPos, lngStartPos) = "ES(" & strTmp & _
                    Me.�����ı� & _
                    Me.Ҫ�ص�λ & "EE(" & strTmp & IIf(Me.�Ƿ���, vbCrLf, "")
            End If
        End If
        .Range(lngStartPos, lngStartPos + 32 + lngLen + lngLen2).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Hidden = False
        '�»���
        If Me.������̬ = 0 Then
            .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.�����ı� = "" Or blnFromPhrase, ELE_UNDERLINE, cprNone)
        Else
            If Me.Ҫ�ر�ʾ = 2 And InStr(Me.�����ı�, "��") = 0 Then
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprwave
            ElseIf Me.Ҫ�ر�ʾ = 3 And InStr(Me.�����ı�, "��") = 0 Then
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = cprwave
            Else
                .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.�����ı� = "" Or blnFromPhrase, ELE_UNDERLINE, cprNone)
            End If
        End If
        'ɾ���ߵ�����
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.Strikethrough = (Me.��ֹ�� > 0)
        'ǰ��ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.ForeColor = GetCharColor(Me.��ʼ��, Me.��ֹ�� + 1)
        '����ɫ
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen + lngLen2).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen + lngLen2, lngStartPos + 32 + lngLen + lngLen2).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen + lngLen2

        If Me.�Ƿ��� Then
            .Range(lngStartPos, lngStartPos + 2).Font.Protected = False
            .Range(lngStartPos, lngStartPos + 2).Font.Hidden = False
            lngStartPos = lngStartPos + 2
        End If
        
        Me.ǩ��Ҫ�� = False
        
        If Me.�滻�� = 1 Then
            Select Case Me.Ҫ������
            Case "����ҽʦǩ��", "����ҽʦǩ��", "����ҽʦǩ��"
                If Me.�����ı� <> "" Then
                    Dim lESS As Long
                    Dim lESE As Long
                    Dim lEES As Long
                    Dim lEEE As Long
                    
                    Dim blnNeeded As Boolean
                    Dim blnFinded As Boolean
        
                    blnFinded = FindKey(edtThis, "E", Me.Key, lESS, lESE, lEES, lEEE, blnNeeded)
                    
                    If blnFinded Then
                        .Range(lESS, lEEE).Font.Protected = True
                        .Range(lESS, lEEE).Font.Hidden = True
                    End If
                Else
                    '�п��õ�ǩ��Ҫ��
                    Me.ǩ��Ҫ�� = True
                End If
            End Select
        End If
        
        .Range(lngStartPos, lngStartPos).Selected

        .ForceEdit = blnForce
        .UnFreeze
    End With
    InsertIntoEditor = True
End Function

'################################################################################################################
'## ���ܣ�  ˢ����ʾ
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = edtThis.Range(lKSE, lKES).Font.GetFontFmt
            CheckFonts FontFmt
            If FontFmt.Underline = cprwave Then FontFmt.Underline = cprNone
            
            .Range(lKSS, lKEE) = ""
            Me.�Ƿ��� = False
            Me.InsertIntoEditor edtThis, lKSS, , , True
            If InStr(",����ҽʦǩ��,����ҽʦǩ��,����ҽʦǩ��", Me.Ҫ������) = 0 Then '��ǩ��Ҫ�صı༭ˢ��
                .Range(lKSE, lKSE + Len(Me.�����ı�) + Len(Me.Ҫ�ص�λ)).Selected
                .Range(lKSE, lKSE + Len(Me.�����ı�) + Len(Me.Ҫ�ص�λ)).Font.SetFontFmt FontFmt
                .Range(lKSE, lKSE + Len(Me.�����ı�) + Len(Me.Ҫ�ص�λ)).Para.SetParaFmt ParaFmt
            End If
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  �ӱ༭����ɾ��
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            .Range(lKSS, lKEE) = ""
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ��ٱ��浽���ݿ��SQL���
'##
'## ������  eEdtType    :��ǰ�༭��ʽ��
'## ����ֵ������ �洢���̵��õ�SQL��䡣
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, ByRef ArraySQL() As String) As Boolean
    Dim lngCount As Long
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String

    Call CheckValidID(eEdtType)         '���ID��Ч�ԣ�
    Select Case eEdtType
    Case cprET_�����ļ�����
        gstrSQL = "Zl_�����ļ��ṹ_Update(" & Me.ID & "," & Me.�ļ�ID & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",4," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.�������� & "',NULL,'" & Replace(Me.�����ı�, "'", "' || chr(39) || '") & "'," & IIf(Me.�Ƿ���, 1, 0) & _
            ",NULL,NULL,NULL," & _
            IIf(Me.����Ҫ��ID = 0, "NULL", Me.����Ҫ��ID) & "," & Me.�滻�� & ",'" & Me.Ҫ������ & "'," & Me.Ҫ������ & "," & Me.Ҫ�س��� & "," & _
            Me.Ҫ��С�� & ",'" & Me.Ҫ�ص�λ & "'," & Me.Ҫ�ر�ʾ & "," & Me.������̬ & ",'" & Me.Ҫ��ֵ�� & "')"
    Case cprET_ȫ��ʾ���༭
        gstrSQL = "Zl_������������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",4," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.�������� & "',NULL,'" & Replace(Me.�����ı�, "'", "' || chr(39) || '") & "'," & IIf(Me.�Ƿ���, 1, 0) & _
            ",NULL,NULL,NULL,NULL," & _
            IIf(Me.����Ҫ��ID = 0, "NULL", Me.����Ҫ��ID) & "," & Me.�滻�� & ",'" & Me.Ҫ������ & "'," & Me.Ҫ������ & "," & Me.Ҫ�س��� & "," & _
            Me.Ҫ��С�� & ",'" & Me.Ҫ�ص�λ & "'," & Me.Ҫ�ر�ʾ & "," & Me.������̬ & ",'" & Me.Ҫ��ֵ�� & "')"
    Case cprET_�������༭, cprET_���������
        '���� ��ʼ�棬��ֹ�� �ֶΣ�������ͬ
        gstrSQL = "ZL_���Ӳ�������_Update(" & Me.ID & "," & Me.�ļ�ID & "," & Me.��ʼ�� & "," & Me.��ֹ�� & "," & IIf(Me.��ID = 0, "NULL", Me.��ID) & "," & _
            Me.������� & ",4," & Me.Key & "," & IIf(Me.��������, 1, 0) & ",'" & Me.�������� & "',NULL,'" & Replace(Me.�����ı�, "'", "' || chr(39) || '") & "'," & IIf(Me.�Ƿ���, 1, 0) & _
            ",NULL,NULL,NULL,NULL," & _
            IIf(Me.����Ҫ��ID = 0, "NULL", Me.����Ҫ��ID) & "," & Me.�滻�� & ",'" & Me.Ҫ������ & "'," & Me.Ҫ������ & "," & Me.Ҫ�س��� & "," & _
            Me.Ҫ��С�� & ",'" & Me.Ҫ�ص�λ & "'," & Me.Ҫ�ر�ʾ & "," & Me.������̬ & ",'" & Me.Ҫ��ֵ�� & "')"
    End Select
    ArraySQL(lngCount) = gstrSQL
    
    GetSaveSQL = True
End Function

'################################################################################################################
'## ���ܣ�  ѡ�и�����Ҫ��
'##
'## ������  edtThis         :��ǰ�ı༭���ؼ�
'################################################################################################################
Public Function Selected(ByRef edtThis As Object) As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    On Error Resume Next
    bFinded = FindKey(edtThis, "E", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Range(lKSE, lKES).Selected
        End With
        Selected = True
    Else
        Selected = False
    End If
End Function

Public Sub CheckFonts(fmtFont As cFontFormat)
Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        If Screen.Fonts(i) = fmtFont.Name Then
            Exit Sub
        End If
    Next i
    fmtFont.Name = "����"
End Sub