'ȱʡ����ֵ:
LocateCol = 1 '��λ��(�������ĳ�б��û�ѡ�У�����û�ѡ����ؼ����Զ���λ������
CmdEnable = False '�ÿؼ��еİ�ťEnable����
CmdVisible = False '�ÿؼ��еİ�ťVisible����
CobEnable = False '�ÿؼ��еİ�ťEnable����
CobVisible = False '�ÿؼ��еİ�ťVisible����
TxtEnable = False '�ÿؼ��еİ�ťEnable����
TxtVisible = False '�ÿؼ��еİ�ťVisible����
MonVisible = False '�ÿؼ��еİ�ťVisible���� 
MonEnable = False '�ÿؼ��еİ�ťEnable����
    
****���øÿؼ�����ֵ***** 'ȱʡΪ0
    '�����ֵΪ1������ʾ��ť
    '�����ֵΪ2������ʾ��ť������ʾ���ڿؼ�
    '�����ֵΪ3������ʾ������
    '�����ֵΪ4������ʾ�ı���
    '�����ֵΪ5��������ѡ��,��ѡ����λ����λ��
    '�����ֵΪ0�����û�����ѡ��,�����ܸ���
    '�����ֵΪ����ֵ�����û�����ѡ��

�����Ϣ������ϸ�Ķ�------���ݿؼ�.Doc

����Դ���룺

Private Sub Form_Load()
    msf.Cols = 8

    msf.Clear
    msf.active=true

    msf.AddItem "Ҽ"
    msf.AddItem "��"
    msf.AddItem "��"
    msf.AddItem "��"
    msf.AddItem "��"
    msf.AddItem "½"
    msf.AddItem "��"
    msf.AddItem "��"
    msf.AddItem "��"
    msf.AddItem "ʰ"
    
    msf.TextMatrix(0, 0) = "��1��"
    msf.TextMatrix(0, 1) = "��2��"
    msf.TextMatrix(0, 2) = "��3��"
    msf.TextMatrix(0, 3) = "��4��"
    msf.TextMatrix(0, 4) = "��5��"
    msf.TextMatrix(0, 5) = "��6��"
    msf.TextMatrix(0, 6) = "��7��"
    msf.TextMatrix(0, 7) = "��8��"
    
    msf.ColData(0) = 1
    msf.ColData(1) = 0
    msf.ColData(2) = 2
    msf.ColData(3) = 3
    msf.ColData(4) = 4
    msf.ColData(5) = 5
    msf.ColData(6) = 4
    msf.ColData(7) = 5
    
    msf.ColAlignment(0) = 1
    msf.ColAlignment(1) = 1
    msf.ColAlignment(2) = 1
    msf.ColAlignment(3) = 1
    msf.ColAlignment(4) = 7
    msf.ColAlignment(5) = 4
    msf.ColAlignment(6) = 7
    msf.ColAlignment(7) = 4
    
    msf.TextMatrix(1, 5) = "����ѡ��"
    msf.TextMatrix(1, 7) = "����ѡ��"
    
    msf.MaxDate = "9999-12-31"
    msf.MinDate = "1901-01-01"

    Dim Lop As Integer
    msf.Row = 0
    For Lop = 0 To msf.Cols - 1
        msf.Col = Lop
        msf.CellAlignment = 4
    Next
    msf.Row = 1
End Sub

Private Sub msf_cmdselectclick()
    if msf.col=0 then
        MsgBox "лл��ʹ���������������˾�������", vbInformation, "����"
        msf.TextMatrix(msf.Row, 1) = "Thanks��"
        msf.Col = msf.LocateCol
        msf.CmdVisible = False
    endif
End Sub
