
--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�����б�
(
  ����id_In In ����ҽ����¼.����id%Type,
  ��ҳid_In In ����ҽ����¼.��ҳid%Type
) Return Varchar2 Is
  Pragma Autonomous_Transaction;
  
  TYPE C_REPORT_LIST IS REF CURSOR;
  C_REPORT_ITEM C_REPORT_LIST;

  v_Return Varchar2(4000);  
  v_Sql    Varchar2(4000);
  v_Temp   Varchar2(2000);
  n_Count  Number;
  
  n_ITEM_Id       Varchar2(64);
  n_ITEM_YZID     Number(18);
  v_ITEM_GP       Varchar2(64);
  v_ITEM_MC       Varchar2(1024);
  n_ITEM_BGLX     Number(18);
  v_ITEM_BGR      Varchar2(64);
  v_ITEM_BGSJ     Varchar2(64);  

Begin
  
    Select Count(*) Into n_Count From user_tables Where table_name =Upper('zlTempReportList');
    
    if n_Count > 0 then
      v_sql := 'Truncate Table zlTempReportList';
      Execute Immediate v_sql;
      Commit;
    Else
      v_sql := 'Create Global Temporary Table zlTempReportList(
               ID Varchar2(64),   
               YZID Number(18),             
               GP Number(1),
               MC Varchar2(1024),
               BGLX Number(1),
               BGR Varchar2(64),
               BGSJ Date
              ) On Commit Preserve Rows'; 
                    
      Execute Immediate v_sql;
    End if;

    v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ) 
                           Select b.����id || '''' As ID, a.Id As YZID, Decode(d.���uid, Null, 0, 1) As GP, a.ҽ������ As MC, 0 As BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                           From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                           Where a.Id = b.ҽ��id And b.����id = c.Id And a.������� = ''D'' And ���id Is Null And B.RISID Is Null And
                           c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                           a.����id = :1 And nvl(a.��ҳid,0) = :2';  
    Begin                   
        Execute Immediate v_Sql Using ����id_In,��ҳid_In;
    Exception
      When Others Then
        Begin
          v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ) 
                                 Select b.����id || '''' As ID, a.Id As YZID, Decode(d.���uid, Null, 0, 1) As GP, a.ҽ������ As MC, 0 As BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                                 From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                                 Where a.Id = b.ҽ��id And b.����id = c.Id And a.������� = ''D'' And ���id Is Null And
                                 c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                                 a.����id = :1 And nvl(a.��ҳid,0) = :2';  
          Execute Immediate v_Sql Using ����id_In,��ҳid_In;
        Exception
          When Others Then Null;
        end; 
    End;
    
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ) 
                          Select b.��鱨��id || '''' As ID, a.Id As YZID, Decode(d.���uid, Null, 0, 1) As GP, a.ҽ������ As MC, 1 as BGLX, c.���༭�� As BGR, c.������ʱ�� As BGSJ
                          From ����ҽ����¼ A, ����ҽ������ B, Ӱ�񱨸��¼ C, Ӱ�����¼ D
                          where a.Id = b.ҽ��id And b.��鱨��id = c.Id And a.������� = ''D'' And ���id Is Null And
                          c.������ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                          a.����id = :1 And nvl(a.��ҳid,0) = :2';
    Begin
        Execute Immediate v_Sql Using ����id_In,��ҳid_In;
    Exception
      When Others Then Null;
    End;
     
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ)
                          Select b.RISID || '''' As ID, a.Id As YZID, 2 As GP, a.ҽ������ As MC, 2 as BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                          From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                          Where a.Id = b.ҽ��id And b.����ID = c.Id And a.������� = ''D'' And ���id Is Null And B.RISID Is Not Null And
                          c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                          a.����id = :1 And nvl(a.��ҳid,0) = :2';     
    Begin
        Execute Immediate v_Sql Using ����id_In,��ҳid_In; 
    Exception
      When Others Then Null;
    End;
    
    Commit;
    
    v_Sql := 'Select Id, YZID, GP, MC, BGLX, BGR, To_Char(BGSJ,''yyyy-mm-dd hh24:mi:ss'') As BGSJ  From zlTempReportList Order by BGSJ';

    Open C_REPORT_ITEM For v_Sql;
    Loop
      Fetch C_REPORT_ITEM INTO n_ITEM_ID, n_ITEM_YZID, v_ITEM_GP, v_ITEM_MC, n_ITEM_BGLX, v_ITEM_BGR, v_ITEM_BGSJ;
      Exit When C_REPORT_ITEM%NotFound;
      
      v_Temp := '<FILE>' || 
                      '<ID>' || n_ITEM_ID || '</ID>' || 
                      '<YZID>' || n_ITEM_YZID || '</YZID>' ||
                      '<GP>' || v_ITEM_GP || '</GP>' || 
                      '<MC>' || v_ITEM_MC || '</MC>' || 
                      '<BGLX>' || n_ITEM_BGLX || '</BGLX>' || 
                      '<BGR>' || v_ITEM_BGR || '</BGR>' ||
                      '<BGSJ>' || v_ITEM_BGSJ || '</BGSJ>' || 
             '</FILE>';

      v_Return := v_Return || v_Temp;
    End Loop;
    Close C_REPORT_ITEM;

    If v_Return <> ' ' Then
      v_Return := '<FILELIST>' || v_Return || '</FILELIST>';
    End If;

    Return v_Return;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�����б�;
/


--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�����б�Ex
(
  ҽ��ID_In In ����ҽ����¼.id%Type
) Return Varchar2 Is
  Pragma Autonomous_Transaction;
  
  TYPE C_REPORT_LIST IS REF CURSOR;
  C_REPORT_ITEM C_REPORT_LIST;

  v_Return Varchar2(4000);  
  v_Sql    Varchar2(4000);
  v_Temp   Varchar2(2000);
  n_Count  Number;
  
  n_ITEM_Id       Varchar2(64);
  n_ITEM_YZID     Number(18);
  v_ITEM_YZNR     Varchar2(1024);
  v_ITEM_MC       Varchar2(60);
  n_ITEM_BGLX     Number(18);
  v_ITEM_BGR      Varchar2(64);
  v_ITEM_BGSJ     Varchar2(64);  

Begin
  
    Select Count(*) Into n_Count From user_tables Where table_name =Upper('zlTempReportList');
    
    if n_Count > 0 then
      v_sql := 'Truncate Table zlTempReportList';
      Execute Immediate v_sql;
      Commit;
    Else
      v_sql := 'Create Global Temporary Table zlTempReportList(
               ID Varchar2(64),   
               YZID Number(18),             
               YZNR Varchar2(1024),
               MC Varchar2(60),
               BGLX Number(1),
               BGR Varchar2(64),
               BGSJ Date
              ) On Commit Preserve Rows'; 
                    
      Execute Immediate v_sql;
    End if;

    v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ) 
                           Select b.����id || '''' As ID, a.Id As YZID, a.ҽ������ As YXNR, c.�������� as MC, 0 As BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                           From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                           Where a.Id = b.ҽ��id And b.����id = c.Id And a.������� = ''D'' And ���id Is Null And B.RISID Is Null And
                           c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                           a.ID = :1';  
    Begin                   
        Execute Immediate v_Sql Using ҽ��ID_In;
    Exception
      When Others Then
        Begin
          v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ) 
                                 Select b.����id || '''' As ID, a.Id As YZID, a.ҽ������ As YZNR, c.�������� As MC, 0 As BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                                 From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                                 Where a.Id = b.ҽ��id And b.����id = c.Id And a.������� = ''D'' And ���id Is Null And
                                 c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                                 a.id = :1';  
          Execute Immediate v_Sql Using ҽ��ID_In;
        Exception
          When Others Then Null;
        end; 
    End;
    
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ) 
                          Select b.��鱨��id || '''' As ID, a.Id As YZID, a.ҽ������ As YZNR, c.�ĵ����� As MC, 1 as BGLX, c.���༭�� As BGR, c.������ʱ�� As BGSJ
                          From ����ҽ����¼ A, ����ҽ������ B, Ӱ�񱨸��¼ C, Ӱ�����¼ D
                          where a.Id = b.ҽ��id And b.��鱨��id = c.Id And a.������� = ''D'' And ���id Is Null And
                          c.������ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                          a.id = :1';
    Begin
        Execute Immediate v_Sql Using ҽ��ID_In;
    Exception
      When Others Then Null;
    End;
     
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ)
                          Select b.RISID || '''' As ID, a.Id As YZID, a.ҽ������ As YZNR, c.�������� As MC, 2 as BGLX, c.������ As BGR, c.���ʱ�� As BGSJ
                          From ����ҽ����¼ A, ����ҽ������ B, ���Ӳ�����¼ C, Ӱ�����¼ D
                          Where a.Id = b.ҽ��id And b.����ID = c.Id And a.������� = ''D'' And ���id Is Null And B.RISID Is Not Null And
                          c.���ʱ�� Is Not Null And a.Id = d.ҽ��id(+) And a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And
                          a.id = :1';     
    Begin
        Execute Immediate v_Sql Using ҽ��ID_In; 
    Exception
      When Others Then Null;
    End;
    
    Commit;
    
    v_Sql := 'Select Id, YZID, YZNR, MC, BGLX, BGR, To_Char(BGSJ,''yyyy-mm-dd hh24:mi:ss'') As BGSJ  From zlTempReportList Order by BGSJ';

    Open C_REPORT_ITEM For v_Sql;
    Loop
      Fetch C_REPORT_ITEM INTO n_ITEM_ID, n_ITEM_YZID, v_ITEM_YZNR, v_ITEM_MC, n_ITEM_BGLX, v_ITEM_BGR, v_ITEM_BGSJ;
      Exit When C_REPORT_ITEM%NotFound;
      
      v_Temp := '<FILE>' || 
                      '<ID>' || n_ITEM_ID || '</ID>' || 
                      '<YZID>' || n_ITEM_YZID || '</YZID>' ||
                      '<YZNR>' || v_ITEM_YZNR || '</YZNR>' || 
                      '<MC>' || v_ITEM_MC || '</MC>' || 
                      '<BGLX>' || n_ITEM_BGLX || '</BGLX>' || 
                      '<BGR>' || v_ITEM_BGR || '</BGR>' ||
                      '<BGSJ>' || v_ITEM_BGSJ || '</BGSJ>' || 
             '</FILE>';

      v_Return := v_Return || v_Temp;
    End Loop;
    Close C_REPORT_ITEM;

    If v_Return <> ' ' Then
      v_Return := '<FILELIST>' || v_Return || '</FILELIST>';
    End If;

    Return v_Return;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�����б�Ex;
/


--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�ĵ����
(
  ����id_In In ����ҽ������.��鱨��ID%Type
) Return Varchar2 Is
  v_������� Varchar2(1000);
  v_������� Varchar2(100);

  x_Content xmltype;
  n_NodeNum number(2);
  Xcdom            Xmldom.Domdocument;
  Section_List     Xmldom.Domnodelist;
Begin
    v_������� := '';

    Select b.�������� Into x_Content From ����ҽ������ a, Ӱ�񱨸��¼ b Where a.��鱨��id=b.id And  a.��鱨��id = ����id_In;

    Xcdom         := Xmldom.Newdomdocument(x_Content);
    Section_List  := Xmldom.Getelementsbytagname(Xcdom, 'section');
    n_NodeNum     := Xmldom.Getlength(Section_List);

    For i in 0..n_NodeNum-1 Loop
      v_������� := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Section_List, i)), 'title');

      If Nvl(v_�������,' ') != ' ' Then
        v_������� := v_������� || '<split>' || v_�������;
      End If;
    End Loop;
    
    Return(Substr(v_�������, 8));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�ĵ����;
/


CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�������
(
  ����id_In In ����ҽ������.����id%Type
) Return Varchar2 Is
  v_������� Varchar2(1000);

  Cursor c_������� Is
    Select Distinct a.�����ı�
    From ���Ӳ������� A, ���Ӳ������� B, ����ҽ������ C
    Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = c.����id And c.����id = ����id_In;
Begin
  For Row_Cols In c_������� Loop
    v_������� := '<split>' || Row_Cols.�����ı� || v_�������;
  End Loop;

  Return(Substr(v_�������, 8));

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�������;
/

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�������
(
  ����id_In In Varchar2,
  ������Դ_In In Number
) Return Varchar2 Is
  v_������� Varchar2(2000);
  n_����ID Number(18);
  
  v_Sql Varchar2(100);
Begin
  If ������Դ_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_��ȡ�ĵ����(:1)  From Dual';  
    Begin                   
        Execute Immediate v_Sql Into v_������� Using ����id_In ;
    Exception
      When Others Then v_������� := '';          
    End;
  Else
    n_����ID := To_Number(����id_In);
      
    If ������Դ_In = 2 Then
      v_Sql := 'Select ����ID From ����ҽ������ Where RISID=:1';
      Execute Immediate v_Sql Into n_����ID Using ����id_In ;
    End If;
      
    v_Sql := 'Select Zlpub_Pacs_��ȡ�������(:1)  From Dual';  
    Begin                   
        Execute Immediate v_Sql Into v_������� Using n_����ID ;
    Exception
      When Others Then v_������� := '';          
    End;
  End If;

  Return v_�������;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�������;
/


--******************************************************************************************


CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�ĵ��ı�
(
  Ids_In   In Varchar2
)Return XmlType Is
  Docxml XmlType;
  
  File_Id Varchar2(32);
  n_Adviceid Number(18);
  
  x_Content    Xmltype;
  Section_Node Xmldom.Domnode;
  Element_Node Xmldom.Domnode;
  Xcdom        Xmldom.Domdocument;
  Node_List    Xmldom.Domnodelist;
  Section_List Xmldom.Domnodelist;
  
  n_Count Number(1);
  --��Ǳ���
  
  n_Len     Number(3);
  n_Width   Number(4);
  n_Height  Number(4);
  v_Id      Varchar2(100);
  v_Title   Varchar2(100);
  v_Newline Varchar2(2);
  v_Text    Varchar2(4000);
  v_Name    Varchar2(100);
  v_Type    Varchar2(20);
Begin
  Select Xmltype('<?xml version="1.0" encoding="' || Value || '"?><ZlEPR></ZlEPR>')
  Into Docxml
  From Nls_Database_Parameters
  Where Parameter = 'NLS_CHARACTERSET';
  
  For J In 1 .. 1000 Loop
    File_Id := 0;
    Select Zl_Eprsplit(Ids_In, '|', J) Into File_Id From Dual;
    
    If File_Id Is Null Then
      Exit;
    End If;      
    
    --��ʼĳ���ļ���ȡ
    Begin
      Select a.ҽ��id,
             Appendchildxml(Docxml, '/ZlEPR',
                             Xmlelement("Document",
                                         Xmlattributes(b.���� As "����", b.����id As "����ID", b.��ҳid As "��ҳID", a.�ĵ����� As "�ļ���",
                                                        Rawtohex(a.Id) As "�ļ�ID")))
      Into n_Adviceid, Docxml
      From Ӱ�񱨸��¼ A, ����ҽ����¼ B
      Where a.Id = Hextoraw(File_Id) And a.ҽ��id = b.Id;
    Exception
      --�������ļ�ID��Ч
      When Others Then
        Return Null;
    End;
    
    Select Insertchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]', 'Compend',
                           Xmlelement("Compend", Xmlattributes('0' As "ID", '����' As "Name")))
    Into Docxml
    From Dual;
      
    --��ʼ��ȡ����
    Select b.�������� Into x_Content From Ӱ�񱨸��¼ B Where b.Id || '' = File_Id;
      
    Xcdom := Xmldom.Newdomdocument(x_Content);
      
    Section_List := Xmldom.Getelementsbytagname(Xcdom, 'zlxml');
    Section_Node := Xmldom.Item(Section_List, 0);
    Node_List    := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Section_Node), '*');
    n_Len        := Xmldom.Getlength(Node_List);
      
    For I In 0 .. n_Len - 1 Loop
      Element_Node := Xmldom.Item(Node_List, I);
        
      v_Name    := Xmldom.Getnodename(Element_Node);
      v_Newline := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'br');
        
      If v_Newline Is Null Then
        v_Newline := '1';
      End If;
        
      If v_Name = 'section' Then
        --���
        v_Title := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'title');
        v_Id    := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'sid');
          
        Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]',
                               Xmlelement("Compend", Xmlattributes(v_Title As "Name", v_Id As "ID")))
        Into Docxml
        From Dual;
      Elsif v_Name = 'utext' Then
        --�ı�
        v_Text := LTrim(LTrim(Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)), ':'), '��');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Text", Xmlattributes(v_Newline As "NewLine"), v_Text))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Text", Xmlattributes(v_Newline As "NewLine"), v_Text))
          Into Docxml
          From Dual;
        End If;
      Elsif v_Name = 'element' Then
        --Ҫ��
        v_Title := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'title');
        v_Text  := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'value') ||
                   Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'unit');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Element", Xmlattributes(v_Title As "Name", v_Newline As "NewLine"),
                                             v_Text))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Element", Xmlattributes(v_Title As "Name", v_Newline As "NewLine"),
                                             v_Text))
          Into Docxml
          From Dual;
        End If;
      Elsif v_Name = 'image' Then
        --ͼƬ
        n_Width  := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'width');
        n_Height := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'height');
        v_Name   := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'key');
        v_Type   := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'class');
          
        If Nvl(v_Name, ' ') <> ' ' Then
          If Nvl(v_Id, ' ') = ' ' Then
            Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                                   Xmlelement("Picture",
                                               Xmlattributes(n_Width As "OrigWidth", n_Height As "OrigHeight",
                                                              n_Width As "ShowWidth", n_Height As "ShowHeight",
                                                              v_Name As "PicName", n_Adviceid As "AdviceID",
                                                              v_Type As "Type")))
            Into Docxml
            From Dual;
          Else
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                   Xmlelement("Picture",
                                               Xmlattributes(n_Width As "OrigWidth", n_Height As "OrigHeight",
                                                              n_Width As "ShowWidth", n_Height As "ShowHeight",
                                                              v_Name As "PicName", n_Adviceid As "AdviceID",
                                                              v_Type As "Type")))
            Into Docxml
            From Dual;
          End If;
        End If;
          
      Elsif v_Name = 'signature' Then
        --ǩ��
        v_Text := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'displayinfo');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Sign", Xmlattributes(v_Newline As "NewLine"), Zl_Eprsplit(v_Text, ';', 1)))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Sign", Xmlattributes(v_Newline As "NewLine"), Zl_Eprsplit(v_Text, ';', 1)))
          Into Docxml
          From Dual;
        End If;
      End If;
    End Loop;
      
    For Aa In (Select '/' || a.FtpĿ¼ || '/ReportImages/' || To_Char(b.����ʱ��, 'YYYYMMDD') || '/' || b.Id || '/' As v_Ftppath
               From Ӱ���豸Ŀ¼ A, Ӱ�񱨸��¼ B
               Where a.�豸�� = b.�豸�� And b.Id = File_Id) Loop
        
      Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                             Xmlelement("FtpPath", Xmlattributes(v_Newline As "NewLine"), Aa.v_Ftppath))
      Into Docxml
      From Dual;
    End Loop;
  End Loop;
  
  Return Docxml;
End Zlpub_Pacs_��ȡ�ĵ��ı�;
/
 

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�����ı�
(
  Ids_In   In Varchar2,
  From_In Number
)Return XmlType Is
  Docxml XmlType;
  
  File_Id Varchar2(32);
  n_Adviceid Number(18);
  
  v_Sql        Varchar2(1000);
  
  --��Ǳ���
  v_Mark     Varchar2(500);
  v_Marks    Varchar2(2500);
  Makxml     Xmltype;
  Maksxml    Xmltype;
  v_Ftppath  Varchar2(200);
  
  v_Newline Varchar2(2);
Begin
  Select Xmltype('<?xml version="1.0" encoding="' || Value || '"?><ZlEPR></ZlEPR>')
  Into Docxml
  From Nls_Database_Parameters
  Where Parameter = 'NLS_CHARACTERSET';
  
  For J In 1 .. 1000 Loop
    File_Id := 0;
    Select Zl_Eprsplit(Ids_In, '|', J) Into File_Id From Dual;
    
    If File_Id Is Null Then
      Exit;
    End If;  
    
    If From_In = 2 Then
       --RIS����
       v_Sql := 'Select ����Id From ����ҽ������ Where RISID = :1';
       Execute Immediate v_Sql Into File_Id Using File_Id;
    End If;
        
    --��ʼĳ�������ļ���ȡ
    Begin
      Select Appendchildxml(Docxml, '/ZlEPR',
                             Xmlelement("Document",
                                         Xmlattributes(b.���� As "����", a.����id As "����ID", a.��ҳid As "��ҳID", a.�������� As "�ļ���",
                                                        a.Id As "�ļ�ID")))
      Into Docxml
      From ���Ӳ�����¼ A, ������Ϣ B
      Where a.Id = File_Id And a.�༭��ʽ = 0 And a.����id = b.����id;
          
      Select ҽ��id Into n_Adviceid From ����ҽ������ Where ����id = File_Id;
    Exception
      --�����Ĳ����ļ�ID��Ч
      When Others Then Return Null;
    End;
        
    Select Insertchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]', 'Compend',
                           Xmlelement("Compend", Xmlattributes('0' As "ID", '����' As "Name")))
    Into Docxml
    From Dual;
      
    Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                           Xmlelement("Text", Xmlattributes(Nvl(Null, 0) As "NewLine"), '�����ı�'))
    Into Docxml
    From Dual;
      
    For Rs In (Select ID, ��id, �������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ҫ������
               From (Select ID, ��id, �������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ҫ������
                      From ���Ӳ�������
                      Where �ļ�id = File_Id And ������� > 0 And ������� <> ID And ��ֹ�� = 0)
               Start With ��id Is Null
               Connect By Prior ID = ��id
               Order Siblings By �������, �����д�) Loop
      If Rs.�������� = 1 Then
        --���
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']',
                                 Xmlelement("Compend", Xmlattributes(Rs.�����ı� As "Name", Rs.Id As "ID")))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Compend", Xmlattributes(Rs.�����ı� As "Name", Rs.Id As "ID")))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.�������� = 2 Then
        --�ı�
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Text", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"), Rs.�����ı�))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Text", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"), Rs.�����ı�))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.�������� = 3 Then
        --���
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Table",
                                             Xmlattributes(Zl_Eprsplit(Rs.��������, ';', 1) As "Rows",
                                                            Zl_Eprsplit(Rs.��������, ';', 2) As "Cols",
                                                            Zl_Eprsplit(Rs.��������, ';', 3) As "Width",
                                                            Zl_Eprsplit(Rs.��������, ';', 4) As "Height",
                                                            Zl_Eprsplit(Rs.��������, ';', 5) As "ColWidthString",
                                                            Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs.Id As "ID"), Rs.�����ı�))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Table",
                                             Xmlattributes(Zl_Eprsplit(Rs.��������, ';', 1) As "Rows",
                                                            Zl_Eprsplit(Rs.��������, ';', 2) As "Cols",
                                                            Zl_Eprsplit(Rs.��������, ';', 3) As "Width",
                                                            Zl_Eprsplit(Rs.��������, ';', 4) As "Height",
                                                            Zl_Eprsplit(Rs.��������, ';', 5) As "ColWidthString",
                                                            Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs.Id As "ID"), Rs.�����ı�))
          Into Docxml
          From Dual;
        End If;
          
        ---�Ա��ĵ�Ԫ��������
        For Rs_Cell In (Select ID, ��id, �������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ҫ������
                        From ���Ӳ�������
                        Where �ļ�id = File_Id And ��id = Rs.Id And ��ֹ�� = 0
                        Order By �����д�, ID) Loop
          If Rs_Cell.�������� = 2 Or Rs_Cell.�������� = 4 Then
            If Zl_Eprsplit(Rs_Cell.��������, '|', 26) Is Null Then
              --������ʷ����
              Select Appendchildxml(Docxml,
                                     '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                     Xmlelement("Cell",
                                                 Xmlattributes(Zl_Eprsplit(Rs_Cell.��������, '|', 2) As "Row",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 3) As "Col",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 2) || '_' ||
                                                                 Zl_Eprsplit(Rs_Cell.��������, '|', 3) As "Row_Col",
                                                                Decode(Rs_Cell.��������, 2, 0, 4, 1) As "Type",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 5) As "Width",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 6) As "Height",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 4) As "MergeNo",
                                                                Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs_Cell.Id As "ID"),
                                                 Rs_Cell.�����ı�))
              Into Docxml
              From Dual;
            Else
              Select Appendchildxml(Docxml,
                                     '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                     Xmlelement("Cell",
                                                 Xmlattributes(Zl_Eprsplit(Rs_Cell.��������, '|', 3) As "Row",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 4) As "Col",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 3) || '_' ||
                                                                 Zl_Eprsplit(Rs_Cell.��������, '|', 4) As "Row_Col",
                                                                Decode(Rs_Cell.��������, 2, 0, 4, 1) As "Type",
                                                                Zl_Eprsplit(Rs_Cell.��������, '', 6) As "Width",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 7) As "Height",
                                                                Zl_Eprsplit(Rs_Cell.��������, '|', 5) As "MergeNo",
                                                                Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs_Cell.Id As "ID"),
                                                 Rs_Cell.�����ı�))
              Into Docxml
              From Dual;
            End If;
          Elsif Rs_Cell.�������� = 5 And Zl_Eprsplit(Rs_Cell.��������, ';', 1) = 2 Then
            --��Ԫ��ͼ��Webserviceֱ�Ӷ�ȡBLOB֮��ֱ��д�ļ�������ٶ�
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                   Xmlelement("Picture",
                                               Xmlattributes(Zl_Eprsplit(Rs_Cell.��������, ';', 2) As "Row",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 3) As "Col",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 8) As "OrigWidth",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 9) As "OrigHeight",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 6) As "ShowWidth",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 7) As "ShowHeight",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.��������, ';', 12), ' ') As "PicName",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.��������, ';', 13), '0') As "AdviceID",
                                                              Rs_Cell.Id As "ID"), Rs_Cell.Id))
            Into Docxml
            From Dual;
              
            If Nvl(n_Adviceid, 0) = 0 Then
              n_Adviceid := Nvl(Zl_Eprsplit(Rs_Cell.��������, ';', 13), '0');
            End If;
          Elsif Rs_Cell.�������� = 5 And Zl_Eprsplit(Rs_Cell.��������, ';', 1) <> 2 Then
            --��Ԫ��ͼ��Webserviceֱ�Ӷ�ȡBLOB֮��ֱ��д�ļ�������ٶ�
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id ||
                                    ']/Cell[@Row_Col="' || Zl_Eprsplit(Rs_Cell.��������, ';', 2) || '_' ||
                                    Zl_Eprsplit(Rs_Cell.��������, ';', 3) || '"]',
                                   Xmlelement("Picture",
                                               Xmlattributes(Zl_Eprsplit(Rs_Cell.��������, ';', 2) As "Row",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 3) As "Col",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 8) As "OrigWidth",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 9) As "OrigHeight",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 6) As "ShowWidth",
                                                              Zl_Eprsplit(Rs_Cell.��������, ';', 7) As "ShowHeight",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.��������, ';', 12), ' ') As "PicName",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.��������, ';', 13), '0') As "AdviceID",
                                                              Rs_Cell.Id As "ID"), Rs_Cell.Id))
            Into Docxml
            From Dual;
            --��������ӽڵ㼯
            v_Mark  := '';
            Makxml  := Null;
            Maksxml := Null;
            For Rs_Mark In (Select ID, ��id, �����ı�, �����д�
                            From ���Ӳ�������
                            Where ��id = Rs_Cell.Id
                            Order By �����д�) Loop
              v_Marks := v_Mark || Rs_Mark.�����ı�;
              v_Marks := Replace(v_Marks, '||', '^');
              For I In 1 .. 100 Loop
                v_Mark := Zl_Eprsplit(v_Marks, '^', I);
                If Zl_Eprsplit(v_Mark, '|', 15) Is Null Then
                  --���һ�������Ϣ��ȫ��������һ����
                  Exit;
                Else
                  Select Xmlelement("Mark",
                                     Xmlforest(Zl_Eprsplit(v_Mark, '|', 2) As "����",
                                                Zl_Eprsplit(v_Mark, '|', 3) As "����", Zl_Eprsplit(v_Mark, '|', 4) As "�㼯",
                                                Zl_Eprsplit(v_Mark, '|', 5) As "X1", Zl_Eprsplit(v_Mark, '|', 6) As "Y1",
                                                Zl_Eprsplit(v_Mark, '|', 7) As "X2", Zl_Eprsplit(v_Mark, '|', 8) As "Y2",
                                                Zl_Eprsplit(v_Mark, '|', 9) As "���ɫ",
                                                Zl_Eprsplit(v_Mark, '|', 10) As "��䷽ʽ",
                                                Zl_Eprsplit(v_Mark, '|', 11) As "����ɫ",
                                                Zl_Eprsplit(v_Mark, '|', 12) As "����ɫ",
                                                Zl_Eprsplit(v_Mark, '|', 13) As "����",
                                                Zl_Eprsplit(v_Mark, '|', 14) As "�߿�",
                                                Zl_Eprsplit(v_Mark, '|', 15) As "����"))
                  Into Makxml
                  From Dual;
                  Select Xmlconcat(Maksxml, Makxml) Into Maksxml From Dual;
                End If;
              End Loop;
            End Loop;
            --��Picture�������ӽڵ�
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Picture[@ID=' || Rs_Cell.Id || ']',
                                   Maksxml)
            Into Docxml
            From Dual;
          End If;
        End Loop;
      Elsif Rs.�������� = 4 Then
        --Ҫ��
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Element", Xmlattributes(Rs.Ҫ������ As "Name", Nvl(Rs.�Ƿ���, 0) As "NewLine"),
                                             Rs.�����ı�))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Element", Xmlattributes(Rs.Ҫ������ As "Name", Nvl(Rs.�Ƿ���, 0) As "NewLine"),
                                             Rs.�����ı�))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.�������� = 5 And Nvl(Rs.�����д�, 0) = 0 Then
        --ͼƬ��Webserviceֱ�Ӷ�ȡBLOB֮��ֱ��д�ļ�������ٶ�
 
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Picture",
                                             Xmlattributes(Zl_Eprsplit(Rs.��������, ';', 8) As "OrigWidth",
                                                            Zl_Eprsplit(Rs.��������, ';', 9) As "OrigHeight",
                                                            Zl_Eprsplit(Rs.��������, ';', 6) As "ShowWidth",
                                                            Zl_Eprsplit(Rs.��������, ';', 7) As "ShowHeight",
                                                            Nvl(Zl_Eprsplit(Rs.��������, ';', 12), ' ') As "PicName",
                                                            Nvl(Zl_Eprsplit(Rs.��������, ';', 13), '0') As "AdviceID",
                                                            Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs.Id As "ID"), Rs.Id))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Picture",
                                             Xmlattributes(Zl_Eprsplit(Rs.��������, ';', 8) As "OrigWidth",
                                                            Zl_Eprsplit(Rs.��������, ';', 9) As "OrigHeight",
                                                            Zl_Eprsplit(Rs.��������, ';', 6) As "ShowWidth",
                                                            Zl_Eprsplit(Rs.��������, ';', 7) As "ShowHeight",
                                                            Nvl(Zl_Eprsplit(Rs.��������, ';', 12), ' ') As "PicName",
                                                            Nvl(Zl_Eprsplit(Rs.��������, ';', 13), '0') As "AdviceID",
                                                            Nvl(Rs.�Ƿ���, 0) As "NewLine", Rs.Id As "ID"), Rs.Id))
          Into Docxml
          From Dual;
        End If;
        --��������ӽڵ㼯
        v_Mark  := '';
        Makxml  := Null;
        Maksxml := Null;
        For Rs_Mark In (Select ID, ��id, �����ı�, �����д� From ���Ӳ������� Where ��id = Rs.Id Order By �����д�) Loop
          v_Marks := v_Mark || Rs_Mark.�����ı�;
          v_Marks := Replace(v_Marks, '||', '^');
          For I In 1 .. 100 Loop
            v_Mark := Zl_Eprsplit(v_Marks, '^', I);
            If Zl_Eprsplit(v_Mark, '|', 15) Is Null Then
              --���һ�������Ϣ��ȫ��������һ����
              Exit;
            Else
              Select Xmlelement("Mark",
                                 Xmlforest(Zl_Eprsplit(v_Mark, '|', 2) As "����", Zl_Eprsplit(v_Mark, '|', 3) As "����",
                                            Zl_Eprsplit(v_Mark, '|', 4) As "�㼯", Zl_Eprsplit(v_Mark, '|', 5) As "X1",
                                            Zl_Eprsplit(v_Mark, '|', 6) As "Y1", Zl_Eprsplit(v_Mark, '|', 7) As "X2",
                                            Zl_Eprsplit(v_Mark, '|', 8) As "Y2", Zl_Eprsplit(v_Mark, '|', 9) As "���ɫ",
                                            Zl_Eprsplit(v_Mark, '|', 10) As "��䷽ʽ",
                                            Zl_Eprsplit(v_Mark, '|', 11) As "����ɫ", Zl_Eprsplit(v_Mark, '|', 12) As "����ɫ",
                                            Zl_Eprsplit(v_Mark, '|', 13) As "����", Zl_Eprsplit(v_Mark, '|', 14) As "�߿�",
                                            Zl_Eprsplit(v_Mark, '|', 15) As "����"))
              Into Makxml
              From Dual;
              Select Xmlconcat(Maksxml, Makxml) Into Maksxml From Dual;
            End If;
          End Loop;
        End Loop;
        --��Picture�������ӽڵ�
        Select Appendchildxml(Docxml,
                               '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Picture[@ID=' || Rs.Id || ']',
                               Maksxml)
        Into Docxml
        From Dual;
      Elsif Rs.�������� = 7 Then
        --���
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Diagnosise", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"), Rs.�����ı�))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Diagnosise", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"), Rs.�����ı�))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.�������� = 8 Then
        --ǩ��
        If Rs.��id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Sign", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"),
                                             Zl_Eprsplit(Rs.�����ı�, ';', 1)))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@�ļ�ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.��id || ']',
                                 Xmlelement("Sign", Xmlattributes(Nvl(Rs.�Ƿ���, 0) As "NewLine"),
                                             Zl_Eprsplit(Rs.�����ı�, ';', 1)))
          Into Docxml
          From Dual;
        End If;
      End If;
    End Loop;
      
    For Aa In (Select A1.FtpĿ¼ || '/' || To_Char(l.��������, 'yyyymmdd') || '/' || l.���uid As v_Ftppath
               From Ӱ�����¼ L, Ӱ���豸Ŀ¼ A1
               Where l.λ��һ = A1.�豸��(+) And l.ҽ��id = n_Adviceid) Loop
        
      Select Appendchildxml(Docxml, '/ZlEPR/Document[@�ļ�ID="' || File_Id || '"]/Compend[@ID=0]',
                             Xmlelement("FtpPath", Xmlattributes(v_Newline As "NewLine"), Aa.v_Ftppath))
      Into Docxml
      From Dual;
    End Loop;
    
  End Loop;
  
  Return Docxml;  
End Zlpub_Pacs_��ȡ�����ı�;
/  

CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�����ı�
(
  Ids_In In Varchar2,
  From_In Number
) Return Xmltype Is
--Ids_In�������� '|' �ָ���ID������ʼ/��β�� '|'
  --���������Ĳ����ļ�ID����������XML������XMLType
  Docxml XmlType;
  v_Sql Varchar2(1000);
Begin
    
  If From_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_��ȡ�ĵ��ı�(:1) From Dual';
    Execute Immediate v_Sql Into Docxml Using Ids_In;
  Else
    v_Sql := 'Select Zlpub_Pacs_��ȡ�����ı�(:1, :2) From Dual';
    Execute Immediate v_Sql Into Docxml Using Ids_In, From_In;
  End If;

  Return Docxml;
Exception
  When Others Then
    Return Null;
End Zlpub_Pacs_��ȡ�����ı�;
/

--******************************************************************************************
CREATE OR REPLACE Function Zlpub_Pacs_��ȡ�ĵ�����
( 
  ����ID_In In Ӱ�񱨸��¼.ID%Type, 
  �������_In In Ӱ�񱨸��¼.������%Type 
) Return Varchar2 Is 
  x_Content        xmltype; 
  Xcdom            Xmldom.Domdocument; 
  Section_List     Xmldom.Domnodelist; 
  Section_Node     Xmldom.Domnode; 
  Node_List        Xmldom.Domnodelist; 
  n_Len            Number; 
  Element_Node     Xmldom.Domnode; 
  p_Node           Xmldom.Domnode; 
  Enum_Node        Xmldom.Domnode; 
  e_Node           Xmldom.Domnodelist; 
  c_Node           Xmldom.Domnode; 
  Enumeration_List Xmldom.Domnodelist; 
  Enumeration_Node Xmldom.Domnode; 
  Item_List        Xmldom.Domnodelist; 
  Item_Node        Xmldom.Domnode; 
  Item_Node1       Xmldom.Domnode; 
  v_Name           Varchar2(100); 
  v_Result         Varchar2(4000); 
  n_i              Number; 
  n_Num            Number; 
  n_j              Number; 
  n_Enum           Number; 
  v_Val            Varchar2(20); 
  v_Content        Varchar2(4000); 
  v_Eleid          Varchar2(50); 
  v_Multisel       Varchar2(10); 
Begin 
    v_Result := ''; 
    
    Select �������� Into x_Content From Ӱ�񱨸��¼ Where id = ����ID_In;

    Select Deletexml(x_Content, '//image') Into x_Content From Dual; 
 
    Xcdom := Xmldom.Newdomdocument(x_Content); 
 
    For Myrow In (Select Column_Value Name From Table(f_Str2list(�������_In))) Loop 
      n_i := -1; 
      --ѭ��������� 
      Section_List := Xmldom.Getelementsbytagname(Xcdom, 'section'); 
      n_Len        := Xmldom.Getlength(Section_List); 
 
      For I In 0 .. n_Len - 1 Loop 
        If Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Section_List, I)), 'title') = Myrow.Name Then 
          n_i := I; 
          Exit; 
        End If; 
      End Loop; 
 
      If n_i >= 0 Then 
        Section_Node := Xmldom.Item(Section_List, n_i); 
        Node_List    := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Section_Node), '*'); 
        n_Len        := Xmldom.Getlength(Node_List); 
 
        For I In 0 .. n_Len - 1 Loop 
          Element_Node := Xmldom.Item(Node_List, I); 
          v_Name       := Xmldom.Getnodename(Element_Node); 
 
          If v_Name = 'element' Then 
            If Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'unit') Is Not Null Then 
              v_Content := Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)); 
 
              If Instr(v_Content, 'textstyleno') > 0 Then 
                v_Content := ''; 
              End If; 
              --����е�λ 
              v_Result := v_Result || v_Content || Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'unit'); 
            Else 
              p_Node := Xmldom.Getparentnode(Element_Node); 
              If Xmldom.Getnodename(p_Node) <> 'enumvalues' Then 
                v_Result := v_Result || Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)); 
              End If; 
            End If; 
          Elsif v_Name = 'utext' Then 
            v_Result := v_Result || LTrim(LTrim(Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)), ':'), '��'); 
          Elsif v_Name = 'e_list' Or v_Name = 'e_enum' Or v_Name = 'e_etree' Or v_Name = 'e_utree' Then 
            Enumeration_List := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Element_Node), 'enumeration'); 
            n_Num            := Xmldom.Getlength(Enumeration_List); 
 
            If v_Name = 'e_enum' And n_Num > 0 Then 
              For J In 0 .. n_Num - 1 Loop 
                Enumeration_Node := Xmldom.Item(Enumeration_List, J); 
                Item_List        := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Element_Node), 'item'); 
                n_j              := Xmldom.Getlength(Item_List); 
 
                For K In 0 .. n_j - 1 Loop 
                  Item_Node := Xmldom.Item(Item_List, K); 
                  If Xmldom.Getattribute(Xmldom.Makeelement(Item_Node), 'checked') = '1' Then 
                    v_Val := Xmldom.Getattribute(Xmldom.Makeelement(Item_Node), 'val'); 
 
                    For Z In 0 .. n_j - 1 Loop 
                      Item_Node1 := Xmldom.Item(Item_List, Z); 
                      If Xmldom.Getattribute(Xmldom.Makeelement(Item_Node1), 'val') = v_Val And 
                         Xmldom.Getattribute(Xmldom.Makeelement(Item_Node1), 'issymbol') = '0' Then 
                        v_Result := v_Result || Xmldom.Getnodevalue(Xmldom.Getfirstchild(Item_Node1)); 
                        Exit; 
                      End If; 
                    End Loop; 
                  End If; 
                End Loop; 
              End Loop; 
            Else 
              --���ﴦ��ö�����޵���� 
              v_Eleid := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'sid'); --��ȡԪ��ID 
 
              Select Extractvalue(b.ֵ������, '/root/multisel') 
              Into v_Multisel 
              From Ӱ�񱨸�Ԫ���嵥 A, Ӱ�񱨸�ֵ���嵥 B 
              Where a.ֵ��id = b.id And a.id = Hextoraw(v_Eleid); 
 
              If v_Multisel = 2 And v_Name = 'e_enum' Then 
                --Ϊ�Ƿ����͵�ö�� 
                v_Result := v_Result || Xmldom.Getnodevalue(Xmldom.getLastChild(Element_Node)); 
              Else 
                Enum_Node := Xmldom.Item(Xmldom.Getelementsbytagname(Xmldom.Makeelement(Element_Node), 'enumvalues'), 0); 
                e_Node := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Enum_Node), 'element'); 
                n_Enum := Xmldom.Getlength(e_Node); 
 
                For K In 0 .. n_Enum - 1 Loop 
                  c_Node   := Xmldom.Item(e_Node, K); 
                  v_Result := v_Result || Xmldom.Getattribute(Xmldom.Makeelement(c_Node), 'showtext'); 
 
                  If K <> n_Enum - 1 Then 
                    v_Result := v_Result || '��'; 
                  End If; 
                End Loop; 
              End If; 
            End If; 
          End If; 
        End Loop; 
      End If; 
    End Loop; 
 
    Xmldom.Freedocument(Xcdom); 
 
    Return translate(v_Result,chr(13)||chr(10),','); 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zlpub_Pacs_��ȡ�ĵ�����; 
/


CREATE OR REPLACE Function Zlpub_Pacs_��ȡ��������
( 
  ����ID_In   In Number, 
  ������Դ_In In Number,
  �������_In In Varchar2 
) Return Varchar2 Is 
  Type t_Str_Table Is Table Of Varchar2(4000);
  a_Return t_Str_Table := t_Str_Table();
    
  v_Return        Varchar2(4000);
  n_Count         Number(2);
  n_����ID        Number(18);
  
  v_Sql           Varchar2(1000);
Begin
  n_����ID := ����id_In;
  v_Return       := '';
   
  If ������Դ_In = 2 Then
     v_Sql := 'Select ����ID From ����ҽ������ Where RISID=:1';
     Execute Immediate v_Sql Into n_����ID Using ����id_In;
  End If;
  
  Begin
    Select Decode(�Ƿ���, 1, �����ı� || Chr(10) || Chr(13), �����ı�) Bulk Collect
    Into a_Return
    From ���Ӳ�������
    Where ��ֹ�� = 0  And ��������=2 And �ļ�id = n_����ID
    Start With  ��id = (Select ID From ���Ӳ������� Where �ļ�id = n_����ID And �����ı� = �������_In And �������� = 1) 
    Connect By Prior ID=��ID
    Order By �������;
      
    For n_Count In 1 .. a_Return.Count Loop
      If v_Return Is Null Then
        v_Return := a_Return(n_Count);
      Else
        v_Return := v_Return || a_Return(n_Count);
      End If;
    End Loop;
      
  Exception
    When Others Then
      v_Return := Null;
  End;
    
  Begin
    If v_Return Is Null Then
      Select Decode(�Ƿ���, 1, �����ı� || Chr(10) || Chr(13), �����ı�) Bulk Collect
      Into a_Return            
      From ���Ӳ�������
      Where ��ֹ�� = 0  And ��������=2 And �ļ�id = n_����ID
      Start With  ��id = (Select ID From ���Ӳ������� Where �ļ�id = n_����ID And �����ı� = �������_In And �������� = 3) 
      Connect By Prior ID=��ID
      Order By �������;
           
      For n_Count In 1 .. a_Return.Count Loop
        If v_Return Is Null Then
          v_Return := a_Return(n_Count);
        Else
          v_Return := v_Return || a_Return(n_Count);
        End If;
      End Loop;   
         
    End If;
  Exception
    When Others Then
      v_Return := Null;
  End;
    
  If v_Return Is Null Then
    Select Decode(�Ƿ���, 1, �����ı� || Chr(10) || Chr(13), �����ı�) Bulk Collect
    Into a_Return
    From ���Ӳ�������
    Where ��ֹ�� = 0  And Substr(��������,1,1) = '0' And �ļ�id = n_����ID
    Start With  ��id = (Select ID From ���Ӳ������� Where �ļ�id = n_����ID And �����ı� = �������_In And �������� = 1) 
    Connect By Prior ID=��ID
    Order By �������;

    For n_Count In 1 .. a_Return.Count Loop
      If v_Return Is Null Then
        v_Return := a_Return(n_Count);
      Else
        v_Return := v_Return || a_Return(n_Count);
      End If;
    End Loop;
  End If;
  
  Return v_Return;  
End Zlpub_Pacs_��ȡ��������;
/

Create Or Replace Function Zlpub_Pacs_��ȡ�������
(
  ����id_In   In Varchar2,
  ������Դ_In In Number,
  �������_In In Varchar2
) Return Varchar2 Is

  v_Result        Varchar2(4000);
  v_Singleresult  Varchar2(4000);
  v_Sql           Varchar2(1000);
  
Begin
  v_Result       := '';
  v_Singleresult := '';

  If ������Դ_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_��ȡ�ĵ�����(:1, :2) From Dual';
    Execute Immediate v_Sql Into v_Singleresult Using ����id_In,�������_In;
  Else
    v_Sql := 'Select Zlpub_Pacs_��ȡ��������(:1, :2, :3) From Dual';
    Execute Immediate v_Sql Into v_Singleresult Using ����id_In,������Դ_In, �������_In;
  End If;
  
  If v_Result Is Null And Not v_Singleresult Is Null Then
    v_Result := v_Singleresult;
  Elsif Not v_Singleresult Is Null Then
    v_Result := v_Result || ';' || v_Singleresult;
  End If;
    
  Return v_Result;
      
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�������;
/
