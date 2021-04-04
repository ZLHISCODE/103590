
--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_获取报告列表
(
  病人id_In In 病人医嘱记录.病人id%Type,
  主页id_In In 病人医嘱记录.主页id%Type
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
                           Select b.病历id || '''' As ID, a.Id As YZID, Decode(d.检查uid, Null, 0, 1) As GP, a.医嘱内容 As MC, 0 As BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                           From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                           Where a.Id = b.医嘱id And b.病历id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And B.RISID Is Null And
                           c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                           a.病人id = :1 And nvl(a.主页id,0) = :2';  
    Begin                   
        Execute Immediate v_Sql Using 病人id_In,主页id_In;
    Exception
      When Others Then
        Begin
          v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ) 
                                 Select b.病历id || '''' As ID, a.Id As YZID, Decode(d.检查uid, Null, 0, 1) As GP, a.医嘱内容 As MC, 0 As BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                                 From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                                 Where a.Id = b.医嘱id And b.病历id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And
                                 c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                                 a.病人id = :1 And nvl(a.主页id,0) = :2';  
          Execute Immediate v_Sql Using 病人id_In,主页id_In;
        Exception
          When Others Then Null;
        end; 
    End;
    
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ) 
                          Select b.检查报告id || '''' As ID, a.Id As YZID, Decode(d.检查uid, Null, 0, 1) As GP, a.医嘱内容 As MC, 1 as BGLX, c.最后编辑人 As BGR, c.最后审核时间 As BGSJ
                          From 病人医嘱记录 A, 病人医嘱报告 B, 影像报告记录 C, 影像检查记录 D
                          where a.Id = b.医嘱id And b.检查报告id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And
                          c.最后审核时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                          a.病人id = :1 And nvl(a.主页id,0) = :2';
    Begin
        Execute Immediate v_Sql Using 病人id_In,主页id_In;
    Exception
      When Others Then Null;
    End;
     
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, GP, MC, BGLX, BGR, BGSJ)
                          Select b.RISID || '''' As ID, a.Id As YZID, 2 As GP, a.医嘱内容 As MC, 2 as BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                          From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                          Where a.Id = b.医嘱id And b.病历ID = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And B.RISID Is Not Null And
                          c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                          a.病人id = :1 And nvl(a.主页id,0) = :2';     
    Begin
        Execute Immediate v_Sql Using 病人id_In,主页id_In; 
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
End Zlpub_Pacs_获取报告列表;
/


--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_获取报告列表Ex
(
  医嘱ID_In In 病人医嘱记录.id%Type
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
                           Select b.病历id || '''' As ID, a.Id As YZID, a.医嘱内容 As YXNR, c.病历名称 as MC, 0 As BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                           From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                           Where a.Id = b.医嘱id And b.病历id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And B.RISID Is Null And
                           c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                           a.ID = :1';  
    Begin                   
        Execute Immediate v_Sql Using 医嘱ID_In;
    Exception
      When Others Then
        Begin
          v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ) 
                                 Select b.病历id || '''' As ID, a.Id As YZID, a.医嘱内容 As YZNR, c.病历名称 As MC, 0 As BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                                 From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                                 Where a.Id = b.医嘱id And b.病历id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And
                                 c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                                 a.id = :1';  
          Execute Immediate v_Sql Using 医嘱ID_In;
        Exception
          When Others Then Null;
        end; 
    End;
    
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ) 
                          Select b.检查报告id || '''' As ID, a.Id As YZID, a.医嘱内容 As YZNR, c.文档标题 As MC, 1 as BGLX, c.最后编辑人 As BGR, c.最后审核时间 As BGSJ
                          From 病人医嘱记录 A, 病人医嘱报告 B, 影像报告记录 C, 影像检查记录 D
                          where a.Id = b.医嘱id And b.检查报告id = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And
                          c.最后审核时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                          a.id = :1';
    Begin
        Execute Immediate v_Sql Using 医嘱ID_In;
    Exception
      When Others Then Null;
    End;
     
    
    v_Sql := 'Insert Into zlTempReportList(Id, YZID, YZNR, MC, BGLX, BGR, BGSJ)
                          Select b.RISID || '''' As ID, a.Id As YZID, a.医嘱内容 As YZNR, c.病历名称 As MC, 2 as BGLX, c.保存人 As BGR, c.完成时间 As BGSJ
                          From 病人医嘱记录 A, 病人医嘱报告 B, 电子病历记录 C, 影像检查记录 D
                          Where a.Id = b.医嘱id And b.病历ID = c.Id And a.诊疗类别 = ''D'' And 相关id Is Null And B.RISID Is Not Null And
                          c.完成时间 Is Not Null And a.Id = d.医嘱id(+) And a.医嘱期效 = 1 And a.医嘱状态 In (3, 5, 6, 7, 8) And
                          a.id = :1';     
    Begin
        Execute Immediate v_Sql Using 医嘱ID_In; 
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
End Zlpub_Pacs_获取报告列表Ex;
/


--******************************************************************************************

CREATE OR REPLACE Function Zlpub_Pacs_获取文档提纲
(
  报告id_In In 病人医嘱报告.检查报告ID%Type
) Return Varchar2 Is
  v_报告提纲 Varchar2(1000);
  v_提纲内容 Varchar2(100);

  x_Content xmltype;
  n_NodeNum number(2);
  Xcdom            Xmldom.Domdocument;
  Section_List     Xmldom.Domnodelist;
Begin
    v_报告提纲 := '';

    Select b.报告内容 Into x_Content From 病人医嘱报告 a, 影像报告记录 b Where a.检查报告id=b.id And  a.检查报告id = 报告id_In;

    Xcdom         := Xmldom.Newdomdocument(x_Content);
    Section_List  := Xmldom.Getelementsbytagname(Xcdom, 'section');
    n_NodeNum     := Xmldom.Getlength(Section_List);

    For i in 0..n_NodeNum-1 Loop
      v_提纲内容 := Xmldom.Getattribute(Xmldom.Makeelement(Xmldom.Item(Section_List, i)), 'title');

      If Nvl(v_提纲内容,' ') != ' ' Then
        v_报告提纲 := v_报告提纲 || '<split>' || v_提纲内容;
      End If;
    End Loop;
    
    Return(Substr(v_报告提纲, 8));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_获取文档提纲;
/


CREATE OR REPLACE Function Zlpub_Pacs_获取病历提纲
(
  报告id_In In 病人医嘱报告.病历id%Type
) Return Varchar2 Is
  v_报告提纲 Varchar2(1000);

  Cursor c_报告提纲 Is
    Select Distinct a.内容文本
    From 电子病历内容 A, 电子病历内容 B, 病人医嘱报告 C
    Where a.对象类型 = 3 And a.Id = b.父id And b.对象类型 = 2 And b.终止版 = 0 And a.文件id = c.病历id And c.病历id = 报告id_In;
Begin
  For Row_Cols In c_报告提纲 Loop
    v_报告提纲 := '<split>' || Row_Cols.内容文本 || v_报告提纲;
  End Loop;

  Return(Substr(v_报告提纲, 8));

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_获取病历提纲;
/

CREATE OR REPLACE Function Zlpub_Pacs_获取报告提纲
(
  报告id_In In Varchar2,
  报告来源_In In Number
) Return Varchar2 Is
  v_报告提纲 Varchar2(2000);
  n_病历ID Number(18);
  
  v_Sql Varchar2(100);
Begin
  If 报告来源_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_获取文档提纲(:1)  From Dual';  
    Begin                   
        Execute Immediate v_Sql Into v_报告提纲 Using 报告id_In ;
    Exception
      When Others Then v_报告提纲 := '';          
    End;
  Else
    n_病历ID := To_Number(报告id_In);
      
    If 报告来源_In = 2 Then
      v_Sql := 'Select 病历ID From 病人医嘱报告 Where RISID=:1';
      Execute Immediate v_Sql Into n_病历ID Using 报告id_In ;
    End If;
      
    v_Sql := 'Select Zlpub_Pacs_获取病历提纲(:1)  From Dual';  
    Begin                   
        Execute Immediate v_Sql Into v_报告提纲 Using n_病历ID ;
    Exception
      When Others Then v_报告提纲 := '';          
    End;
  End If;

  Return v_报告提纲;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_获取报告提纲;
/


--******************************************************************************************


CREATE OR REPLACE Function Zlpub_Pacs_获取文档文本
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
  --标记变量
  
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
    
    --开始某个文件读取
    Begin
      Select a.医嘱id,
             Appendchildxml(Docxml, '/ZlEPR',
                             Xmlelement("Document",
                                         Xmlattributes(b.姓名 As "姓名", b.病人id As "病人ID", b.主页id As "主页ID", a.文档标题 As "文件名",
                                                        Rawtohex(a.Id) As "文件ID")))
      Into n_Adviceid, Docxml
      From 影像报告记录 A, 病人医嘱记录 B
      Where a.Id = Hextoraw(File_Id) And a.医嘱id = b.Id;
    Exception
      --给定的文件ID无效
      When Others Then
        Return Null;
    End;
    
    Select Insertchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]', 'Compend',
                           Xmlelement("Compend", Xmlattributes('0' As "ID", '内容' As "Name")))
    Into Docxml
    From Dual;
      
    --开始读取内容
    Select b.报告内容 Into x_Content From 影像报告记录 B Where b.Id || '' = File_Id;
      
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
        --提纲
        v_Title := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'title');
        v_Id    := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'sid');
          
        Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]',
                               Xmlelement("Compend", Xmlattributes(v_Title As "Name", v_Id As "ID")))
        Into Docxml
        From Dual;
      Elsif v_Name = 'utext' Then
        --文本
        v_Text := LTrim(LTrim(Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)), ':'), '：');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Text", Xmlattributes(v_Newline As "NewLine"), v_Text))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Text", Xmlattributes(v_Newline As "NewLine"), v_Text))
          Into Docxml
          From Dual;
        End If;
      Elsif v_Name = 'element' Then
        --要素
        v_Title := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'title');
        v_Text  := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'value') ||
                   Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'unit');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Element", Xmlattributes(v_Title As "Name", v_Newline As "NewLine"),
                                             v_Text))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Element", Xmlattributes(v_Title As "Name", v_Newline As "NewLine"),
                                             v_Text))
          Into Docxml
          From Dual;
        End If;
      Elsif v_Name = 'image' Then
        --图片
        n_Width  := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'width');
        n_Height := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'height');
        v_Name   := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'key');
        v_Type   := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'class');
          
        If Nvl(v_Name, ' ') <> ' ' Then
          If Nvl(v_Id, ' ') = ' ' Then
            Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                                   Xmlelement("Picture",
                                               Xmlattributes(n_Width As "OrigWidth", n_Height As "OrigHeight",
                                                              n_Width As "ShowWidth", n_Height As "ShowHeight",
                                                              v_Name As "PicName", n_Adviceid As "AdviceID",
                                                              v_Type As "Type")))
            Into Docxml
            From Dual;
          Else
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@文件ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
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
        --签名
        v_Text := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'displayinfo');
          
        If Nvl(v_Id, ' ') = ' ' Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                                 Xmlelement("Sign", Xmlattributes(v_Newline As "NewLine"), Zl_Eprsplit(v_Text, ';', 1)))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID="' || File_Id || '"]/descendant::Compend[@ID="' || v_Id || '"]',
                                 Xmlelement("Sign", Xmlattributes(v_Newline As "NewLine"), Zl_Eprsplit(v_Text, ';', 1)))
          Into Docxml
          From Dual;
        End If;
      End If;
    End Loop;
      
    For Aa In (Select '/' || a.Ftp目录 || '/ReportImages/' || To_Char(b.创建时间, 'YYYYMMDD') || '/' || b.Id || '/' As v_Ftppath
               From 影像设备目录 A, 影像报告记录 B
               Where a.设备号 = b.设备号 And b.Id = File_Id) Loop
        
      Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                             Xmlelement("FtpPath", Xmlattributes(v_Newline As "NewLine"), Aa.v_Ftppath))
      Into Docxml
      From Dual;
    End Loop;
  End Loop;
  
  Return Docxml;
End Zlpub_Pacs_获取文档文本;
/
 

CREATE OR REPLACE Function Zlpub_Pacs_获取病历文本
(
  Ids_In   In Varchar2,
  From_In Number
)Return XmlType Is
  Docxml XmlType;
  
  File_Id Varchar2(32);
  n_Adviceid Number(18);
  
  v_Sql        Varchar2(1000);
  
  --标记变量
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
       --RIS报告
       v_Sql := 'Select 病历Id From 病人医嘱报告 Where RISID = :1';
       Execute Immediate v_Sql Into File_Id Using File_Id;
    End If;
        
    --开始某个病历文件读取
    Begin
      Select Appendchildxml(Docxml, '/ZlEPR',
                             Xmlelement("Document",
                                         Xmlattributes(b.姓名 As "姓名", a.病人id As "病人ID", a.主页id As "主页ID", a.病历名称 As "文件名",
                                                        a.Id As "文件ID")))
      Into Docxml
      From 电子病历记录 A, 病人信息 B
      Where a.Id = File_Id And a.编辑方式 = 0 And a.病人id = b.病人id;
          
      Select 医嘱id Into n_Adviceid From 病人医嘱报告 Where 病历id = File_Id;
    Exception
      --给定的病历文件ID无效
      When Others Then Return Null;
    End;
        
    Select Insertchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]', 'Compend',
                           Xmlelement("Compend", Xmlattributes('0' As "ID", '内容' As "Name")))
    Into Docxml
    From Dual;
      
    Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                           Xmlelement("Text", Xmlattributes(Nvl(Null, 0) As "NewLine"), '内容文本'))
    Into Docxml
    From Dual;
      
    For Rs In (Select ID, 父id, 对象序号, 对象类型, 对象属性, 内容行次, 内容文本, 是否换行, 要素名称
               From (Select ID, 父id, 对象序号, 对象类型, 对象属性, 内容行次, 内容文本, 是否换行, 要素名称
                      From 电子病历内容
                      Where 文件id = File_Id And 对象序号 > 0 And 对象序号 <> ID And 终止版 = 0)
               Start With 父id Is Null
               Connect By Prior ID = 父id
               Order Siblings By 对象序号, 内容行次) Loop
      If Rs.对象类型 = 1 Then
        --提纲
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']',
                                 Xmlelement("Compend", Xmlattributes(Rs.内容文本 As "Name", Rs.Id As "ID")))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Compend", Xmlattributes(Rs.内容文本 As "Name", Rs.Id As "ID")))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.对象类型 = 2 Then
        --文本
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Text", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"), Rs.内容文本))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Text", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"), Rs.内容文本))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.对象类型 = 3 Then
        --表格
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Table",
                                             Xmlattributes(Zl_Eprsplit(Rs.对象属性, ';', 1) As "Rows",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 2) As "Cols",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 3) As "Width",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 4) As "Height",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 5) As "ColWidthString",
                                                            Nvl(Rs.是否换行, 0) As "NewLine", Rs.Id As "ID"), Rs.内容文本))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Table",
                                             Xmlattributes(Zl_Eprsplit(Rs.对象属性, ';', 1) As "Rows",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 2) As "Cols",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 3) As "Width",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 4) As "Height",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 5) As "ColWidthString",
                                                            Nvl(Rs.是否换行, 0) As "NewLine", Rs.Id As "ID"), Rs.内容文本))
          Into Docxml
          From Dual;
        End If;
          
        ---对表格的单元格进行填充
        For Rs_Cell In (Select ID, 父id, 对象序号, 对象类型, 对象属性, 内容行次, 内容文本, 是否换行, 要素名称
                        From 电子病历内容
                        Where 文件id = File_Id And 父id = Rs.Id And 终止版 = 0
                        Order By 内容行次, ID) Loop
          If Rs_Cell.对象类型 = 2 Or Rs_Cell.对象类型 = 4 Then
            If Zl_Eprsplit(Rs_Cell.对象属性, '|', 26) Is Null Then
              --兼容历史病历
              Select Appendchildxml(Docxml,
                                     '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                     Xmlelement("Cell",
                                                 Xmlattributes(Zl_Eprsplit(Rs_Cell.对象属性, '|', 2) As "Row",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 3) As "Col",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 2) || '_' ||
                                                                 Zl_Eprsplit(Rs_Cell.对象属性, '|', 3) As "Row_Col",
                                                                Decode(Rs_Cell.对象类型, 2, 0, 4, 1) As "Type",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 5) As "Width",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 6) As "Height",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 4) As "MergeNo",
                                                                Nvl(Rs.是否换行, 0) As "NewLine", Rs_Cell.Id As "ID"),
                                                 Rs_Cell.内容文本))
              Into Docxml
              From Dual;
            Else
              Select Appendchildxml(Docxml,
                                     '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                     Xmlelement("Cell",
                                                 Xmlattributes(Zl_Eprsplit(Rs_Cell.对象属性, '|', 3) As "Row",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 4) As "Col",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 3) || '_' ||
                                                                 Zl_Eprsplit(Rs_Cell.对象属性, '|', 4) As "Row_Col",
                                                                Decode(Rs_Cell.对象类型, 2, 0, 4, 1) As "Type",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '', 6) As "Width",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 7) As "Height",
                                                                Zl_Eprsplit(Rs_Cell.对象属性, '|', 5) As "MergeNo",
                                                                Nvl(Rs.是否换行, 0) As "NewLine", Rs_Cell.Id As "ID"),
                                                 Rs_Cell.内容文本))
              Into Docxml
              From Dual;
            End If;
          Elsif Rs_Cell.对象类型 = 5 And Zl_Eprsplit(Rs_Cell.对象属性, ';', 1) = 2 Then
            --单元格图由Webservice直接读取BLOB之后直接写文件以提高速度
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id || ']',
                                   Xmlelement("Picture",
                                               Xmlattributes(Zl_Eprsplit(Rs_Cell.对象属性, ';', 2) As "Row",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 3) As "Col",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 8) As "OrigWidth",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 9) As "OrigHeight",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 6) As "ShowWidth",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 7) As "ShowHeight",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.对象属性, ';', 12), ' ') As "PicName",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.对象属性, ';', 13), '0') As "AdviceID",
                                                              Rs_Cell.Id As "ID"), Rs_Cell.Id))
            Into Docxml
            From Dual;
              
            If Nvl(n_Adviceid, 0) = 0 Then
              n_Adviceid := Nvl(Zl_Eprsplit(Rs_Cell.对象属性, ';', 13), '0');
            End If;
          Elsif Rs_Cell.对象类型 = 5 And Zl_Eprsplit(Rs_Cell.对象属性, ';', 1) <> 2 Then
            --单元格图由Webservice直接读取BLOB之后直接写文件以提高速度
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Table[@ID=' || Rs.Id ||
                                    ']/Cell[@Row_Col="' || Zl_Eprsplit(Rs_Cell.对象属性, ';', 2) || '_' ||
                                    Zl_Eprsplit(Rs_Cell.对象属性, ';', 3) || '"]',
                                   Xmlelement("Picture",
                                               Xmlattributes(Zl_Eprsplit(Rs_Cell.对象属性, ';', 2) As "Row",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 3) As "Col",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 8) As "OrigWidth",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 9) As "OrigHeight",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 6) As "ShowWidth",
                                                              Zl_Eprsplit(Rs_Cell.对象属性, ';', 7) As "ShowHeight",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.对象属性, ';', 12), ' ') As "PicName",
                                                              Nvl(Zl_Eprsplit(Rs_Cell.对象属性, ';', 13), '0') As "AdviceID",
                                                              Rs_Cell.Id As "ID"), Rs_Cell.Id))
            Into Docxml
            From Dual;
            --制作标记子节点集
            v_Mark  := '';
            Makxml  := Null;
            Maksxml := Null;
            For Rs_Mark In (Select ID, 父id, 内容文本, 内容行次
                            From 电子病历内容
                            Where 父id = Rs_Cell.Id
                            Order By 内容行次) Loop
              v_Marks := v_Mark || Rs_Mark.内容文本;
              v_Marks := Replace(v_Marks, '||', '^');
              For I In 1 .. 100 Loop
                v_Mark := Zl_Eprsplit(v_Marks, '^', I);
                If Zl_Eprsplit(v_Mark, '|', 15) Is Null Then
                  --最后一个标记信息不全，存在下一行中
                  Exit;
                Else
                  Select Xmlelement("Mark",
                                     Xmlforest(Zl_Eprsplit(v_Mark, '|', 2) As "类型",
                                                Zl_Eprsplit(v_Mark, '|', 3) As "内容", Zl_Eprsplit(v_Mark, '|', 4) As "点集",
                                                Zl_Eprsplit(v_Mark, '|', 5) As "X1", Zl_Eprsplit(v_Mark, '|', 6) As "Y1",
                                                Zl_Eprsplit(v_Mark, '|', 7) As "X2", Zl_Eprsplit(v_Mark, '|', 8) As "Y2",
                                                Zl_Eprsplit(v_Mark, '|', 9) As "填充色",
                                                Zl_Eprsplit(v_Mark, '|', 10) As "填充方式",
                                                Zl_Eprsplit(v_Mark, '|', 11) As "线条色",
                                                Zl_Eprsplit(v_Mark, '|', 12) As "字体色",
                                                Zl_Eprsplit(v_Mark, '|', 13) As "线型",
                                                Zl_Eprsplit(v_Mark, '|', 14) As "线宽",
                                                Zl_Eprsplit(v_Mark, '|', 15) As "字体"))
                  Into Makxml
                  From Dual;
                  Select Xmlconcat(Maksxml, Makxml) Into Maksxml From Dual;
                End If;
              End Loop;
            End Loop;
            --向Picture插入标记子节点
            Select Appendchildxml(Docxml,
                                   '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Picture[@ID=' || Rs_Cell.Id || ']',
                                   Maksxml)
            Into Docxml
            From Dual;
          End If;
        End Loop;
      Elsif Rs.对象类型 = 4 Then
        --要素
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Element", Xmlattributes(Rs.要素名称 As "Name", Nvl(Rs.是否换行, 0) As "NewLine"),
                                             Rs.内容文本))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Element", Xmlattributes(Rs.要素名称 As "Name", Nvl(Rs.是否换行, 0) As "NewLine"),
                                             Rs.内容文本))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.对象类型 = 5 And Nvl(Rs.内容行次, 0) = 0 Then
        --图片由Webservice直接读取BLOB之后直接写文件以提高速度
 
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Picture",
                                             Xmlattributes(Zl_Eprsplit(Rs.对象属性, ';', 8) As "OrigWidth",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 9) As "OrigHeight",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 6) As "ShowWidth",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 7) As "ShowHeight",
                                                            Nvl(Zl_Eprsplit(Rs.对象属性, ';', 12), ' ') As "PicName",
                                                            Nvl(Zl_Eprsplit(Rs.对象属性, ';', 13), '0') As "AdviceID",
                                                            Nvl(Rs.是否换行, 0) As "NewLine", Rs.Id As "ID"), Rs.Id))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Picture",
                                             Xmlattributes(Zl_Eprsplit(Rs.对象属性, ';', 8) As "OrigWidth",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 9) As "OrigHeight",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 6) As "ShowWidth",
                                                            Zl_Eprsplit(Rs.对象属性, ';', 7) As "ShowHeight",
                                                            Nvl(Zl_Eprsplit(Rs.对象属性, ';', 12), ' ') As "PicName",
                                                            Nvl(Zl_Eprsplit(Rs.对象属性, ';', 13), '0') As "AdviceID",
                                                            Nvl(Rs.是否换行, 0) As "NewLine", Rs.Id As "ID"), Rs.Id))
          Into Docxml
          From Dual;
        End If;
        --制作标记子节点集
        v_Mark  := '';
        Makxml  := Null;
        Maksxml := Null;
        For Rs_Mark In (Select ID, 父id, 内容文本, 内容行次 From 电子病历内容 Where 父id = Rs.Id Order By 内容行次) Loop
          v_Marks := v_Mark || Rs_Mark.内容文本;
          v_Marks := Replace(v_Marks, '||', '^');
          For I In 1 .. 100 Loop
            v_Mark := Zl_Eprsplit(v_Marks, '^', I);
            If Zl_Eprsplit(v_Mark, '|', 15) Is Null Then
              --最后一个标记信息不全，存在下一行中
              Exit;
            Else
              Select Xmlelement("Mark",
                                 Xmlforest(Zl_Eprsplit(v_Mark, '|', 2) As "类型", Zl_Eprsplit(v_Mark, '|', 3) As "内容",
                                            Zl_Eprsplit(v_Mark, '|', 4) As "点集", Zl_Eprsplit(v_Mark, '|', 5) As "X1",
                                            Zl_Eprsplit(v_Mark, '|', 6) As "Y1", Zl_Eprsplit(v_Mark, '|', 7) As "X2",
                                            Zl_Eprsplit(v_Mark, '|', 8) As "Y2", Zl_Eprsplit(v_Mark, '|', 9) As "填充色",
                                            Zl_Eprsplit(v_Mark, '|', 10) As "填充方式",
                                            Zl_Eprsplit(v_Mark, '|', 11) As "线条色", Zl_Eprsplit(v_Mark, '|', 12) As "字体色",
                                            Zl_Eprsplit(v_Mark, '|', 13) As "线型", Zl_Eprsplit(v_Mark, '|', 14) As "线宽",
                                            Zl_Eprsplit(v_Mark, '|', 15) As "字体"))
              Into Makxml
              From Dual;
              Select Xmlconcat(Maksxml, Makxml) Into Maksxml From Dual;
            End If;
          End Loop;
        End Loop;
        --向Picture插入标记子节点
        Select Appendchildxml(Docxml,
                               '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Picture[@ID=' || Rs.Id || ']',
                               Maksxml)
        Into Docxml
        From Dual;
      Elsif Rs.对象类型 = 7 Then
        --诊断
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Diagnosise", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"), Rs.内容文本))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Diagnosise", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"), Rs.内容文本))
          Into Docxml
          From Dual;
        End If;
      Elsif Rs.对象类型 = 8 Then
        --签名
        If Rs.父id Is Null Then
          Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID=' || File_Id || ']/Compend[@ID=0]',
                                 Xmlelement("Sign", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"),
                                             Zl_Eprsplit(Rs.内容文本, ';', 1)))
          Into Docxml
          From Dual;
        Else
          Select Appendchildxml(Docxml,
                                 '/ZlEPR/Document[@文件ID=' || File_Id || ']/descendant::Compend[@ID=' || Rs.父id || ']',
                                 Xmlelement("Sign", Xmlattributes(Nvl(Rs.是否换行, 0) As "NewLine"),
                                             Zl_Eprsplit(Rs.内容文本, ';', 1)))
          Into Docxml
          From Dual;
        End If;
      End If;
    End Loop;
      
    For Aa In (Select A1.Ftp目录 || '/' || To_Char(l.接收日期, 'yyyymmdd') || '/' || l.检查uid As v_Ftppath
               From 影像检查记录 L, 影像设备目录 A1
               Where l.位置一 = A1.设备号(+) And l.医嘱id = n_Adviceid) Loop
        
      Select Appendchildxml(Docxml, '/ZlEPR/Document[@文件ID="' || File_Id || '"]/Compend[@ID=0]',
                             Xmlelement("FtpPath", Xmlattributes(v_Newline As "NewLine"), Aa.v_Ftppath))
      Into Docxml
      From Dual;
    End Loop;
    
  End Loop;
  
  Return Docxml;  
End Zlpub_Pacs_获取病历文本;
/  

CREATE OR REPLACE Function Zlpub_Pacs_获取报告文本
(
  Ids_In In Varchar2,
  From_In Number
) Return Xmltype Is
--Ids_In规则是以 '|' 分隔的ID串，开始/结尾无 '|'
  --根所给定的病历文件ID串生成内容XML并返回XMLType
  Docxml XmlType;
  v_Sql Varchar2(1000);
Begin
    
  If From_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_获取文档文本(:1) From Dual';
    Execute Immediate v_Sql Into Docxml Using Ids_In;
  Else
    v_Sql := 'Select Zlpub_Pacs_获取病历文本(:1, :2) From Dual';
    Execute Immediate v_Sql Into Docxml Using Ids_In, From_In;
  End If;

  Return Docxml;
Exception
  When Others Then
    Return Null;
End Zlpub_Pacs_获取报告文本;
/

--******************************************************************************************
CREATE OR REPLACE Function Zlpub_Pacs_获取文档内容
( 
  报告ID_In In 影像报告记录.ID%Type, 
  报告提纲_In In 影像报告记录.诊断意见%Type 
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
    
    Select 报告内容 Into x_Content From 影像报告记录 Where id = 报告ID_In;

    Select Deletexml(x_Content, '//image') Into x_Content From Dual; 
 
    Xcdom := Xmldom.Newdomdocument(x_Content); 
 
    For Myrow In (Select Column_Value Name From Table(f_Str2list(报告提纲_In))) Loop 
      n_i := -1; 
      --循环提纲名称 
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
              --如果有单位 
              v_Result := v_Result || v_Content || Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'unit'); 
            Else 
              p_Node := Xmldom.Getparentnode(Element_Node); 
              If Xmldom.Getnodename(p_Node) <> 'enumvalues' Then 
                v_Result := v_Result || Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)); 
              End If; 
            End If; 
          Elsif v_Name = 'utext' Then 
            v_Result := v_Result || LTrim(LTrim(Xmldom.Getnodevalue(Xmldom.Getfirstchild(Element_Node)), ':'), '：'); 
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
              --这里处理枚举有无的情况 
              v_Eleid := Xmldom.Getattribute(Xmldom.Makeelement(Element_Node), 'sid'); --获取元素ID 
 
              Select Extractvalue(b.值域描述, '/root/multisel') 
              Into v_Multisel 
              From 影像报告元素清单 A, 影像报告值域清单 B 
              Where a.值域id = b.id And a.id = Hextoraw(v_Eleid); 
 
              If v_Multisel = 2 And v_Name = 'e_enum' Then 
                --为是否类型的枚举 
                v_Result := v_Result || Xmldom.Getnodevalue(Xmldom.getLastChild(Element_Node)); 
              Else 
                Enum_Node := Xmldom.Item(Xmldom.Getelementsbytagname(Xmldom.Makeelement(Element_Node), 'enumvalues'), 0); 
                e_Node := Xmldom.Getelementsbytagname(Xmldom.Makeelement(Enum_Node), 'element'); 
                n_Enum := Xmldom.Getlength(e_Node); 
 
                For K In 0 .. n_Enum - 1 Loop 
                  c_Node   := Xmldom.Item(e_Node, K); 
                  v_Result := v_Result || Xmldom.Getattribute(Xmldom.Makeelement(c_Node), 'showtext'); 
 
                  If K <> n_Enum - 1 Then 
                    v_Result := v_Result || '、'; 
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
End Zlpub_Pacs_获取文档内容; 
/


CREATE OR REPLACE Function Zlpub_Pacs_获取病历内容
( 
  报告ID_In   In Number, 
  报告来源_In In Number,
  报告提纲_In In Varchar2 
) Return Varchar2 Is 
  Type t_Str_Table Is Table Of Varchar2(4000);
  a_Return t_Str_Table := t_Str_Table();
    
  v_Return        Varchar2(4000);
  n_Count         Number(2);
  n_病历ID        Number(18);
  
  v_Sql           Varchar2(1000);
Begin
  n_病历ID := 报告id_In;
  v_Return       := '';
   
  If 报告来源_In = 2 Then
     v_Sql := 'Select 病历ID From 病人医嘱报告 Where RISID=:1';
     Execute Immediate v_Sql Into n_病历ID Using 报告id_In;
  End If;
  
  Begin
    Select Decode(是否换行, 1, 内容文本 || Chr(10) || Chr(13), 内容文本) Bulk Collect
    Into a_Return
    From 电子病历内容
    Where 终止版 = 0  And 对象类型=2 And 文件id = n_病历ID
    Start With  父id = (Select ID From 电子病历内容 Where 文件id = n_病历ID And 内容文本 = 报告提纲_In And 对象类型 = 1) 
    Connect By Prior ID=父ID
    Order By 对象序号;
      
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
      Select Decode(是否换行, 1, 内容文本 || Chr(10) || Chr(13), 内容文本) Bulk Collect
      Into a_Return            
      From 电子病历内容
      Where 终止版 = 0  And 对象类型=2 And 文件id = n_病历ID
      Start With  父id = (Select ID From 电子病历内容 Where 文件id = n_病历ID And 内容文本 = 报告提纲_In And 对象类型 = 3) 
      Connect By Prior ID=父ID
      Order By 对象序号;
           
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
    Select Decode(是否换行, 1, 内容文本 || Chr(10) || Chr(13), 内容文本) Bulk Collect
    Into a_Return
    From 电子病历内容
    Where 终止版 = 0  And Substr(对象属性,1,1) = '0' And 文件id = n_病历ID
    Start With  父id = (Select ID From 电子病历内容 Where 文件id = n_病历ID And 内容文本 = 报告提纲_In And 对象类型 = 1) 
    Connect By Prior ID=父ID
    Order By 对象序号;

    For n_Count In 1 .. a_Return.Count Loop
      If v_Return Is Null Then
        v_Return := a_Return(n_Count);
      Else
        v_Return := v_Return || a_Return(n_Count);
      End If;
    End Loop;
  End If;
  
  Return v_Return;  
End Zlpub_Pacs_获取病历内容;
/

Create Or Replace Function Zlpub_Pacs_获取提纲内容
(
  报告id_In   In Varchar2,
  报告来源_In In Number,
  报告提纲_In In Varchar2
) Return Varchar2 Is

  v_Result        Varchar2(4000);
  v_Singleresult  Varchar2(4000);
  v_Sql           Varchar2(1000);
  
Begin
  v_Result       := '';
  v_Singleresult := '';

  If 报告来源_In = 1 Then
    v_Sql := 'Select Zlpub_Pacs_获取文档内容(:1, :2) From Dual';
    Execute Immediate v_Sql Into v_Singleresult Using 报告id_In,报告提纲_In;
  Else
    v_Sql := 'Select Zlpub_Pacs_获取病历内容(:1, :2, :3) From Dual';
    Execute Immediate v_Sql Into v_Singleresult Using 报告id_In,报告来源_In, 报告提纲_In;
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
End Zlpub_Pacs_获取提纲内容;
/
