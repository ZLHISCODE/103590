--��Ҫ�����������µ��༭����������Һ��ѡ��������ű���󣬲��˻�����ϸ��Ŀ���=0�����ݡ�
DECLARE
  --��ȡ��Ժ���˵����µ��ļ�
  CURSOR Cur_File IS
    SELECT d.Rowid, d.��Ŀ���, d.��¼����
    FROM ������Ϣ e, ������ҳ f, ���˻����ļ� b, �����ļ��б� a, ���˻������� c, ���˻�����ϸ d
    WHERE e.��Ժ = 1 AND e.����id = f.����id AND e.סԺ���� = f.��ҳid AND e.����id = b.����id AND e.סԺ���� = b.��ҳid AND f.��Ժ���� IS NULL AND
          a.���� = 3 AND a.���� = -1 AND b.��ʽid = a.Id AND b.Id = c.�ļ�id AND c.Id = d.��¼id AND d.��Ŀ���� = 1 AND d.��Ŀ��� = 0;
  TYPE t_File IS TABLE OF Cur_File%ROWTYPE;
  t_Count t_File;
BEGIN

  OPEN Cur_File;
  LOOP
    FETCH Cur_File BULK COLLECT
      INTO t_Count LIMIT 100;
    EXIT WHEN t_Count.Count = 0;
    FOR i IN 1 .. t_Count.Count LOOP
      IF t_Count(i).��¼���� LIKE '%C' THEN
        UPDATE ���˻�����ϸ SET ��¼����=1,��Ŀ��� = 9, ��Ŀ���� = '��Һ��' WHERE ROWID = t_Count(i).Rowid;
      END IF;
      IF t_Count(i).��¼���� = '��' OR t_Count(i).��¼���� = '��' OR t_Count(i).��¼���� LIKE '%E' THEN
        UPDATE ���˻�����ϸ SET ��¼����=1,��Ŀ��� = 10, ��Ŀ���� = '������' WHERE ROWID = t_Count(i).Rowid;
      END IF;
    END LOOP;
  END LOOP;
  CLOSE Cur_File;
END;
/ 
COMMIT;
