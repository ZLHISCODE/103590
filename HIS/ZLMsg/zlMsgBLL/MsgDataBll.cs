using System;
using System.Collections.Generic;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using zlMsgDAL;
using zlMsgModel;

namespace zlMsgBLL
{
    public class MsgDataBll
    {
        private MsgDataDal DataDal = new MsgDataDal();

        public MsgData GetMsgDataByRowid(string strRowid)
        {
            return DataDal.GetShortMsg(strRowid);
        }

        public List<MsgData> GetMsgDatas()
        {
            return DataDal.GetErrorMsg();
        }

        public bool UpdateMsgdata(MsgData msgData)
        {
            return DataDal.UpdateShortMsg(msgData);
        }

        public bool DeleteMsgData(List<long> IDList)
        {
            return DataDal.DeleteShortMsg(IDList);
        }

        public void RegistDcn(OnChangeEventHandler changeEventHandler)
        {
            DataDal.ShortMsgRegist(changeEventHandler);
        }

        public void UnRegistDcn()
        {
            DataDal.ShortMsgUnRegist();
        }

    }
}
