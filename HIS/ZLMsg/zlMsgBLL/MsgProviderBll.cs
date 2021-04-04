using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using zlMsgDAL;
using zlMsgModel;

namespace zlMsgBLL
{
    public class MsgProviderBll
    {
        private MsgProviderDal providerDal ;

        public MsgProviderBll()
        {
            this.providerDal = new MsgProviderDal();
        }
        public long GetProviderCode()
        {
            return providerDal.GetProviderCode();
        }

        public MsgProvider GetMsgProviderByCode(long code)
        {
            return providerDal.GetMsgProvider(code);
        }

        public MsgProvider GetMsgProviderByName(string name)
        {
            return providerDal.GetMsgProvider(0,name);
        }

        public bool ProviderInsert(MsgProvider provider)
        {
            return providerDal.AddProvider(provider);
        }

        public bool ProviderUpdate(MsgProvider provider)
        {
            return providerDal.UpdateProvider(provider);
        }
    }
}
