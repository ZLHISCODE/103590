
using System;

namespace zlMsgModel
{
    public class MsgProvider
    {
        public long ProviderCode { get; set; }
        public string PrividerName { get; set; }
        public string AppKey { get; set; }
        public string AppSecret { get; set; }

        public MsgProvider(long providerCode, string prividerName, string appKey, string appSecret)
        {
            ProviderCode = providerCode;
            PrividerName = prividerName;
            AppKey = appKey;
            AppSecret = appSecret;
        }

    }
}
