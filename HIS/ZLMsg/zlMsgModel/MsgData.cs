using System;

namespace zlMsgModel
{
    public class MsgData
    {
        public long ID { get; set; }
        public MsgTemplate Template{ get; set; }
        public long Receiver { get; set; }
        public string Para { get; set; }
        public string Sender { get; set; }
        public string Terminal { get; set; }
        public int State { get; set; }
        public int Kind { get; set; }
        public string MsgText { get; set; }
        public string Extend { get; set; }
        public DateTime SendDate { get; set; }

        public MsgData(long iD, MsgTemplate template, long receiver, string para, string sender, string terminal, int state, int kind, string msgText, string extend, DateTime sendDate)
        {
            ID = iD;
            Template = template;
            Receiver = receiver;
            Para = para;
            Sender = sender;
            Terminal = terminal;
            State = state;
            Kind = kind;
            MsgText = msgText;
            Extend = extend;
            SendDate = sendDate;
        }
    }
}
