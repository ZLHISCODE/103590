using System.Runtime.InteropServices;//这是用到DllImport时候要引入的包

namespace ZLSOFT.HIS.ZLIDCard
{
    /// <summary>
    /// 身份证阅读类
    /// </summary>
     class CVRSDK
    {
        [DllImport("Termb.dll", EntryPoint = "CVR_InitComm", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_InitComm(int Port);


        [DllImport("Termb.dll", EntryPoint = "CVR_Authenticate", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_Authenticate();


        [DllImport("Termb.dll", EntryPoint = "CVR_Read_Content", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_Read_Content(int Active);


        [DllImport("Termb.dll", EntryPoint = "CVR_Read_FPContent", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_Read_FPContent();

        [DllImport("Termb.dll", EntryPoint = "CVR_FindCard", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_FindCard();

        [DllImport("Termb.dll", EntryPoint = "CVR_SelectCard", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_SelectCard();


        [DllImport("Termb.dll", EntryPoint = "CVR_CloseComm", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int CVR_CloseComm();

        [DllImport("Termb.dll", EntryPoint = "GetCertType", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern  int GetCertType(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleName", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern  int GetPeopleName(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleChineseName", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern  int GetPeopleChineseName(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleNation", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int GetPeopleNation(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetNationCode", CharSet = CharSet.Ansi, SetLastError = false)]
        public static extern int GetNationCode(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleBirthday", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPeopleBirthday(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleAddress", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPeopleAddress(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleIDCode", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPeopleIDCode(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetDepartment", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetDepartment(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetStartDate", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetStartDate(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetEndDate", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetEndDate(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "GetPeopleSex", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPeopleSex(ref byte strTmp, ref int strLen);


        [DllImport("Termb.dll", EntryPoint = "CVR_GetSAMID", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int CVR_GetSAMID(ref byte strTmp);

        [DllImport("Termb.dll", EntryPoint = "GetBMPData", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetBMPData(ref byte btBmp, ref int nLen);

        [DllImport("Termb.dll", EntryPoint = "GetJpgData", CharSet = CharSet.Ansi, SetLastError = false, CallingConvention = CallingConvention.StdCall)]
        public static extern int GetJpgData(ref byte btBmp, ref int nLen);
    }
}
