using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace zlShortMsg
{
    public static class ZLSM4
    {
        private enum CrypeMode
        {
            CM_Encrypt = 1,  //加密
            CM_Decrypt = 0   //解密
        };

        private enum KeyType
        {
            KT_Default = 0,  //演示密钥
            KT_IV = 1,      //分组加密的分组密钥
            KT_Key = 2      //密钥
        };

        //'SM4加密
        //'/**
        //' * \brief          SM4-ECB block encryption/decryption
        //' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
        //' * \param length   length of the input data
        //' * \param input    input block
        //' * \param output   output block
        //' */
        [DllImport("zlSm4.dll")]
        private static extern void sm4_crypt_ecb(Int32 Mode, Int32 Length, byte[] Key, byte[] in_put, byte[] out_put);

        //'SM4分组密码加密
        //'/**
        //' * \brief          SM4-CBC buffer encryption/decryption
        //' * \param mode     SM4_ENCRYPT or SM4_DECRYPT
        //' * \param length   length of the input data
        //' * \param iv       initialization vector (updated after use)
        //' * \param input    buffer holding the input data
        //' * \param output   buffer holding the output data
        //' */
        [DllImport("zlSm4.dll")]
        private static extern void sm4_crypt_cbc(Int32 Mode, Int32 Length, byte[] iv, byte[] Key, byte[] in_put, byte[] out_put);


        //'获取字符串的哈希编码
        //'/**
        //' * \brief          Output = SM3( input buffer )
        //' *
        //' * \param input    buffer holding the  data
        //' * \param ilen     length of the input data
        //' * \param output   SM3 checksum result
        //' */
        [DllImport("zlSm4.dll", EntryPoint = "sm3")]
        private static extern void sm3_hash(byte[] in_put, Int32 Length, byte[] out_put);
        //'获取文件的sm哈希编码
        //'/**
        //' * \brief          Output = SM3( file contents )
        //' *
        //' * \param path     input file name
        //' * \param output   SM3 checksum result
        //' *
        //' * \return         0 if successful, 1 if fopen failed,
        //' *                 or 2 if fread failed
        //' */
        [DllImport("zlSm4.dll", EntryPoint = "sm3_file")]
        private static extern Int32 sm3_file_hash(byte[] in_path, byte[] out_put);

        //'HMAC是密钥相关的哈希运算消息认证码，HMAC运算利用哈希算法，以一个密钥和一个消息为输入，生成一个消息摘要作为输出。
        //'/**
        //' * \brief          Output = HMAC-SM3( hmac key, input buffer )
        //' *
        //' * \param key      HMAC secret key
        //' * \param keylen   length of the HMAC key
        //' * \param input    buffer holding the  data
        //' * \param ilen     length of the input data
        //' * \param output   HMAC-SM3 result
        //' */
        [DllImport("zlSm4.dll", EntryPoint = "sm3_hmac")]
        private static extern Int32 sm3_hmac_hash(byte[] key, Int32 keylen, byte[] in_put, Int32 inputLen, byte[] out_put);


        //'获取ZLSM4的修改版本
        //'1:只支持sm4_crypt_ecb,sm4_crypt_cbc
        //'2:增加支持sm3，sm3_file，sm3_hmac，sm_version
        //'/**
        //' * \brief          Output = zlSM4.DLL Version
        //' */
        [DllImport("zlSm4.dll", EntryPoint = "sm_version")]
        private static extern Int32 get_sm_version();
        private static byte[] mbtyKey = { 231, 167, 243, 94, 93, 14, 161, 231, 69, 7, 221, 246, 10, 37, 130, 78 };
        private static byte[] mbtyIVKey = { 179, 6, 136, 130, 67, 95, 163, 5, 127, 158, 77, 11, 160, 231, 177, 102 };
        public static Int32 M_SM4_VERSION;
        //'======================================================================================================================
        //'方法           Sm4EncryptEcb           SM4加密
        //'返回值         String                  加密后的值,格式：ZLSV+版本号+:+加密后的字符串
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  要加密的字符串
        //'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'======================================================================================================================
        public static string Sm4EncryptEcb(string strInput, string strKey = null)
        {
            if (M_SM4_VERSION == 0)
            {
                M_SM4_VERSION = sm_version();
            }
            string strRet = null;
            if (strInput != null)
            {
                byte[] bytKey = GetKey(strKey, KeyType.KT_Key);
                byte[] byteIn = BytePadding(strInput);
                byte[] byteOut = new byte[byteIn.Length];
                ZLSM4.sm4_crypt_ecb(Convert.ToInt32(CrypeMode.CM_Encrypt), byteIn.Length, bytKey, byteIn, byteOut);
                strRet = "ZLSV" + M_SM4_VERSION + ":" + ByteToHexString(byteOut);
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           Sm4DecryptEcb           SM4解密
        //'返回值         String                  解密后的值
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  要解密的字符串（该字符串是Sm4EncryptEcb生成的结果）
        //'strKey         String(Optional)        加密密钥也就是解密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'======================================================================================================================
        public static string Sm4DecryptEcb(string strInput, string strKey = null)
        {
            if (M_SM4_VERSION == 0)
            {
                M_SM4_VERSION = sm_version();
            }
            string strRet = null;
            Regex regZLSV = new Regex(@"^ZLSV\d+:");
            Match maVersion = regZLSV.Match(strInput);
            if (strInput != null && maVersion.Success)
            {
                string StrPre = maVersion.Groups[0].Value;
                int intVersion = Convert.ToInt32(StrPre.Substring(4, StrPre.Length - 5));
                string strIn = strInput.Substring(StrPre.Length);
                byte[] bytKey = GetKey(strKey, KeyType.KT_Key);
                byte[] byteIn = HexStringToByte(strIn);
                byte[] byteOut = new byte[byteIn.Length];
                ZLSM4.sm4_crypt_ecb(Convert.ToInt32(CrypeMode.CM_Decrypt), byteIn.Length, bytKey, byteIn, byteOut);
                strRet = Encoding.ASCII.GetString(byteOut);
                if (intVersion == 1)
                {
                    strRet = strRet.Trim();
                }
                else
                {
                    strRet = TruncZero(strRet);
                }
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           Sm4EncryptCbc           SM4分组加密
        //'返回值         String                  加密后的值
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  要加密的字符串
        //'strKey         String(Optional)        加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'strIv          String(Optional)        分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'======================================================================================================================
        public static string Sm4EncryptCbc(string strInput, string strKey = null, string strIv = null)
        {
            if (M_SM4_VERSION == 0)
            {
                M_SM4_VERSION = sm_version();
            }
            string strRet = null;
            if (strInput != null)
            {
                byte[] bytKey = GetKey(strKey, KeyType.KT_Key);
                byte[] bytIV = GetKey(strIv, KeyType.KT_IV);

                byte[] byteIn = BytePadding(strInput);
                byte[] byteOut = new byte[byteIn.Length];
                ZLSM4.sm4_crypt_cbc(Convert.ToInt32(CrypeMode.CM_Encrypt), byteIn.Length, bytIV, bytKey, byteIn, byteOut);
                strRet = "ZLSV" + M_SM4_VERSION + ":" + ByteToHexString(byteOut);
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           Sm4EncryptCbc           SM4分组加密对应的解密过程
        //'返回值         String                  解密后的值
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  已经加密的字符串
        //'strKey         String(Optional)        解密密钥也就是加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'strIv          String(Optional)        分组解密密钥也就是分组加密密钥（32位的16进制字符串，可以通过HexStringToByte返回）
        //'======================================================================================================================
        public static string Sm4DecryptCbc(string strInput, string strKey = null, string strIv = null)
        {
            if (M_SM4_VERSION == 0)
            {
                M_SM4_VERSION = sm_version();
            }
            string strRet = null;
            Regex regZLSV = new Regex(@"^ZLSV\d+:");
            Match maVersion = regZLSV.Match(strInput);
            if (strInput != null && maVersion.Success)
            {
                string StrPre = maVersion.Groups[0].Value;
                int intVersion = Convert.ToInt32(StrPre.Substring(4, StrPre.Length - 5));
                string strIn = strInput.Substring(StrPre.Length);
                byte[] bytKey = GetKey(strKey, KeyType.KT_Key);
                byte[] bytIV = GetKey(strIv, KeyType.KT_IV);
                byte[] byteIn = HexStringToByte(strIn);
                byte[] byteOut = new byte[byteIn.Length];
                ZLSM4.sm4_crypt_cbc(Convert.ToInt32(CrypeMode.CM_Decrypt), byteIn.Length, bytIV, bytKey, byteIn, byteOut);
                strRet = Encoding.ASCII.GetString(byteOut);
                if (intVersion == 1)
                {
                    strRet = strRet.Trim();
                }
                else
                {
                    strRet = TruncZero(strRet);
                }

            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           Sm3                     计算字符串的哈希值（用来检测字符串的变动）
        //'返回值         String(32)              字符串的哈希值
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  字符串内容
        //'======================================================================================================================
        public static string Sm3(string strInput)
        {
            byte[] bytIn = Encoding.ASCII.GetBytes(strInput);
            byte[] bytOut = new byte[32];
            sm3_hash(bytIn, bytIn.Length, bytOut);
            string strRet = ByteToHexString(bytOut);
            return strRet;
        }
        //'======================================================================================================================
        //'方法           Sm3_File                计算文件的哈希值（用来检测 文件内容的变动）
        //'返回值         String(32)              文件的哈希值
        //'入参列表:
        //'参数名         类型                    说明
        //'strFile        String                  文件路径
        //'======================================================================================================================
        public static string Sm3_File(string strFile)
        {
            byte[] bytIn = Encoding.ASCII.GetBytes(strFile + "\0");
            byte[] bytOut = new byte[32];
            int intRet = sm3_file_hash(bytIn, bytOut);
            string strRet = null;
            switch (intRet)
            {
                case 0:
                    strRet = ByteToHexString(bytOut);
                    break;
                case 1:
                    strRet = "ERROR:文件打开失败";
                    break;
                case 2:
                    strRet = "ERROR:文件读取失败";
                    break;
                default:
                    strRet = "ERROR:未知错误";
                    break;
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           sm3_hmac                给定义一个密钥对传入的消息产生消息摘要
        //'返回值         String(32)              密钥加密消息后生成的消息摘要
        //'入参列表:
        //'参数名         类型                    说明
        //'strKey         String                  密钥
        //'strMsg         String                  消息内容
        //'======================================================================================================================
        public static string sm3_hmac(string strKey, string strMsg)
        {
            byte[] bytKey = Encoding.ASCII.GetBytes(strKey);
            byte[] bytMsg = Encoding.ASCII.GetBytes(strMsg);
            byte[] bytOut = new byte[32];
            sm3_hmac_hash(bytKey, bytKey.Length, bytMsg, bytMsg.Length, bytOut);
            string strRet = ByteToHexString(bytOut);
            return strRet;
        }
        //'======================================================================================================================
        //'方法           sm_version              获取ZLSM4的版本号
        //'返回值         Long                    ZLSM4的版本号
        //'入参列表:
        //'======================================================================================================================
        public static int sm_version()
        {
            int intRet;
            try
            {
                intRet = get_sm_version();
            }
            catch
            {
                intRet = 1;
            }
            return intRet;
        }
        //'======================================================================================================================
        //'方法           BytePadding             将指定字符串按照16字节补齐，
        //'返回值         Byte()                  补齐后的字符串字节组
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  字符串
        //'lngVersion     Long(Optional,2)        字符串补齐的版本（ZLSM4.DLL的版本，以及加密算法前缀中的版本），1-空格补齐，>1:Chr(0)补齐
        //'lngPaddingNum  Long(Optional,16)        补齐的字节数，缺省按照16进制补齐
        //'======================================================================================================================
        private static byte[] BytePadding(string strInput, int intVersion = 2, int intPaddingNum = 16)
        {

            byte[] byteIn = Encoding.ASCII.GetBytes(strInput);
            //计算填充的后的长度
            int intLen = ((int)(byteIn.Length / intPaddingNum) + ((byteIn.Length % intPaddingNum > 0) ? 1 : 0)) * intPaddingNum;
            byte[] byteNew = new byte[intLen];
            for (int i = 0; i < byteIn.Length; i++)
            {
                byteNew[i] = byteIn[i];
            }
            for (int i = byteIn.Length; i < intLen; i++)
            {
                if (intVersion > 1)
                {
                    byteNew[i] = 0;
                }
                else
                {
                    byteNew[i] = 32;
                }
            }
            return byteNew;
        }
        //'======================================================================================================================
        //'方法           ByteToHexString         将16进制字符串转换为字节组
        //'返回值         Byte()                  16进制字符串转换的字节组
        //'入参列表:
        //'参数名         类型                    说明
        //'bstrInput      String                  16进制字符串
        //'lngRetBytLen   Long(Optional)          指定返回的字节组的长度,0-按原始长度返回，<>0返回指定的长度，不足补齐（补0），多了截取
        //'======================================================================================================================
        private static byte[] HexStringToByte(string strInput, int intRetBytLen = 0)
        {
            byte[] bytRet = null;
            int lngLen = (int)(strInput.Length / 2);
            if (intRetBytLen != 0)
            {
                bytRet = new byte[intRetBytLen];
            }
            else
            {
                bytRet = new byte[lngLen];
            }
            for (int i = 0; i < bytRet.Length; i++)
            {
                if (i < lngLen)
                {
                    bytRet[i] = Convert.ToByte(("0x" + strInput.Substring(i * 2, 2)), 16);
                }
                else
                {
                    bytRet[i] = 0;
                }
            }
            return bytRet;
        }
        //'======================================================================================================================
        //'方法           ByteToHexString         将字节组转换为16进制字符串
        //'返回值         String                  字节组转换的16进制字符串
        //'入参列表:
        //'参数名         类型                    说明
        //'bytInpu        Byte(）                 字节数组
        //'======================================================================================================================
        private static string ByteToHexString(byte[] bytInput)
        {
            string strRet = null;
            for (int i = 0; i < bytInput.Length; i++)
            {
                strRet = strRet + bytInput[i].ToString("X2");
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           TruncZero              去掉字字符串中的\0以及其后的字符串
        //'返回值         String                 返回处理后的字符串
        //'入参列表:
        //'参数名         类型                    说明
        //'strInput       String                  传入的字符串
        //'======================================================================================================================
        private static string TruncZero(string strInput)
        {
            string strRet = strInput;
            int length = strInput.IndexOf('\0');
            if (length >= 0)
            {
                strRet = strInput.Substring(0, length);
            }
            return strRet;
        }
        //'======================================================================================================================
        //'方法           GetKey                 根据传入的16进制字符串生成密钥，没有密钥则自动生成一个
        //'返回值         byte()                 返回处理后的字符串
        //'入参列表:
        //'参数名         类型                    说明
        //'strKey         String                  传入的16进制字符串密钥
        //'intType        Integer                 没有密钥时生成密钥的类型，0-演示密钥，1-分组加密的分组加密密钥，2-加密密钥
        //'======================================================================================================================
        private static byte[] GetKey(string strKey, KeyType ktType)
        {
            byte[] bytRet = new byte[16];
            if (strKey != null)
            {
                bytRet = HexStringToByte(strKey, 16);
            }
            else
            {
                switch (ktType)
                {
                    case KeyType.KT_Default:
                        for (int i = 0; i < bytRet.Length; i++)
                        {
                            bytRet[i] = Convert.ToByte(i * 15);
                        }
                        break;
                    case KeyType.KT_IV:
                        bytRet = mbtyIVKey;
                        break;
                    case KeyType.KT_Key:
                        bytRet = mbtyKey;
                        break;
                    default:
                        for (int i = 0; i < bytRet.Length; i++)
                        {
                            bytRet[i] = Convert.ToByte(i * 15);
                        }
                        break;
                }
            }
            return bytRet;
        }
        internal static string GetGeneralAccountKey(string strKey)
        {
            byte[] bytTmp = new byte[16];
            if (strKey != null)
            {
                bytTmp = HexStringToByte(strKey, 16);
                for (int i = 0; i < bytTmp.Length; i++)
                {
                    if (i % 2 == 0)
                    {
                        bytTmp[i] = Convert.ToByte(Convert.ToInt32(255 - bytTmp[i]));
                    }
                    else if (i % 3 == 0)
                    {
                        bytTmp[i] = Convert.ToByte(Convert.ToInt32(bytTmp[i] + i) % 256);
                    }
                }
            }
            return ByteToHexString(bytTmp);
        }
    }
}
