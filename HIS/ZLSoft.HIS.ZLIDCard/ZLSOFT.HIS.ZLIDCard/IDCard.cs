using System;
using System.IO;
using System.Text;
using System.Runtime.ExceptionServices;

namespace ZLSOFT.HIS.ZLIDCard
{
    //取到身份证的值的类

    public class IDCard
    {
        public string ErrMeassge = "";
        public IDCard()
        {//构造函数初始化身份证读卡机
            ErrMeassge = "";
            int iPort, iRetUSB = 0;
            try
            {
                for (iPort = 1001; iPort <= 1016; iPort++)
                {
                    iRetUSB = CVRSDK.CVR_InitComm(iPort);
                    if (iRetUSB == 1) break;
                }
                if (iRetUSB != 1) ErrMeassge = "身份证读卡设备初始化失败！";
            }
            catch (Exception e)
            {
                ErrMeassge = "身份证读卡设备初始化失败！";
            }
        }

        [HandleProcessCorruptedStateExceptions] //.net4支持对非托管异常的捕获

        public PersonInfor ReadIDCard()//对身份证是否读取进行验证，并读取信息
        {
            ErrMeassge = "";
            try
            {
                int authenticate = CVRSDK.CVR_Authenticate();
                if (authenticate == 1)
                {
                    int readContent = CVRSDK.CVR_Read_FPContent();
                    if (readContent == 1)
                    {
                        return GetPersonInfor();
                    }
                    else
                    {
                        ErrMeassge="设备认证成功，但读卡失败！";
                    }
                }
                else
                {
                    ErrMeassge="请将身份证摆放到合适位置再试一次";
                }
            }
            catch (Exception ex)
            {
                ErrMeassge= "身份证信息读取失败！";
            }
            return new PersonInfor(); 
        }

        private PersonInfor GetPersonInfor()
        {
            PersonInfor personInfor = new PersonInfor();//实例化对象

            try
            {
                byte[] imgData = new byte[40960];//身份证图片封装

                int length = 40960;
                CVRSDK.GetJpgData(ref imgData[0], ref length);
                MemoryStream myStream = new MemoryStream();

                for (int i = 0; i < length; i++)
                {
                    myStream.WriteByte(imgData[i]);
                }
                /*Image myImage = Image.FromStream(myStream);解析图片
                pictureBoxPhoto.Image = myImage;*/

                byte[] name = new byte[128];
                length = 128;
                CVRSDK.GetPeopleName(ref name[0], ref length);

                byte[] cnName = new byte[128];
                length = 128;
                CVRSDK.GetPeopleChineseName(ref cnName[0], ref length);

                byte[] IDnumber = new byte[128];
                length = 128;
                CVRSDK.GetPeopleIDCode(ref IDnumber[0], ref length);

                byte[] peopleNation = new byte[128];
                length = 128;
                CVRSDK.GetPeopleNation(ref peopleNation[0], ref length);

                byte[] peopleNationCode = new byte[128];
                length = 128;
                CVRSDK.GetNationCode(ref peopleNationCode[0], ref length);

                byte[] validtermOfStart = new byte[128];
                length = 128;
                CVRSDK.GetStartDate(ref validtermOfStart[0], ref length);

                byte[] birthday = new byte[128];
                length = 128;
                CVRSDK.GetPeopleBirthday(ref birthday[0], ref length);

                byte[] address = new byte[128];
                length = 128;
                CVRSDK.GetPeopleAddress(ref address[0], ref length);

                byte[] validtermOfEnd = new byte[128];
                length = 128;
                CVRSDK.GetEndDate(ref validtermOfEnd[0], ref length);

                byte[] signdate = new byte[128];
                length = 128;
                CVRSDK.GetDepartment(ref signdate[0], ref length);

                byte[] sex = new byte[128];
                length = 128;
                CVRSDK.GetPeopleSex(ref sex[0], ref length);

                byte[] samid = new byte[128];
                CVRSDK.CVR_GetSAMID(ref samid[0]);

                bool bCivic = true;
                byte[] certType = new byte[32];
                length = 32;
                CVRSDK.GetCertType(ref certType[0], ref length);

                string strType = Encoding.ASCII.GetString(certType);
                int nStart = strType.IndexOf("I");
                if (nStart != -1) bCivic = false;

                personInfor.Sex = Encoding.GetEncoding("GB2312").GetString(sex).Replace("\0", "").Trim();//获取性别  
                personInfor.Birthday = Encoding.GetEncoding("GB2312").GetString(birthday).Replace("\0", "").Trim();//获取出生日期
                personInfor.Identity = Encoding.GetEncoding("GB2312").GetString(IDnumber).Replace("\0", "").Trim();//获取身份证号
                personInfor.Signdate = Encoding.GetEncoding("GB2312").GetString(signdate).Replace("\0", "").Trim();//获取签发机关
                personInfor.ValidtermOfStart = Encoding.GetEncoding("GB2312").GetString(validtermOfStart).Replace("\0", "").Trim();//获取身份证发证日期

                personInfor.ValidtermOfEnd = Encoding.GetEncoding("GB2312").GetString(validtermOfEnd).Replace("\0", "").Trim();//获取身份证失效日期" 
                personInfor.Samid = Encoding.GetEncoding("GB2312").GetString(samid).Replace("\0", "").Trim();//获取安全模块号

                personInfor.Picture = myStream;

                if (bCivic)
                {
                    personInfor.Name = Encoding.GetEncoding("GB2312").GetString(name).Replace("\0", "").Trim();//"获取姓名"                                                                                                
                    personInfor.Nation = Encoding.GetEncoding("GB2312").GetString(peopleNation).Replace("\0", "").Trim();//获取民族
                    personInfor.Address = Encoding.GetEncoding("GB2312").GetString(address).Replace("\0", "").Trim();//获取地址
                    personInfor.PeopleNation = "中国";
                }
                else
                {
                    personInfor.ForeigNername = Encoding.GetEncoding("GB2312").GetString(name).Replace("\0", "").Trim();//获得外国人本身姓名

                    personInfor.CnName = Encoding.GetEncoding("GB2312").GetString(cnName).Replace("\0", "").Trim();//获得中文姓名
                    personInfor.PeopleNation = Encoding.GetEncoding("GB2312").GetString(peopleNation).Replace("\0", "").Trim();//获得国籍
                    personInfor.PeopleNationCode = Encoding.GetEncoding("GB2312").GetString(peopleNationCode).Replace("\0", "").Trim();// "获得国籍代码
                }
                // return JsonConvert.SerializeObject(user);//返回json字符串

               
                return personInfor;
            }
            catch (Exception ex)
            {
                ErrMeassge = ex.ToString();
            }
            return personInfor;
        }
        ~IDCard()  // 析构函数关闭端口
        {
            try
            { 
                CVRSDK.CVR_CloseComm();
            }
            catch (Exception)
            {
            }
        }
    }
}
