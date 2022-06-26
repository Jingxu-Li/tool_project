using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Threading;
using System.IO;
using System.Data;
using System.Runtime.InteropServices;

namespace serial_new
{
    class Program
    {
        /// <summary>
        /// 将字符串写入txt文件
        /// </summary>
        /// <param name="message"></param>
        public static void FileWrite(string message)
        {
            FileStream _file = new FileStream(fileName, FileMode.Append, FileAccess.Write);
            using (StreamWriter writer1 = new StreamWriter(_file))
            {
                writer1.WriteLine(message);
                writer1.Flush();
                writer1.Close();

                _file.Close();
            }
        }

        /// <summary>
        /// Write DataTable to csv File
        /// </summary>
        /// <param name="dt">DataTable to be write</param>
        /// <param name="fullPath">csv path</param>
        public static void SaveCSV(DataTable dt, string fullPath)
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            System.IO.FileStream fs = new System.IO.FileStream(fullPath, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite);
            System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    data += str;

                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
        }

        /// <summary>
        /// Read CSV file and return DataTable
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public static DataTable OpenCSV(string filepath)
        {
            System.Text.Encoding encoding = System.Text.Encoding.UTF8;
            DataTable dt = new DataTable();
            //System.IO.FileStream fs = new System.IO.FileStream(filepath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
            System.IO.StreamReader sr = new System.IO.StreamReader(filepath, encoding);

            string strLine = "";
            string[] aryLine = null;
            string[] tableHead = null;
            int columnCount = 0;
            bool IsFirst = true;
            while ((strLine = sr.ReadLine()) != null)
            {
                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;

                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            //fs.Close();
            return dt;

        }

        /// <summary>
        /// 
        /// </summary>
        public static DataTable WriteDataTable(DataTable dt, string fullPath)
        {
            return dt;
        }

        /// <summary>
        /// 校验带有BCC的字符串
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static bool Bcccheck(string data)
        {
            bool result = true;//返回检验值
            string bcccheck = data.Substring(data.Length - 3, 2);//待检验BCC值
            byte[] dataToCheck = System.Text.Encoding.ASCII.GetBytes(data);//待检验字符串
            string calbcc;
            byte[] datacheck = new byte[data.Length - 3];
            //Console.WriteLine(bcccheck);
            Array.Copy(dataToCheck, datacheck, datacheck.Length);//将待检验字符串字节数组放到检验字符串字节数组中
            int len = datacheck.Length;
            byte check = 0;//检验结果
            for (int i = 0; i < len; i++)
            {
                check ^= datacheck[i];
            }
            calbcc = Convert.ToString(check, 16).ToUpper();
            if (calbcc.Length == 1)
                calbcc = '0' + calbcc;
            //Console.WriteLine(calbcc);
            if (!calbcc.Equals(bcccheck))
                result = false;
            return result;
        }

        /// <summary>
        /// 将未加bcc字符串末尾添加bcc校验值
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string GetCommandWithBcc(string data)
        {
            byte bcc = 0;//检验结果
            string check;
            byte[] dataToCheck = System.Text.Encoding.ASCII.GetBytes(data);
            int len = dataToCheck.Length;

            for (int i = 0; i < len; i++)
            {
                bcc ^= dataToCheck[i];
            }

            check = Convert.ToString(bcc, 16).ToUpper();
            if (check.Length == 1)
            {
                check = '0' + check;
            }
            return data + check;
        }

        /// <summary>
        /// 得到标签对应的PLC指令
        /// </summary>
        /// <param name="startAddr">PLC标签首地址,10000-59500</param>
        /// <param name="endAddr">PLC标签末地址,10487-59987</param>
        /// <returns>带有BCC的PLC指令字符串</returns>
        public static string GetCommand(int startAddr, int endAddr)
        {
            string command = "<01#RDD";
            string sa = Convert.ToString(startAddr);
            if (sa.Length < 5)
            {
                int num0 = 5 - sa.Length;
                while (num0 != 0)
                {
                    sa = "0" + sa;
                    num0--;
                }
            }
            string ea = Convert.ToString(endAddr);
            if (ea.Length < 5)
            {
                int num0 = 5 - ea.Length;
                while (num0 != 0)
                {
                    ea = "0" + ea;
                    num0--;
                }
            }

            command = command + sa + ea;
            command = GetCommandWithBcc(command) + Environment.NewLine;

            return command;
        }

        /// <summary>
        /// 时间同步PLC命令
        /// </summary>
        /// <returns></returns>
        public static string TimeSync()
        {
            string tscm = "%01#WDD0300003002";//Time Sync Command
            string datetime = DateTime.Now.ToString("yyMMddHHmmss");
            tscm = tscm + datetime;
            tscm = GetCommandWithBcc(tscm) + Environment.NewLine;
            return tscm;
        }

        /// <summary>
        /// 读取500寄存器和501寄存器的值，如果同时置为0则开始读取数据
        /// </summary>
        /// <returns></returns>
        public static bool ReadSingleContact500501()
        {
            bool re = false;
            bool re1 = false;
            bool re2 = false;

            string s500 = "<01#RCSR05000B" + Environment.NewLine;
            string s501 = "<01#RCSR05010A" + Environment.NewLine;

            //<01$RC0+BCC+CR
            string msg500 = CheckMessage(s500);
            string msg501 = CheckMessage(s501);

            if (Timeout)
                return false;

            if (msg500[6].Equals('0'))
            {
                re1 = true;
            }
            if (msg501[6].Equals('0'))
            {
                re2 = true;
            }

            if (re1 == true && re2 == true)
            {
                re = true;
            }

            return re;
        }

        /// <summary>
        /// test received data, repeat three times, if received correct data, return correct data,
        /// else receive error type
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="rec"></param>
        /// <returns></returns>
        public static string CheckMessage(string cmd)
        {
            string recv = null;
            //readover = false;
            if (!Port.IsOpen)
            {
                try
                {
                    Port.Open();
                }
                catch (Exception ex)
                {
                    if (!Timeout)
                    {
                        if (ex.Message != HardwareString)
                        {
                            Console.WriteLine(ex.Message);
                            HardwareString = ex.Message;
                            Timeout = true;
                        }
                    }
                    return ex.Message;
                }
            }
            try
            {
                Port.Write(cmd);
                Thread.Sleep(200);
                recv = Port.ReadExisting();
                if (recv.Contains('\r') && recv[3] == 36)
                {
                    Timeout = false;
                    //readover = true;
                    return recv;
                }
                else
                {
                    Timeout = true;
                    if (Port.IsOpen == false)
                    {
                        Console.WriteLine("上位机串口故障");
                        return "上位机串口故障";
                    }
                }

                for (int trytime = 0; trytime < 3; trytime++)
                {
                    Port.Write(cmd);
                    Thread.Sleep(3000);
                    recv = Port.ReadExisting();
                    //readover = true;
                    if (recv.Contains('\r') && recv[3] == 36)
                    {
                        Timeout = false;                        
                        return recv;
                    }
                }
                Timeout = true;
                Console.WriteLine("PLC与转换器间故障");
                return "PLC与转换器间故障";
            }
            catch (Exception ex)
            {
                Timeout = true;
                Console.WriteLine(ex.Message);
                return ex.Message;
            }
        }

        /// <summary>
        /// Post data to database
        /// </summary>
        /// <param name="Time"></param>
        /// <param name="Inout"></param>
        /// <param name="People"></param>
        /// <param name="Tool"></param>
        /// <returns>Transmit success or not</returns>
        public static bool Postdata(string Time,string Inout,string People,string Tool)
        {
            return true;
        }

        public static string message = null;
        public static string error;
        public static bool comerror = false;
        public static string fileName = String.Format(@"c:\logdir\{0}.txt", DateTime.Now.ToString("yyyyMMdd"));
        public static string csvName = String.Format(@"c:\datadir\{0}.csv", DateTime.Now.ToString("yyyyMMdd"));
        public static DataTable dt = new DataTable();

        public static SerialPort Port = new SerialPort("COM6", 115200, Parity.Odd, 8, StopBits.One);
        public static string WCS500 = "<01#WCSR050013F" + Environment.NewLine;
        public static string WCS501 = "<01#WCSR050113E" + Environment.NewLine;
        public static bool readover = false;
        public static string Time;
        public static string Inout;
        public static string People;
        public static string Tool;
        public static string Errorcode;
        public static string Time_ym;
        public static string Time_dh;
        public static string Time_ms;
        public static string Tagnum;
        public static int comflag = 0;
        public static int inout = 0;
        public static int ymcheck = 0;
        public static bool Check;
        public static bool systemerror = false;
        public static string systemstring;
        public static bool Timeout = false;
        public static string HardwareString = null;
        public static bool pause = false;
        public delegate bool ControlCtrlDelegate(int CtrlType);

        [DllImport("kernel32.dll")]
        private static extern bool SetConsoleCtrlHandler(ControlCtrlDelegate HandlerRoutine, bool Add);
        private static ControlCtrlDelegate cancelHandler = new ControlCtrlDelegate(HandlerRoutine);

        public static bool HandlerRoutine(int CtrlType)
        {
            pause = true;
            string Clear500 = "<01#WCSR050003E" + Environment.NewLine;
            string Clear501 = "<01#WCSR050103F" + Environment.NewLine;
            switch (CtrlType)
            {
                case 0:
                    {
                        string CancelMsg1 = CheckMessage(Clear500);
                        string CancelMsg2 = CheckMessage(Clear501);
                        Console.WriteLine("shutdown with Ctrl+C");
                        FileWrite("shutdown with Ctrl+C");
                        break;
                    }
                case 2:
                    {
                        if(!readover)
                        {
                            Thread.Sleep(2000);
                        }
                        Thread.Sleep(1000);
                        Port.Write(Clear500);
                        Thread.Sleep(1000);
                        Port.Write(Clear501);
                        Console.WriteLine("shutdown with Console");
                        FileWrite("shutdown with Console");
                        break;
                    }
            }
            return true;
        }

        static void Main(string[] args)
        {
            SetConsoleCtrlHandler(cancelHandler, true);
            //log directory test
            if (!Directory.Exists(@"c:\logdir"))
            {
                Directory.CreateDirectory(@"c:\logdir");
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                }
            }
            else
            {
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                }
            }
            //csv directory test
            if (!Directory.Exists(@"c:\datadir"))
            {
                Directory.CreateDirectory(@"c:\datadir");
                if (!File.Exists(csvName))
                {
                    File.Create(csvName).Close();
                }
            }
            else
            {
                if (!File.Exists(csvName))
                {
                    File.Create(csvName).Close();
                }
            }

            //dt.Columns.Add("ReadTime", Type.GetType("System.String"));
            //dt.Columns.Add("RFID_TYPE", Type.GetType("System.String"));
            //dt.Columns.Add("RFID", Type.GetType("System.String"));
            //dt.Columns.Add("IS_BUSINESS_PROCESSING", Type.GetType("System.String"));
            //dt.Columns.Add("DATE", Type.GetType("System.String"));

            Console.WriteLine("程序启动时间：{0}", DateTime.Now.ToString());
            FileWrite(string.Format("程序启动时间：{0}", DateTime.Now.ToString()));

            //Check if Port is open
            if (!Port.IsOpen)
            {
                try
                {
                    Port.Open();
                }
                catch (Exception ex)
                {
                    if (ex.Message != error)
                    {
                        Console.WriteLine("串口发生故障的时间：{0}", DateTime.Now.ToString());
                        Console.WriteLine("串口故障原因：{0}", ex.Message);
                        error = ex.Message;
                        FileWrite(string.Format("串口发生故障的时间：{0}", DateTime.Now.ToString()));
                        FileWrite(string.Format("串口故障原因：{0}", ex.Message));
                    }
                }
            }
            //Communication Test for the first time
            message = CheckMessage(GetCommand(10000, 10487));
            if (Timeout)
            {
                Console.WriteLine("通信测试阶段系统发生故障的时间：{0}", DateTime.Now.ToString());
                Console.WriteLine("通信测试阶段故障原因：{0}", message);
                FileWrite(string.Format("通信测试阶段系统发生故障的时间：{0}", DateTime.Now.ToString()));
                FileWrite(string.Format("通信测试阶段故障原因：{0}", message));
            }
            //Time sync
            if (!Timeout)
            {
                string timesSyncResult =  CheckMessage(TimeSync());
            }
            //Start main circle
            while (true)
            {
                message = CheckMessage(GetCommand(10000, 10487));
                //clear 500 and 501 circle
                if (!Timeout)
                {
                    while (true)
                    {
                        if (ReadSingleContact500501())
                        {
                            break;
                        }
                    }
                    string msgret501 = CheckMessage(WCS501);
                }
                //串口可以通信，500和501寄存器已经复位则读取数据并且写到文件中
                if ((!Timeout) && (!systemerror)&&(!comerror))
                {
                    try
                    {
                        for (int sendtime = 0; sendtime < 100; sendtime++)
                        {
                            readover = false;
                            message = CheckMessage(GetCommand(10000 + 500 * sendtime, 10000 + 500 * sendtime + 487));
                            if (message.Length != 0 && Timeout==false)
                            {
                                comerror = false;
                                Time = DateTime.Now.ToString();
                                if(pause)
                                    Thread.Sleep(5000);
                                if (message[3] != 36 || (Bcccheck(message) == false))
                                {
                                    Console.WriteLine("通讯数据错误（第三位不为$或BCC异常），接收到的数据为：{0}", message);
                                    FileWrite(string.Format("通讯数据缺失（第三位不为$或BCC异常），接收到的数据为：{0}", message));
                                    continue;
                                }

                                Console.WriteLine("当前数据条数：{0}", sendtime);

                                Check = Bcccheck(message);
                                Inout = message.Substring(34, 4);
                                inout = int.Parse(Inout);
                                Errorcode = message.Substring(6, 4);
                                Tagnum = message.Substring(18, 4);
                                Time_ym = message.Substring(22, 4);

                                ymcheck = int.Parse(Time_ym);
                                //Console.WriteLine(ymcheck);
                                //Console.WriteLine(inout);
                                Time_dh = message.Substring(26, 4);
                                Time_ms = message.Substring(30, 4);
                                People = message.Substring(38, 24);
                                Tool = message.Substring(62, 1896);
                                Thread.Sleep(50);
                            }
                            //analyze the data and break the circle
                            if (inout == 0 && ymcheck == 0)
                            {
                                break;
                            }
                            //Check inout
                            if (inout == 1)
                            {
                                Inout = "进库";
                            }
                            else if (inout == 2)
                            {
                                Inout = "出库";
                            }

                            if (inout == 1 || inout == 2)
                            {
                                FileWrite(string.Format("读取时间：{0} \n\r人员：{1} \n\r工具：{2} \n\rBCC校验：{3} \n\r出入库状态：{4} \n\r错误码：{5} \n\r记录时间：{6} \n\r", Time, People, Tool, Check, Inout, Errorcode,Time_ym + Time_dh + Time_ms));
                                Console.WriteLine("读取时间：{0}", Time);
                                Console.WriteLine("人员：{0}", People);
                                Console.WriteLine("工具：{0}", Tool);
                                Console.WriteLine("BCC校验：{0}", Check);
                                Console.WriteLine("出入库：{0}", Inout);
                                Console.WriteLine("记录时间：{0}", Time_ym + Time_dh + Time_ms);
                            }
                            readover = true;
                            Thread.Sleep(200);
                        }
                        string msgret500 = CheckMessage(WCS500);
                    }
                    catch (Exception ex)
                    {
                        systemstring = ex.Message;
                        if (systemerror == false)
                        {
                            Console.WriteLine("系统发生故障的时间：{0}", DateTime.Now.ToString());
                            Console.WriteLine("故障原因：{0}", systemstring);
                            FileWrite(string.Format("系统发生故障的时间：{0}", DateTime.Now.ToString()));
                            FileWrite(string.Format("故障原因：{0}", systemstring));
                            systemerror = true;
                        }
                    }
                }

            }
        }
    }
}

