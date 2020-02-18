using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;

namespace PickPosition
{
    class PickPosition
    {
        //宣告結束按鈕旗標
        public static short flagflag = 1;
        public static void Picking()
        {
            ScanCmdPacket ScanCmdPack = new ScanCmdPacket();
            ScanEchoPacket ScanEchoPack = new ScanEchoPacket();
            MachIDCmdPacket MachIDCmdPack = new MachIDCmdPacket();
            MachIDEchoPacket MachIDEchoPack = new MachIDEchoPacket();
            MachConnectCmdPacket MachConnectCmdPack = new MachConnectCmdPacket();
            MachConnectEchoPacket MachConnectEchoPack = new MachConnectEchoPacket();
            MachDataCmdPacket MachDataCmdPack = new MachDataCmdPacket();
            MachDataEchoPacket MachDataEchoPack = new MachDataEchoPacket();
            //MachConnectCmdPacket[] PassWord = new MachConnectCmdPacket[1];
            //MachDataCmdPacket[] DataBuf = new MachDataCmdPacket[1];
            MachConnectCmdPack.Password = new char[60];
            MachDataCmdPack.DataBuf = new byte[800];

            LaserPACKET LaserPack = new LaserPACKET(); //Data Package

            #region package initialize

            ScanCmdPack.ID = 0x0;
            ScanCmdPack.Sz = 0x0;
            ScanCmdPack.Cmd = 0x20;
            ScanCmdPack.Count = 0x0;
            ScanCmdPack.Sum = 0xe0;

            //   MachIDCmdPacket
            MachIDCmdPack.ID = 0x0;
            MachIDCmdPack.Sz = 0x0;
            MachIDCmdPack.Cmd = 0x21;
            MachIDCmdPack.Count = 0x0;
            MachIDCmdPack.Sum = 0xdf;

            //   MachConnectCmdPacket
            MachConnectCmdPack.ID = 1;
            MachConnectCmdPack.Sz = 0x4a;
            MachConnectCmdPack.Cmd = 0x22;
            MachConnectCmdPack.Count = 0;
            MachConnectCmdPack.DataSz = 0x42;
            MachConnectCmdPack.DataCmd0 = 0x03;
            MachConnectCmdPack.DataCmd1 = 0;
            MachConnectCmdPack.Part = 0;
            MachConnectCmdPack.ver1 = 4;
            MachConnectCmdPack.ver2 = 3;
            MachConnectCmdPack.BugFix = 7;
            MachConnectCmdPack.TypeID = 0x10;
            MachConnectCmdPack.SubTypeID = 0xa0;
            Array.Clear(MachConnectCmdPack.Password, 0x00, 60);
            for (int i = 0; i < 4; i++)
            {
                MachConnectCmdPack.Password[i] = '0';
            }
            //+ Array.ConvertAll<char, int>(MachConnectCmdPack.Password, value => Convert.ToInt32(value));
            MachConnectCmdPack.Sum = (byte)0xd0;

            //   MachDataCmdPacket
            MachDataCmdPack.ID = 1;
            MachDataCmdPack.Sz = 0x330;
            MachDataCmdPack.Cmd = 0x01;
            MachDataCmdPack.Count = 0x00;
            MachDataCmdPack.DataSz = 0x328;
            MachDataCmdPack.DataCmd0 = 0x50;
            MachDataCmdPack.DataCmd1 = 0;
            MachDataCmdPack.Part = 0;
            MachDataCmdPack.Code = 0xa000;
            MachDataCmdPack.Len = 0x320;
            //memset((char*)MachDataCmdPack.DataBuf, 0x00, sizeof(DataSent));
            Array.Clear(MachDataCmdPack.DataBuf, 0, 800);
            MachDataCmdPack.Sum = (byte)0x8d;
            Console.WriteLine(MachDataCmdPack.Sum);

            #endregion
            //PackageSetting(ref ScanCmdPack, ref MachIDCmdPack, ref MachConnectCmdPack, ref MachDataCmdPack);

            //設定本機IP位址，可用IPAddress.Any自行尋找，或直接使用IPAddress.Parse("X.X.X.X")設定
            IPAddress localAddress = IPAddress.Parse("10.1.10.210");
            IPAddress destAddress = IPAddress.Parse("10.1.10.200");
            ushort portNumber = 0x869C;
            try
            {
                localAddress = IPAddress.Parse(Form1.localIP);
                //如果為傳送端，使用IPAddress.Parse("X.X.X.X")設定
                destAddress = IPAddress.Parse(Form1.lasorIP);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.GetType().FullName);
                Console.WriteLine(ex.Message);
                Form1.textMessage += "\r\n" + ex.Message;
                Form1.ResetThread();
            }

            //*******************選擇是否為傳送端*******************//
            bool udpSender = true;

            //預設byte陣列數量
            int bufferSize = 512;

            //初始化物件
            UdpClient udpSocket = null; //UDP instance的宣告(類似初始化)
            byte[] sendBuffer = new byte[bufferSize], receiveBuffer = new byte[bufferSize];  //為了暫時存放序列化資料的變數
            int byteSize; //為顯示傳送出去封包大小而宣告的變數
            ushort cmdState = 0;

            List<string> mpBuffer = new List<string>(); //為一次存取所有資料的暫存變數
            long records = 0;
            

            try
            {    
                while (true)
                {
                    if(flagflag == 0)
                    {
                        break;
                    }

                    if (udpSender == false)
                    {
                        udpSocket = new UdpClient(new IPEndPoint(localAddress, portNumber));
                    }
                    else
                    {
                        udpSocket = new UdpClient(new IPEndPoint(localAddress, 0));
                    }


                    if (udpSender == true)
                    {
                        udpSocket.Connect(destAddress, portNumber);
                        Console.WriteLine("Connect() is OK...");
                    }

                    udpSocket.Client.ReceiveTimeout = 200; //等待訊息接收時間上限

                    if (udpSender == true)
                    {
                        if (cmdState == 0)
                        {
                            Console.WriteLine("Sending the first requested number of packets to the destination, Send()...");

                            MemoryStream stream = new MemoryStream();
                            BinaryWriter bw = new BinaryWriter(stream);
                            Console.WriteLine(ScanCmdPack.Sum);
                            bw.Write(ScanCmdPack.ID);
                            bw.Write(ScanCmdPack.Sz);
                            bw.Write(ScanCmdPack.Cmd);
                            bw.Write(ScanCmdPack.Count);
                            bw.Write(ScanCmdPack.Sum);

                            sendBuffer = stream.ToArray();
                            byteSize = udpSocket.Send(sendBuffer, sendBuffer.Length);
                            cmdState = (ushort)ME.scan;

                            Console.WriteLine("Sent {0} bytes to {1}", byteSize, destAddress.ToString());
                            stream.Close();
                            udpSender = false;
                            Console.WriteLine("Change to receiving mode");
                        }

                        if (cmdState == (ushort)ME.scanE)
                        {
                            Console.WriteLine("Sending the second requested number of packets to the destination, Send()...");

                            MemoryStream stream = new MemoryStream();
                            BinaryWriter bw = new BinaryWriter(stream);
                            Console.WriteLine(ScanCmdPack.Sum);
                            bw.Write(MachIDCmdPack.ID);
                            bw.Write(MachIDCmdPack.Sz);
                            bw.Write(MachIDCmdPack.Cmd);
                            bw.Write(MachIDCmdPack.Count);
                            bw.Write(MachIDCmdPack.Sum);
                            Console.WriteLine("sfgvafgv {0}", MachIDCmdPack.Sum);

                            sendBuffer = stream.ToArray();
                            byteSize = udpSocket.Send(sendBuffer, sendBuffer.Length);
                            cmdState = (ushort)ME.machid;

                            Console.WriteLine("Sent {0} bytes to {1}", byteSize, destAddress.ToString());
                            stream.Close();
                            udpSender = false;
                            Console.WriteLine("Change to receiving mode");
                        }

                        if (cmdState == (ushort)ME.machidE)
                        {
                            Console.WriteLine("Sending the third requested number of packets to the destination, Send()...");

                            MemoryStream stream = new MemoryStream();
                            BinaryWriter bw = new BinaryWriter(stream);
                            Console.WriteLine(ScanCmdPack.Sum);
                            bw.Write(MachConnectCmdPack.ID);
                            bw.Write(MachConnectCmdPack.Sz);
                            bw.Write(MachConnectCmdPack.Cmd);
                            bw.Write(MachConnectCmdPack.Count);
                            bw.Write(MachConnectCmdPack.DataSz);
                            bw.Write(MachConnectCmdPack.DataCmd0);
                            bw.Write(MachConnectCmdPack.DataCmd1);
                            bw.Write(MachConnectCmdPack.Part);
                            bw.Write(MachConnectCmdPack.ver1);
                            bw.Write(MachConnectCmdPack.ver2);
                            bw.Write(MachConnectCmdPack.BugFix);
                            bw.Write(MachConnectCmdPack.TypeID);
                            bw.Write(MachConnectCmdPack.SubTypeID);
                            //bw.Write(MachConnectCmdPack.Password);
                            bw.Write(MachConnectCmdPack.Password);
                            bw.Write(MachConnectCmdPack.Sum);

                            sendBuffer = stream.ToArray();
                            byteSize = udpSocket.Send(sendBuffer, sendBuffer.Length);
                            cmdState = (ushort)ME.machcon;

                            Console.WriteLine("Sent {0} bytes to {1}", byteSize, destAddress.ToString());
                            stream.Close();
                            udpSender = false;
                            Console.WriteLine("Change to receiving mode");
                        }

                        if (cmdState == (ushort)ME.machconE)
                        {
                            Form1.textMessage += "\r\nPicking...";
                            Console.WriteLine("Sending the forth requested number of packets to the destination, Send()...");

                            MemoryStream stream = new MemoryStream();
                            BinaryWriter bw = new BinaryWriter(stream);
                            Console.WriteLine(MachDataCmdPack.Sum);

                            bw.Write(MachDataCmdPack.ID);
                            bw.Write(MachDataCmdPack.Sz);
                            bw.Write(MachDataCmdPack.Cmd);
                            bw.Write(MachDataCmdPack.Count);
                            bw.Write(MachDataCmdPack.DataSz);
                            bw.Write(MachDataCmdPack.DataCmd0);
                            bw.Write(MachDataCmdPack.DataCmd1);
                            bw.Write(MachDataCmdPack.Part);
                            bw.Write(MachDataCmdPack.Code);
                            bw.Write(MachDataCmdPack.Len);
                            //bw.Write(MachDataCmdPack.DataBuf);
                            bw.Write(MachDataCmdPack.DataBuf);
                            bw.Write(MachDataCmdPack.Sum);


                            sendBuffer = stream.ToArray();
                            Console.WriteLine(sendBuffer);
                            byteSize = udpSocket.Send(sendBuffer, sendBuffer.Length);
                            cmdState = (ushort)ME.machdata;

                            Console.WriteLine("Sent {0} bytes to {1}", byteSize, destAddress.ToString());
                            stream.Close();
                            udpSender = false;
                            Console.WriteLine("Change to receiving mode");
                        }

                        if (cmdState == (ushort)ME.machdataE)
                        {
                            Console.WriteLine("Receive successfully, keep receiving.");
                            cmdState = (ushort)ME.machconE;
                        }
                        udpSocket.Close();
                    }
                    else
                    {
                        IPEndPoint senderEndPoint = new IPEndPoint(localAddress, 0);
                        Console.WriteLine("Receiving datagrams in a little time...");

                        if (cmdState == (ushort)ME.scan)
                        {
                            try
                            {
                                receiveBuffer = udpSocket.Receive(ref senderEndPoint);
                                cmdState = (ushort)ME.scanE;
                                Console.WriteLine("It is {0} bytes from {1}", receiveBuffer.Length, senderEndPoint.ToString());

                                MemoryStream stream = new MemoryStream(receiveBuffer);
                                BinaryReader br = new BinaryReader(stream);

                                ScanEchoPack.ID = BitConverter.ToUInt16(br.ReadBytes(2), 0);//反序列化
                                Console.WriteLine("Receive ID = {0}", ScanEchoPack.ID);

                                ScanEchoPack.Sz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Sz = {0}", ScanEchoPack.Sz);

                                ScanEchoPack.Cmd = br.ReadByte();
                                Console.WriteLine("Receive Cmd = {0}", ScanEchoPack.Cmd);

                                ScanEchoPack.Count = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Count = {0}", ScanEchoPack.Count);

                                ScanEchoPack.Sum = br.ReadByte();
                                Console.WriteLine("Receive Sum = {0}", ScanEchoPack.Sum);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.GetType().FullName);
                                Console.WriteLine(ex.Message);
                                Form1.textMessage += "\r\n" + ex.Message;
                                cmdState = 0;
                            }
                            udpSender = true;
                        }

                        if (cmdState == (ushort)ME.machid)
                        {
                            try
                            {
                                receiveBuffer = udpSocket.Receive(ref senderEndPoint);

                                cmdState = (ushort)ME.machidE;
                                Console.WriteLine("It is {0} bytes from {1}", receiveBuffer.Length, senderEndPoint.ToString());

                                MemoryStream stream = new MemoryStream(receiveBuffer);
                                BinaryReader br = new BinaryReader(stream);

                                MachIDEchoPack.ID = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive ID = {0}", MachIDEchoPack.ID);

                                MachIDEchoPack.Sz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Sz = {0}", MachIDEchoPack.Sz);

                                MachIDEchoPack.Cmd = br.ReadByte();
                                Console.WriteLine("Receive Cmd = {0}", MachIDEchoPack.Cmd);

                                MachIDEchoPack.Count = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Count = {0}", MachIDEchoPack.Count);

                                MachIDEchoPack.ID0 = br.ReadByte();
                                Console.WriteLine("Receive ID0 = {0}", MachIDEchoPack.ID0);

                                MachIDEchoPack.Ver1 = br.ReadByte();
                                Console.WriteLine("Receive Ver1 = {0}", MachIDEchoPack.Ver1);

                                MachIDEchoPack.Ver2 = br.ReadByte();
                                Console.WriteLine("Receive Ver2 = {0}", MachIDEchoPack.Ver2);

                                MachIDEchoPack.BugFix = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive BugFix = {0}", MachIDEchoPack.BugFix);

                                MachIDEchoPack.TypeID = br.ReadByte();
                                Console.WriteLine("Receive TypeID = {0}", MachIDEchoPack.TypeID);

                                MachIDEchoPack.SubTypeID = br.ReadByte();
                                Console.WriteLine("Receive SubTypeID = {0}", MachIDEchoPack.SubTypeID);

                                MachIDEchoPack.UserDef = br.ReadBytes(60);
                                Console.WriteLine("Receive UserDef");

                                MachIDEchoPack.Sum = br.ReadByte();
                                Console.WriteLine("Receive Sum = {0}", MachIDEchoPack.Sum);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.GetType().FullName);
                                Console.WriteLine(ex.Message);
                                Form1.textMessage += "\r\n" + ex.Message;
                                cmdState = (ushort)ME.scanE;
                            }
                            udpSender = true;
                        }

                        if (cmdState == (ushort)ME.machcon)
                        {
                            try
                            {
                                receiveBuffer = udpSocket.Receive(ref senderEndPoint);
                                cmdState = (ushort)ME.machconE;
                                Console.WriteLine("It is {0} bytes from {1}", receiveBuffer.Length, senderEndPoint.ToString());

                                MemoryStream stream = new MemoryStream(receiveBuffer);
                                BinaryReader br = new BinaryReader(stream);

                                MachConnectEchoPack.ID = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive ID = {0}", MachConnectEchoPack.ID);

                                MachConnectEchoPack.Sz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Sz = {0}", MachConnectEchoPack.Sz);

                                MachConnectEchoPack.Cmd = br.ReadByte();
                                Console.WriteLine("Receive Cmd = {0}", MachConnectEchoPack.Cmd);

                                MachConnectEchoPack.Count = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Count = {0}", MachConnectEchoPack.Count);

                                MachConnectEchoPack.DataSz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive DataSz = {0}", MachConnectEchoPack.DataSz);

                                MachConnectEchoPack.DataCmd0 = br.ReadByte();
                                Console.WriteLine("Receive DataCmd0 = {0}", MachConnectEchoPack.DataCmd0);

                                MachConnectEchoPack.DataCmd1 = br.ReadByte();
                                Console.WriteLine("Receive DataCmd1 = {0}", MachConnectEchoPack.DataCmd1);

                                MachConnectEchoPack.Part = BitConverter.ToUInt16(br.ReadBytes(4), 0);
                                Console.WriteLine("Receive Part = {0}", MachConnectEchoPack.Part);

                                MachConnectEchoPack.Security = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Security = {0}", MachConnectEchoPack.Security);

                                MachConnectEchoPack.MachID = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive MachID = {0}", MachConnectEchoPack.MachID);

                                MachConnectEchoPack.Sum = br.ReadByte();
                                Console.WriteLine("Receive Sum = {0}", MachIDEchoPack.Sum);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.GetType().FullName);
                                Console.WriteLine(ex.Message);
                                Form1.textMessage += "\r\n" + ex.Message;
                                cmdState = (ushort)ME.machidE;
                            }
                            udpSender = true;
                        }

                        if (cmdState == (ushort)ME.machdata)
                        {
                            try
                            {
                                receiveBuffer = udpSocket.Receive(ref senderEndPoint);
                                cmdState = (ushort)ME.machdataE;
                                Console.WriteLine("It is {0} bytes from {1}", receiveBuffer.Length, senderEndPoint.ToString());

                                MemoryStream stream = new MemoryStream(receiveBuffer);
                                BinaryReader br = new BinaryReader(stream);

                                MachDataEchoPack.ID = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive ID = {0}", MachDataEchoPack.ID);

                                MachDataEchoPack.Sz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Sz = {0}", MachDataEchoPack.Sz);

                                MachDataEchoPack.Cmd = br.ReadByte();
                                Console.WriteLine("Receive Cmd = {0}", MachDataEchoPack.Cmd);

                                MachDataEchoPack.Count = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive Count = {0}", MachDataEchoPack.Count);

                                MachDataEchoPack.DataSz = BitConverter.ToUInt16(br.ReadBytes(2), 0);
                                Console.WriteLine("Receive DataSz = {0}", MachDataEchoPack.DataSz);

                                MachDataEchoPack.DataCmd0 = br.ReadByte();
                                Console.WriteLine("Receive DataCmd0 = {0}", MachDataEchoPack.DataCmd0);

                                MachDataEchoPack.DataCmd1 = br.ReadByte();
                                Console.WriteLine("Receive DataCmd1 = {0}", MachDataEchoPack.DataCmd1);

                                MachDataEchoPack.Part = BitConverter.ToUInt32(br.ReadBytes(4), 0);
                                Console.WriteLine("Receive Part = {0}", MachDataEchoPack.Part);

                                MachDataEchoPack.Code = BitConverter.ToUInt32(br.ReadBytes(4), 0);
                                Console.WriteLine("Receive Code = {0}", MachDataEchoPack.Code);

                                MachDataEchoPack.Len = BitConverter.ToUInt32(br.ReadBytes(4), 0);
                                Console.WriteLine("Receive Len = {0}", MachDataEchoPack.Len);

                                MachDataEchoPack.ActctLen = BitConverter.ToUInt32(br.ReadBytes(4), 0);
                                Console.WriteLine("Receive ActctLen = {0}", MachDataEchoPack.ActctLen);

                                MachDataEchoPack.DataBuf = br.ReadBytes(532);
                                Console.WriteLine("Receive DataBuf");

                                MachDataEchoPack.Sum = br.ReadByte();
                                Console.WriteLine("Receive Sum = {0}", MachDataEchoPack.Sum);

                                LaserPack = (LaserPACKET)ByteToStruct(MachDataEchoPack.DataBuf, typeof(LaserPACKET));
                                
                                byte start = 0;
                                if (MachDataEchoPack.DataBuf[96] == 0xFC) { start = 1; } else { start = 0; }
                                mpBuffer.Add(
                                    DateTime.Now.ToString("HH: mm:ss.ffffzzz") + "," +
                                    BitConverter.ToInt32(MachDataEchoPack.DataBuf, 100).ToString() + "," +
                                    BitConverter.ToInt32(MachDataEchoPack.DataBuf, 104).ToString() + "," +
                                    BitConverter.ToInt32(MachDataEchoPack.DataBuf, 108).ToString() + "," +
                                    BitConverter.ToDouble(MachDataEchoPack.DataBuf, 316).ToString() + "," +
                                    start.ToString()
                                    );
                                
                                Thread.Sleep(10);
                                records++;
                                if (records % 500 == 0)
                                {
                                    Form1.textMessage += "\r\nPicking...(have picked above " + records.ToString() + " records)";
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.GetType().FullName);
                                Console.WriteLine(ex.Message);
                                Form1.textMessage += "\r\n" + ex.Message;
                                cmdState = (ushort)ME.machconE;
                            }
                            udpSender = true;
                        }
                        udpSocket.Close();
                    }
                }
            }
            catch (SocketException err)
            {
                Console.WriteLine("Socket error occurred: {0}", err.Message);
                Console.WriteLine("Stack: {0}", err.StackTrace);
                Form1.textMessage += "\r\nIP or Port " + err.Message;
                Form1.ResetThread();
            }
            catch (Exception err)
            {
                Form1.textMessage += "\r\n" +　err.Message;
                Form1.ResetThread();
            }
            finally
            {
                if (udpSocket != null)
                {
                    // Free up the underlying network resources
                    Console.WriteLine("Closing the socket...");
                    // udpSocket.Close();
                }
                //初始化檔案
                try
                {
                    string path = @Form1.localPath;
                    if (!Directory.Exists(Path.GetDirectoryName(path)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(path));
                    }
                    FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
                    TextWriter sw = new StreamWriter(fs);

                    sw.WriteLine("Time,x,y,z,Speed,Cycle Start");
                    foreach (string eachLine in mpBuffer)
                    {
                        sw.WriteLine("{0}", eachLine);
                    }
                    sw.Close();
                    Form1.textMessage += "\r\nTotal records: " + records.ToString();
                }
                catch (Exception ex)
                {
                    if (!Directory.Exists(Path.GetDirectoryName(@"C:\temp\temp.csv")))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(@"C:\temp\temp.csv"));
                    }
                    FileStream fs = new FileStream(@"C:\temp\temp.csv", FileMode.OpenOrCreate);
                    TextWriter sw = new StreamWriter(fs);

                    sw.WriteLine("Time,x,y,z,Speed,Cycle Start");
                    foreach (string eachLine in mpBuffer)
                    {
                        sw.WriteLine("{0}", eachLine);
                    }
                    sw.Close();
                    
                    Form1.textMessage += "\r\n" + ex.Message + ", 已先暫存到C:\\temp\\temp.csv";
                    Form1.textMessage += "\r\nTotal records: " + records.ToString();
                }
                
            }
        }

        public static object ByteToStruct(byte[] bytes, Type type)
        {

            int size = Marshal.SizeOf(type);
            Console.WriteLine("LaserPACKET size = " + size);
            Console.WriteLine("DataBuf size = " + bytes.Length);

            //分配結構體記憶體空間
            IntPtr structPtr = Marshal.AllocHGlobal(size);
            //將byte陣列拷貝到分配好的記憶體空間
            Marshal.Copy(bytes, 0, structPtr, size);
            try
            {
                //將記憶體空間轉換為目標結構體
                object obj = Marshal.PtrToStructure(structPtr, type);
                return obj;
            }
            finally
            {
                //釋放記憶體空間
                Marshal.FreeHGlobal(structPtr);
            }

        }
    }
}
