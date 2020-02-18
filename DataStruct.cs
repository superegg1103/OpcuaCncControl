using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

//宣告各項參數(不賦值)
namespace PickPosition
{
    [Serializable] //序列化
    [StructLayout(LayoutKind.Sequential, Pack = 2)] 
    public struct ScanCmdPacket
    {
        //外部
        public ushort ID;
        public ushort Sz;
        public byte Cmd; //its type is char originally
        public ushort Count;
        public byte Sum; //its type is char originally
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct ScanEchoPacket
    {
        public ushort ID;
        public ushort Sz;
        public byte Cmd;
        public ushort Count;
        public byte Sum; //its type is char originally
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    public struct MachIDCmdPacket
    {
        public ushort ID;
        public ushort Sz;
        public byte Cmd;   //its type is char originally
        public short Count;
        public int Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct MachIDEchoPacket
    {
        public ushort ID;
        public ushort Sz;
        public byte Cmd;        //its type is char originally
        public ushort Count;
        public byte ID0;        //its type is char originally
        public byte Ver1;       //its type is char originally
        public byte Ver2;       //its type is char originally
        public ushort BugFix;
        public byte TypeID;
        public byte SubTypeID;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 60)]
        public byte[] UserDef;
        public byte Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct MachConnectCmdPacket
    {
        public ushort ID;
        public ushort Sz;
        public byte Cmd;
        public ushort Count;
        public ushort DataSz;
        public byte DataCmd0;
        public byte DataCmd1;
        public uint Part;
        public byte ver1;
        public byte ver2;
        public ushort BugFix;
        public byte TypeID;
        public byte SubTypeID;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 60)]
        public char[] Password;
        public byte Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct MachConnectEchoPacket
    {
        public int ID;
        public int Sz;
        public int Cmd;
        public int Count;
        public int DataSz;
        public int DataCmd0;
        public int DataCmd1;
        public int Part;
        public int Security;
        public int MachID;
        public int Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct MachDataCmdPacket
    {
        public ushort ID;
        public ushort Sz;
        public byte Cmd;
        public ushort Count;
        public ushort DataSz;
        public byte DataCmd0;
        public byte DataCmd1;
        public uint Part;
        public uint Code;
        public uint Len;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 800)]
        public byte[] DataBuf;
        public byte Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct MachDataEchoPacket
    {
        public int ID;
        public int Sz;
        public int Cmd;
        public int Count;
        public int DataSz;
        public int DataCmd0;
        public int DataCmd1;
        public uint Part;
        public uint Code;
        public uint Len;
        public uint ActctLen;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 532)]
        public byte[] DataBuf;
        public byte Sum;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct LaserPACKET  //雷射封包結構體
    {
        public uint DateTime;   //系統時間
        public uint TotalWCount;   //累計工作時間
        public uint ThisWCount;   //目前加工時間
        public byte Unit;
        public byte CStart;
        public byte CStop;
        public byte MemStart;  //MemStart
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
        public char[] Fn;   //加工程式名稱 FileName
        public uint Nn, Bn;
        public ushort MMode;   //系統模式
        public ushort MStatus;   //機台狀態
        public ushort SCode, Ms;   //抓資料的兩個變數
                                   //------------------------------
        public LaserCOORD Coord;
        public LaserSDATA SData;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct LaserSDATA//<64Byte> //雷射本身加工參數
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public char[] Flag;
        public byte Mode;      //CW\QCW
        public byte Gas;       //Air,N2,O2
        public ushort Power;     //0~10000 unit:0.01%
        public ushort ResW;     //不理他
        public uint Dura;     //0~50000 unit:0.001ms 佔空比
        public uint Hz;       //0~5000000 unit:0.01hz
        public uint RefHt;    //pulse
        public uint Dead;     //pulse
        public ushort AirBar;       //0~50000 unit:0.001bar
        public ushort Kp;        //0~50000 unit:0.001 
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
        public char[] Comment;  //comment
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 28)]
        public char[] ResvB1;
    }

    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct LaserCOORD  //各軸參數
    {
        public PT9L Mp, Pp, Lp, Dp;                     //座標群
        public PT9L RefWp, OftWp;
        public double Speed;                          //移動速度
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public uint[] Rpm;      //各軸轉速
    }
    //---------------------------------
    [Serializable]
    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public struct PT9L
    {
        public int x, y, z, a, b, c, u, v, w;
    }

    public enum ME
    {
        Yes = 1,
        scan = 1,
        scanE = 2,
        machid = 3,
        machidE = 4,
        machcon = 5,
        machconE = 6,
        machdata = 7,
        machdataE = 8,
        ConnectEnd = 9
    }

}