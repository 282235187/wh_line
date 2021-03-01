#define Debug
#undef  Debug     //屏蔽后，打开工装测试，不屏蔽关闭工装测试
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;

using System.Threading;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace synthetic
{
    public partial class Form1 : Form
    {
        string address1, address2, address3;
        long SIZE1, SIZE2;
        long num = 0;
        int MODE = 0;
        int mode = 0;
        int Version_num;

        byte[] read1 = new byte[102400];   //49152
        byte[] read2 = new byte[102400];
        byte[] read3 = new byte[102400];   //升级文件数组
        byte[] write = new byte[204800];
        byte[] CR = new byte[4];

        UInt32[] crc32_table = new UInt32[256]{
                     0x00000000,  0x77073096,  0xEE0E612C,  0x990951BA,
                     0x076DC419,  0x706AF48F,  0xE963A535,  0x9E6495A3,
                     0x0EDB8832,  0x79DCB8A4,  0xE0D5E91E,  0x97D2D988,
                     0x09B64C2B,  0x7EB17CBD,  0xE7B82D07,  0x90BF1D91,
                     // 0x10
                     0x1DB71064,  0x6AB020F2,  0xF3B97148,  0x84BE41DE,
                     0x1ADAD47D,  0x6DDDE4EB,  0xF4D4B551,  0x83D385C7,
                     0x136C9856,  0x646BA8C0,  0xFD62F97A,  0x8A65C9EC,
                     0x14015C4F,  0x63066CD9,  0xFA0F3D63,  0x8D080DF5,
                     // 0x20
                     0x3B6E20C8,  0x4C69105E,  0xD56041E4,  0xA2677172,
                     0x3C03E4D1,  0x4B04D447,  0xD20D85FD,  0xA50AB56B,
                     0x35B5A8FA,  0x42B2986C,  0xDBBBC9D6,  0xACBCF940,
                     0x32D86CE3,  0x45DF5C75,  0xDCD60DCF,  0xABD13D59,
                     // 0x30
                     0x26D930AC,  0x51DE003A,  0xC8D75180,  0xBFD06116,
                     0x21B4F4B5,  0x56B3C423,  0xCFBA9599,  0xB8BDA50F,
                     0x2802B89E,  0x5F058808,  0xC60CD9B2,  0xB10BE924,
                     0x2F6F7C87,  0x58684C11,  0xC1611DAB,  0xB6662D3D,
                     // 0x40
                     0x76DC4190,  0x01DB7106,  0x98D220BC,  0xEFD5102A,
                     0x71B18589,  0x06B6B51F,  0x9FBFE4A5,  0xE8B8D433,
                     0x7807C9A2,  0x0F00F934,  0x9609A88E,  0xE10E9818,
                     0x7F6A0DBB,  0x086D3D2D,  0x91646C97,  0xE6635C01,
                     // 0x50
                     0x6B6B51F4,  0x1C6C6162,  0x856530D8,  0xF262004E,
                     0x6C0695ED,  0x1B01A57B,  0x8208F4C1,  0xF50FC457,
                     0x65B0D9C6,  0x12B7E950,  0x8BBEB8EA,  0xFCB9887C,
                     0x62DD1DDF,  0x15DA2D49,  0x8CD37CF3,  0xFBD44C65,
                     // 0x60
                     0x4DB26158,  0x3AB551CE,  0xA3BC0074,  0xD4BB30E2,
                     0x4ADFA541,  0x3DD895D7,  0xA4D1C46D,  0xD3D6F4FB,
                     0x4369E96A,  0x346ED9FC,  0xAD678846,  0xDA60B8D0,
                     0x44042D73,  0x33031DE5,  0xAA0A4C5F,  0xDD0D7CC9,
                     // 0x70
                     0x5005713C,  0x270241AA,  0xBE0B1010,  0xC90C2086,
                     0x5768B525,  0x206F85B3,  0xB966D409,  0xCE61E49F,
                     0x5EDEF90E,  0x29D9C998,  0xB0D09822,  0xC7D7A8B4,
                     0x59B33D17,  0x2EB40D81,  0xB7BD5C3B,  0xC0BA6CAD,
                    // 0x80
                    0xEDB88320,  0x9ABFB3B6,  0x03B6E20C,  0x74B1D29A,
                    0xEAD54739,  0x9DD277AF,  0x04DB2615,  0x73DC1683,
                    0xE3630B12,  0x94643B84,  0x0D6D6A3E,  0x7A6A5AA8,
                    0xE40ECF0B,  0x9309FF9D,  0x0A00AE27,  0x7D079EB1,
                    // 0x90
                    0xF00F9344,  0x8708A3D2,  0x1E01F268,  0x6906C2FE,
                    0xF762575D,  0x806567CB,  0x196C3671,  0x6E6B06E7,
                    0xFED41B76,  0x89D32BE0,  0x10DA7A5A,  0x67DD4ACC,
                    0xF9B9DF6F,  0x8EBEEFF9,  0x17B7BE43,  0x60B08ED5,
                    // 0xA0
                    0xD6D6A3E8,  0xA1D1937E,  0x38D8C2C4,  0x4FDFF252,
                    0xD1BB67F1,  0xA6BC5767,  0x3FB506DD,  0x48B2364B,
                    0xD80D2BDA,  0xAF0A1B4C,  0x36034AF6,  0x41047A60,
                    0xDF60EFC3,  0xA867DF55,  0x316E8EEF,  0x4669BE79,
                    // 0xB0
                    0xCB61B38C,  0xBC66831A,  0x256FD2A0,  0x5268E236,
                    0xCC0C7795,  0xBB0B4703,  0x220216B9,  0x5505262F,
                    0xC5BA3BBE,  0xB2BD0B28,  0x2BB45A92,  0x5CB36A04,
                    0xC2D7FFA7,  0xB5D0CF31,  0x2CD99E8B,  0x5BDEAE1D,
                    // 0xC0
                    0x9B64C2B0,  0xEC63F226,  0x756AA39C,  0x026D930A,
                    0x9C0906A9,  0xEB0E363F,  0x72076785,  0x05005713,
                    0x95BF4A82,  0xE2B87A14,  0x7BB12BAE,  0x0CB61B38,
                    0x92D28E9B,  0xE5D5BE0D,  0x7CDCEFB7,  0x0BDBDF21,
                    // 0xD0
                    0x86D3D2D4,  0xF1D4E242,  0x68DDB3F8,  0x1FDA836E,
                    0x81BE16CD,  0xF6B9265B,  0x6FB077E1,  0x18B74777,
                    0x88085AE6,  0xFF0F6A70,  0x66063BCA,  0x11010B5C,
                    0x8F659EFF,  0xF862AE69,  0x616BFFD3,  0x166CCF45,
                    // 0xE0
                    0xA00AE278,  0xD70DD2EE,  0x4E048354,  0x3903B3C2,
                    0xA7672661,  0xD06016F7,  0x4969474D,  0x3E6E77DB,
                    0xAED16A4A,  0xD9D65ADC,  0x40DF0B66,  0x37D83BF0,
                    0xA9BCAE53,  0xDEBB9EC5,  0x47B2CF7F,  0x30B5FFE9,
                    // 0xF0
                    0xBDBDF21C,  0xCABAC28A,  0x53B39330,  0x24B4A3A6,
                    0xBAD03605,  0xCDD70693,  0x54DE5729,  0x23D967BF,
                    0xB3667A2E,  0xC4614AB8,  0x5D681B02,  0x2A6F2B94,
                    0xB40BBE37,  0xC30C8EA1,  0x5A05DF1B,  0x2D02EF8D
                    };
        const UInt32 readCTWorkCfg = 0x07000001;                           //获取CT互感器工况信息
        const UInt32 readCTVerCfg = 0x07000002;                           //读取CT互感器版本
        const UInt32 readCoreVerCfg = 0x07000003;                           //读取算法核心板版本
        const UInt32 writeCTfileCfg = 0x0f000001;                           //写CT互感器文件
        const UInt32 writeIDCfg = 0x07020005;                           //设置巡检仪核心算法板ID
        const UInt32 readIDCfg = 0x07020006;                           //读取巡检仪核心算法板ID
        const UInt32 writeCTIDCfg = 0x07020003;                           //设置CT互感器ID 
        const UInt32 readCTIDCfg = 0x07020004;                           //读取CT互感器ID 

        public Form1()
        {
            InitializeComponent();
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived); //串口数据接收事件
            serialPort2.DataReceived += new SerialDataReceivedEventHandler(serialPort2_DataReceived); //串口数据接收事件
            serialPort3.DataReceived += new SerialDataReceivedEventHandler(serialPort3_DataReceived); //串口数据接收事件
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.AutoScrollMinSize = new Size(300, 250);
            this.BackColor = Color.FromArgb(25, 50, 100);
            comboBox1.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
            textBox41.Text = "0";
            textBox46.Text = "1";
            comboBox2.Text = "9600";//波特率默认值
            //加载停止位
            comboBox3.Items.Add("0");
            comboBox3.Items.Add("1");
            comboBox3.Items.Add("1.5");
            comboBox3.Items.Add("2");
            comboBox3.SelectedIndex = 1;
            //加载数据位
            comboBox5.Items.Add("8");
            comboBox5.Items.Add("7");
            comboBox5.Items.Add("6");
            comboBox5.Items.Add("5");
            comboBox5.SelectedIndex = 0;
            //加载奇偶校验位
            comboBox4.Items.Add("无");
            comboBox4.Items.Add("奇校验");
            comboBox4.Items.Add("偶校验");
            comboBox4.SelectedIndex = 0;
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);//添加事件处理程序

            comboBox24.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
            comboBox25.Text = "115200";//波特率默认值
            comboBox15.Text = "57.7";//电压限量初值
            comboBox16.Text = "5";//电流限量初值
            //加载停止位
            comboBox23.Items.Add("0");
            comboBox23.Items.Add("1");
            comboBox23.Items.Add("1.5");
            comboBox23.Items.Add("2");
            comboBox23.SelectedIndex = 1;
            //加载数据位
            comboBox22.Items.Add("8");
            comboBox22.Items.Add("7");
            comboBox22.Items.Add("6");
            comboBox22.Items.Add("5");
            comboBox22.SelectedIndex = 0;
            //加载奇偶校验位
            comboBox21.Items.Add("无");
            comboBox21.Items.Add("奇校验");
            comboBox21.Items.Add("偶校验");
            comboBox21.SelectedIndex = 0;
            serialPort2.DataReceived += new SerialDataReceivedEventHandler(serialPort2_DataReceived);//添加事件处理程序

            comboBox30.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
            comboBox29.Text = "9600";//波特率默认值
            numericUpDown1.Text = "5000";
            numericUpDown2.Text = "20000";
            timer1.Stop();
            //加载停止位
            comboBox28.Items.Add("0");
            comboBox28.Items.Add("1");
            comboBox28.Items.Add("1.5");
            comboBox28.Items.Add("2");
            comboBox28.SelectedIndex = 1;
            //加载数据位
            comboBox27.Items.Add("8");
            comboBox27.Items.Add("7");
            comboBox27.Items.Add("6");
            comboBox27.Items.Add("5");
            comboBox27.SelectedIndex = 0;
            //加载奇偶校验位
            comboBox26.Items.Add("无");
            comboBox26.Items.Add("奇校验");
            comboBox26.Items.Add("偶校验");
            comboBox26.SelectedIndex = 2;
            checkBox1.Checked = true;
            serialPort3.DataReceived += new SerialDataReceivedEventHandler(serialPort3_DataReceived);//添加事件处理程序
        }

        /// <summary>
        /// hex转换成string
        /// </summary>
        /// <param name="hex"></param>
        /// <returns></returns>
        private string HexToString(Int32 data)
        {
            string str = data.ToString("X2");
            return str;
        }
        /// <summary>
        /// 数组转换成字符串
        /// </summary>
        /// <param name="data"></param>
        /// <param name="datalen"></param>
        /// <returns></returns>
        private string ArrayTosString(byte[] data, byte start, byte datalen)
        {
            string str = "";
            byte[] byte1 = new byte[datalen];

            for (int i = 0; i < datalen; i++)
            {
                byte1[i] = data[start + i];
            }
            DLT64507.rever_char(ref byte1, byte1.Length);

            foreach (byte Member in byte1)
            {
                string str1 = Member.ToString("X2");
                str = str + str1;
            }
            return str;
        }

        /// <summary>
        /// serialPort1按键发送程序
        /// </summary>
        /// <param name="data"></param>
        /// <param name="datalen"></param>
        /// <returns></returns>
        private void Key_Sentences(byte[] pBuffer)
        {
            if (serialPort1.IsOpen)//判断串口是否打开，如果打开执行下一步操作输出12个字节 
            {
                serialPort1.Write(pBuffer, 0, 12);
            }
        }

        /// <summary>
        /// serialPort2按键发送程序
        /// </summary>
        /// <param name="data"></param>
        /// <param name="datalen"></param>
        /// <returns></returns>
        private void Key2_Sentences(byte[] pBuffer, int i)
        {
            if (serialPort2.IsOpen)//判断串口是否打开，如果打开执行下一步操作输出12个字节 
            {
                serialPort2.Write(pBuffer, 0, i);
            }
        }
        /// <summary>
        /// serialPort3按键发送程序
        /// </summary>
        /// <param name="data"></param>
        /// <param name="datalen"></param>
        /// <returns></returns>
        private void Key3_Sentences(byte[] pBuffer,int i)
        {
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作输出12个字节 
            {
                serialPort3.Write(pBuffer, 0, i);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public static void Delay(int milliSecond)
        {
            int start = Environment.TickCount;
            while (Math.Abs(Environment.TickCount - start) < milliSecond)//毫秒
            {
                Application.DoEvents();//可执行某无聊的操作
            }
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int byteNumber = serialPort1.BytesToRead;
                Delay(20);
                //延时等待数据接收完毕。
                while ((byteNumber < serialPort1.BytesToRead) && (serialPort1.BytesToRead < 4800))
                {
                    byteNumber = serialPort1.BytesToRead;
                    Delay(20);
                }
                int n = serialPort1.BytesToRead; //记录下缓冲区的字节个数 
                byte[] buf = new byte[n]; //声明一个临时数组存储当前来的串口数据 
                serialPort1.Read(buf, 0, n); //读取缓冲数据到buf中，同时将这串数据从缓冲区移除 
                //设置文字显示
                foreach (byte Member in buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox1.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox1.AppendText("\r\n");
                if(buf[0]==0x68&&buf[1]==0x00&&buf[2]==0x82&&buf[3]==0x06&&buf[11]==0x16)
                {
                    if (buf[4] == 0x00 && buf[5] == 0x00 && buf[6] == 0x00 && buf[7] == 0x00 && buf[8] == 0x00 && buf[9] == 0x00&&buf[10]==0xf0) textBox1.AppendText("A相电流切入计量回路" + "\r\n");
                    else if (buf[4] == 0x00 && buf[5] == 0x01 && buf[6] == 0x00 && buf[7] == 0x01 && buf[8] == 0x00 && buf[9] == 0x01 && buf[10] == 0xf3) textBox1.AppendText("A相电流切入标准阻抗箱" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x00 && buf[6] == 0x01 && buf[7] == 0x00 && buf[8] == 0x01 && buf[9] == 0x00 && buf[10] == 0xf3) textBox1.AppendText("A相一次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x01 && buf[6] == 0x01 && buf[7] == 0x01 && buf[8] == 0x01 && buf[9] == 0x01 && buf[10] == 0xf6) textBox1.AppendText("A相一次侧开路" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x02 && buf[6] == 0x01 && buf[7] == 0x02 && buf[8] == 0x01 && buf[9] == 0x02 && buf[10] == 0xf9) textBox1.AppendText("A相一次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x00 && buf[6] == 0x02 && buf[7] == 0x00 && buf[8] == 0x02 && buf[9] == 0x00 && buf[10] == 0xf6) textBox1.AppendText("A相二次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x01 && buf[6] == 0x02 && buf[7] == 0x01 && buf[8] == 0x02 && buf[9] == 0x01 && buf[10] == 0xf9) textBox1.AppendText("A相二次侧开路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x02 && buf[6] == 0x02 && buf[7] == 0x02 && buf[8] == 0x02 && buf[9] == 0x02 && buf[10] == 0xfc) textBox1.AppendText("A相二次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x03 && buf[6] == 0x02 && buf[7] == 0x03 && buf[8] == 0x02 && buf[9] == 0x03 && buf[10] == 0xff) textBox1.AppendText("A相二次侧二极管接入" + "\r\n");

                    else if (buf[4] == 0x03 && buf[5] == 0x00 && buf[6] == 0x03 && buf[7] == 0x00 && buf[8] == 0x03 && buf[9] == 0x00 && buf[10] == 0xf9) textBox1.AppendText("A相 250Ω" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x01 && buf[6] == 0x03 && buf[7] == 0x01 && buf[8] == 0x03 && buf[9] == 0x01 && buf[10] == 0xfc) textBox1.AppendText("A相 2KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x02 && buf[6] == 0x03 && buf[7] == 0x02 && buf[8] == 0x03 && buf[9] == 0x02 && buf[10] == 0xff) textBox1.AppendText("A相 3KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x03 && buf[6] == 0x03 && buf[7] == 0x03 && buf[8] == 0x03 && buf[9] == 0x03 && buf[10] == 0x02) textBox1.AppendText("A相 3.6KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x04 && buf[6] == 0x03 && buf[7] == 0x04 && buf[8] == 0x03 && buf[9] == 0x04 && buf[10] == 0x05) textBox1.AppendText("A相 3.9KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x05 && buf[6] == 0x03 && buf[7] == 0x05 && buf[8] == 0x03 && buf[9] == 0x05 && buf[10] == 0x08) textBox1.AppendText("A相 20KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x06 && buf[6] == 0x03 && buf[7] == 0x06 && buf[8] == 0x03 && buf[9] == 0x06 && buf[10] == 0x0b) textBox1.AppendText("A相 25KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x07 && buf[6] == 0x03 && buf[7] == 0x07 && buf[8] == 0x03 && buf[9] == 0x07 && buf[10] == 0x0e) textBox1.AppendText("A相 40KΩ" + "\r\n");
                    else textBox1.AppendText("无效应答" + "\r\n");
                }
                else if (buf[0] == 0x68 && buf[1] == 0x01 && buf[2] == 0x82 && buf[3] == 0x06 && buf[11] == 0x16)
                {
                    if (buf[4] == 0x00 && buf[5] == 0x00 && buf[6] == 0x00 && buf[7] == 0x00 && buf[8] == 0x00 && buf[9] == 0x00 && buf[10] == 0xf1) textBox1.AppendText("B相电流切入计量回路" + "\r\n");
                    else if (buf[4] == 0x00 && buf[5] == 0x01 && buf[6] == 0x00 && buf[7] == 0x01 && buf[8] == 0x00 && buf[9] == 0x01 && buf[10] == 0xf4) textBox1.AppendText("B相电流切入标准阻抗箱" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x00 && buf[6] == 0x01 && buf[7] == 0x00 && buf[8] == 0x01 && buf[9] == 0x00 && buf[10] == 0xf4) textBox1.AppendText("B相一次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x01 && buf[6] == 0x01 && buf[7] == 0x01 && buf[8] == 0x01 && buf[9] == 0x01 && buf[10] == 0xf7) textBox1.AppendText("B相一次侧开路" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x02 && buf[6] == 0x01 && buf[7] == 0x02 && buf[8] == 0x01 && buf[9] == 0x02 && buf[10] == 0xfa) textBox1.AppendText("B相一次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x00 && buf[6] == 0x02 && buf[7] == 0x00 && buf[8] == 0x02 && buf[9] == 0x00 && buf[10] == 0xf7) textBox1.AppendText("B相二次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x01 && buf[6] == 0x02 && buf[7] == 0x01 && buf[8] == 0x02 && buf[9] == 0x01 && buf[10] == 0xfa) textBox1.AppendText("B相二次侧开路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x02 && buf[6] == 0x02 && buf[7] == 0x02 && buf[8] == 0x02 && buf[9] == 0x02 && buf[10] == 0xfd) textBox1.AppendText("B相二次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x03 && buf[6] == 0x02 && buf[7] == 0x03 && buf[8] == 0x02 && buf[9] == 0x03 && buf[10] == 0x00) textBox1.AppendText("B相二次侧二极管接入" + "\r\n");

                    else if (buf[4] == 0x03 && buf[5] == 0x00 && buf[6] == 0x03 && buf[7] == 0x00 && buf[8] == 0x03 && buf[9] == 0x00 && buf[10] == 0xfa) textBox1.AppendText("B相 250Ω" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x01 && buf[6] == 0x03 && buf[7] == 0x01 && buf[8] == 0x03 && buf[9] == 0x01 && buf[10] == 0xfd) textBox1.AppendText("B相 2KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x02 && buf[6] == 0x03 && buf[7] == 0x02 && buf[8] == 0x03 && buf[9] == 0x02 && buf[10] == 0x00) textBox1.AppendText("B相 3KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x03 && buf[6] == 0x03 && buf[7] == 0x03 && buf[8] == 0x03 && buf[9] == 0x03 && buf[10] == 0x03) textBox1.AppendText("B相 3.6KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x04 && buf[6] == 0x03 && buf[7] == 0x04 && buf[8] == 0x03 && buf[9] == 0x04 && buf[10] == 0x06) textBox1.AppendText("B相 3.9KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x05 && buf[6] == 0x03 && buf[7] == 0x05 && buf[8] == 0x03 && buf[9] == 0x05 && buf[10] == 0x09) textBox1.AppendText("B相 20KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x06 && buf[6] == 0x03 && buf[7] == 0x06 && buf[8] == 0x03 && buf[9] == 0x06 && buf[10] == 0x0c) textBox1.AppendText("B相 25KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x07 && buf[6] == 0x03 && buf[7] == 0x07 && buf[8] == 0x03 && buf[9] == 0x07 && buf[10] == 0x0f) textBox1.AppendText("B相 40KΩ" + "\r\n");
                    else textBox1.AppendText("无效应答" + "\r\n");
                }
                else if (buf[0] == 0x68 && buf[1] == 0x02 && buf[2] == 0x82 && buf[3] == 0x06 && buf[11] == 0x16)
                {
                    if (buf[4] == 0x00 && buf[5] == 0x00 && buf[6] == 0x00 && buf[7] == 0x00 && buf[8] == 0x00 && buf[9] == 0x00 && buf[10] == 0xf2) textBox1.AppendText("C相电流切入计量回路" + "\r\n");
                    else if (buf[4] == 0x00 && buf[5] == 0x01 && buf[6] == 0x00 && buf[7] == 0x01 && buf[8] == 0x00 && buf[9] == 0x01 && buf[10] == 0xf5) textBox1.AppendText("C相电流切入标准阻抗箱" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x00 && buf[6] == 0x01 && buf[7] == 0x00 && buf[8] == 0x01 && buf[9] == 0x00 && buf[10] == 0xf5) textBox1.AppendText("C相一次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x01 && buf[6] == 0x01 && buf[7] == 0x01 && buf[8] == 0x01 && buf[9] == 0x01 && buf[10] == 0xf8) textBox1.AppendText("C相一次侧开路" + "\r\n");
                    else if (buf[4] == 0x01 && buf[5] == 0x02 && buf[6] == 0x01 && buf[7] == 0x02 && buf[8] == 0x01 && buf[9] == 0x02 && buf[10] == 0xfb) textBox1.AppendText("C相一次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x00 && buf[6] == 0x02 && buf[7] == 0x00 && buf[8] == 0x02 && buf[9] == 0x00 && buf[10] == 0xf8) textBox1.AppendText("C相二次侧正常接入" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x01 && buf[6] == 0x02 && buf[7] == 0x01 && buf[8] == 0x02 && buf[9] == 0x01 && buf[10] == 0xfb) textBox1.AppendText("C相二次侧开路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x02 && buf[6] == 0x02 && buf[7] == 0x02 && buf[8] == 0x02 && buf[9] == 0x02 && buf[10] == 0xfe) textBox1.AppendText("C相二次侧短路" + "\r\n");
                    else if (buf[4] == 0x02 && buf[5] == 0x03 && buf[6] == 0x02 && buf[7] == 0x03 && buf[8] == 0x02 && buf[9] == 0x03 && buf[10] == 0x01) textBox1.AppendText("C相二次侧二极管接入" + "\r\n");

                    else if (buf[4] == 0x03 && buf[5] == 0x00 && buf[6] == 0x03 && buf[7] == 0x00 && buf[8] == 0x03 && buf[9] == 0x00 && buf[10] == 0xfb) textBox1.AppendText("C相 250Ω" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x01 && buf[6] == 0x03 && buf[7] == 0x01 && buf[8] == 0x03 && buf[9] == 0x01 && buf[10] == 0xfe) textBox1.AppendText("C相 2KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x02 && buf[6] == 0x03 && buf[7] == 0x02 && buf[8] == 0x03 && buf[9] == 0x02 && buf[10] == 0x01) textBox1.AppendText("C相 3KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x03 && buf[6] == 0x03 && buf[7] == 0x03 && buf[8] == 0x03 && buf[9] == 0x03 && buf[10] == 0x04) textBox1.AppendText("C相 3.6KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x04 && buf[6] == 0x03 && buf[7] == 0x04 && buf[8] == 0x03 && buf[9] == 0x04 && buf[10] == 0x07) textBox1.AppendText("C相 3.9KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x05 && buf[6] == 0x03 && buf[7] == 0x05 && buf[8] == 0x03 && buf[9] == 0x05 && buf[10] == 0x0a) textBox1.AppendText("C相 20KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x06 && buf[6] == 0x03 && buf[7] == 0x06 && buf[8] == 0x03 && buf[9] == 0x06 && buf[10] == 0x0d) textBox1.AppendText("C相 25KΩ" + "\r\n");
                    else if (buf[4] == 0x03 && buf[5] == 0x07 && buf[6] == 0x03 && buf[7] == 0x07 && buf[8] == 0x03 && buf[9] == 0x07 && buf[10] == 0x10) textBox1.AppendText("C相 40KΩ" + "\r\n");
                    else textBox1.AppendText("无效应答" + "\r\n");
                }
                else if (buf[0] == 0x68 && buf[2] == 0xc2 && buf[3] == 0x01 && buf[4] == 0x01 && buf[6] == 0x16) textBox1.AppendText("装置否认" + "\r\n");
                else textBox1.AppendText("无效应答" + "\r\n");
            }
            catch 
            {
            
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();    //关闭串口
                    button16.Text = "打开串口";
                    button16.BackColor = Color.ForestGreen;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    comboBox4.Enabled = true;
                    comboBox5.Enabled = true;
                }
                else
                {
                    serialPort1.PortName = comboBox1.Text;//设置串口号
                    serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text, 10);//十进制数据转换，设置波特率
                    serialPort1.Open();     //打开串口
                    comboBox1.Enabled = false;
                    comboBox2.Enabled = false;
                    comboBox3.Enabled = false;
                    comboBox4.Enabled = false;
                    comboBox5.Enabled = false;
                    button16.Text = "关闭串口";
                    button16.BackColor = Color.Firebrick;
                }
            }
            catch
            {
                MessageBox.Show("端口错误,请检查串口", "错误");
            }
            if (comboBox3.Text.Trim() == "0")    //设置停止位
            {
                serialPort1.StopBits = StopBits.None;
            }
            else if (comboBox3.Text.Trim() == "1.5")
            {
                serialPort1.StopBits = StopBits.OnePointFive;
            }
            else if (comboBox3.Text.Trim() == "2")
            {
                serialPort1.StopBits = StopBits.Two;
            }
            else
            {
                serialPort1.StopBits = StopBits.One;
            }
            serialPort1.DataBits = Convert.ToInt16(comboBox5.Text.Trim());    //设置数据位

            if (comboBox4.Text.Trim() == "奇校验")    //设置校验
            {
                serialPort1.Parity = Parity.Odd;
            }
            else if (comboBox4.Text.Trim() == "偶校验")
            {
                serialPort1.Parity = Parity.Even;
            }
            else
            {
                serialPort1.Parity = Parity.None;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
           
        }

        private void button18_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";//清屏
        }

        #region 阻抗箱设置
        private void Control_Color1(int num)
        {
            button1.BackColor = Color.White;
            button2.BackColor = Color.White;
            button3.BackColor = Color.White;
            button4.BackColor = Color.White;
            button5.BackColor = Color.White;
            button8.BackColor = Color.White;
            button6.BackColor = Color.White;
            button7.BackColor = Color.White;
            button9.BackColor = Color.White;
            button59.BackColor = Color.White;
            button61.BackColor = Color.White;
            button10.BackColor = Color.White;
            button19.BackColor = Color.White;
            button11.BackColor = Color.White;
            button12.BackColor = Color.White;
            button20.BackColor = Color.White;
            button13.BackColor = Color.White;
            button14.BackColor = Color.White;
            button15.BackColor = Color.White;
            button63.BackColor = Color.White;
            button64.BackColor = Color.White;
            button65.BackColor = Color.White;
            switch (num)
            {
                case 1: button1.BackColor = Color.ForestGreen; break;
                case 2: button2.BackColor = Color.ForestGreen; break;
                case 3: button3.BackColor = Color.ForestGreen; break;
                case 4: button4.BackColor = Color.ForestGreen; break;
                case 5: button5.BackColor = Color.ForestGreen; break;
                case 6: button8.BackColor = Color.ForestGreen; break;
                case 7: button6.BackColor = Color.ForestGreen; break;
                case 8: button7.BackColor = Color.ForestGreen; break;
                case 9: button9.BackColor = Color.ForestGreen; break;
                case 10: button59.BackColor = Color.ForestGreen; break;
                case 11: button61.BackColor = Color.ForestGreen; break;
                case 12: button10.BackColor = Color.ForestGreen; break;
                case 13: button19.BackColor = Color.ForestGreen; break;
                case 14: button11.BackColor = Color.ForestGreen; break;
                case 15: button12.BackColor = Color.ForestGreen; break;
                case 16: button20.BackColor = Color.ForestGreen; break;
                case 17: button13.BackColor = Color.ForestGreen; break;
                case 18: button14.BackColor = Color.ForestGreen; break;
                case 19: button15.BackColor = Color.ForestGreen; break;
                case 20: button63.BackColor = Color.ForestGreen; break;
                case 21: button64.BackColor = Color.ForestGreen; break;
                case 22: button65.BackColor = Color.ForestGreen; break;
            }
        }

        private void Control_Color2(int num)
        {
            button24.BackColor = Color.White;
            button25.BackColor = Color.White;
            button26.BackColor = Color.White;
            button27.BackColor = Color.White;
            button28.BackColor = Color.White;
            button31.BackColor = Color.White;
            button29.BackColor = Color.White;
            button30.BackColor = Color.White;
            button32.BackColor = Color.White;
            button30.BackColor = Color.White;
            button66.BackColor = Color.White;
            button83.BackColor = Color.White;
            button33.BackColor = Color.White;
            button22.BackColor = Color.White;
            button34.BackColor = Color.White;
            button35.BackColor = Color.White;
            button21.BackColor = Color.White;
            button36.BackColor = Color.White;
            button37.BackColor = Color.White;
            button23.BackColor = Color.White;
            button84.BackColor = Color.White;
            button88.BackColor = Color.White;
            button89.BackColor = Color.White;
            switch (num)
            {
                case 1:  button24.BackColor = Color.ForestGreen; break;
                case 2:  button25.BackColor = Color.ForestGreen; break;
                case 3:  button26.BackColor = Color.ForestGreen; break;
                case 4:  button27.BackColor = Color.ForestGreen; break;
                case 5:  button28.BackColor = Color.ForestGreen; break;
                case 6:  button31.BackColor = Color.ForestGreen; break;
                case 7:  button29.BackColor = Color.ForestGreen; break;
                case 8:  button30.BackColor = Color.ForestGreen; break;
                case 9:  button32.BackColor = Color.ForestGreen; break;
                case 10: button66.BackColor = Color.ForestGreen; break;
                case 11: button83.BackColor = Color.ForestGreen; break;
                case 12: button33.BackColor = Color.ForestGreen; break;
                case 13: button22.BackColor = Color.ForestGreen; break;
                case 14: button34.BackColor = Color.ForestGreen; break;
                case 15: button35.BackColor = Color.ForestGreen; break;
                case 16: button21.BackColor = Color.ForestGreen; break;
                case 17: button36.BackColor = Color.ForestGreen; break;
                case 18: button37.BackColor = Color.ForestGreen; break;
                case 19: button23.BackColor = Color.ForestGreen; break;
                case 20: button84.BackColor = Color.ForestGreen; break;
                case 21: button88.BackColor = Color.ForestGreen; break;
                case 22: button89.BackColor = Color.ForestGreen; break;
            }
        }

        private void Control_Color3(int num)
        {
            button41.BackColor = Color.White;
            button42.BackColor = Color.White;
            button43.BackColor = Color.White;
            button44.BackColor = Color.White;
            button45.BackColor = Color.White;
            button48.BackColor = Color.White;
            button46.BackColor = Color.White;
            button47.BackColor = Color.White;
            button49.BackColor = Color.White;
            button90.BackColor = Color.White;
            button91.BackColor = Color.White;
            button50.BackColor = Color.White;
            button39.BackColor = Color.White;
            button51.BackColor = Color.White;
            button52.BackColor = Color.White;
            button38.BackColor = Color.White;
            button53.BackColor = Color.White;
            button54.BackColor = Color.White;
            button40.BackColor = Color.White;
            button92.BackColor = Color.White;
            button93.BackColor = Color.White;
            button94.BackColor = Color.White;
            switch (num)
            {
                case 1:  button41.BackColor = Color.ForestGreen; break;
                case 2:  button42.BackColor = Color.ForestGreen; break;
                case 3:  button43.BackColor = Color.ForestGreen; break;
                case 4:  button44.BackColor = Color.ForestGreen; break;
                case 5:  button45.BackColor = Color.ForestGreen; break;
                case 6:  button48.BackColor = Color.ForestGreen; break;
                case 7:  button46.BackColor = Color.ForestGreen; break;
                case 8:  button47.BackColor = Color.ForestGreen; break;
                case 9:  button49.BackColor = Color.ForestGreen; break;
                case 10: button90.BackColor = Color.ForestGreen; break;
                case 11: button91.BackColor = Color.ForestGreen; break;
                case 12: button50.BackColor = Color.ForestGreen; break;
                case 13: button39.BackColor = Color.ForestGreen; break;
                case 14: button51.BackColor = Color.ForestGreen; break;
                case 15: button52.BackColor = Color.ForestGreen; break;
                case 16: button38.BackColor = Color.ForestGreen; break;
                case 17: button53.BackColor = Color.ForestGreen; break;
                case 18: button54.BackColor = Color.ForestGreen; break;
                case 19: button40.BackColor = Color.ForestGreen; break;
                case 20: button92.BackColor = Color.ForestGreen; break;
                case 21: button93.BackColor = Color.ForestGreen; break;
                case 22: button94.BackColor = Color.ForestGreen; break;
            }
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x70, 0x16 };
            Key_Sentences(Data);
            Control_Color1(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x73, 0x16 };
            Key_Sentences(Data);
            Control_Color1(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00, 0x73, 0x16 };
            Key_Sentences(Data);
            Control_Color1(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x76, 0x16 };
            Key_Sentences(Data);
            Control_Color1(4);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x79, 0x16 };
            Key_Sentences(Data);
            Control_Color1(5);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x02, 0x00, 0x02, 0x00, 0x02, 0x00, 0x76, 0x16 };
            Key_Sentences(Data);
            Control_Color1(6);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01, 0x79, 0x16 };
            Key_Sentences(Data);
            Control_Color1(7);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x02, 0x02, 0x02, 0x02, 0x02, 0x02, 0x7C, 0x16 };
            Key_Sentences(Data);
            Control_Color1(8);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x02, 0x03, 0x02, 0x03, 0x02, 0x03, 0x7f, 0x16 };
            Key_Sentences(Data);
            Control_Color1(9);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x00, 0x03, 0x00, 0x03, 0x00, 0x79, 0x16 };
            Key_Sentences(Data);
            Control_Color1(12);
        }

        private void button19_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x01, 0x03, 0x01, 0x03, 0x01, 0x7C, 0x16 };
            Key_Sentences(Data);
            Control_Color1(13);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x02, 0x03, 0x02, 0x03, 0x02, 0x7f, 0x16 };
            Key_Sentences(Data);
            Control_Color1(14);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x03, 0x03, 0x03, 0x03, 0x03, 0x82, 0x16 };
            Key_Sentences(Data);
            Control_Color1(15);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x04, 0x03, 0x04, 0x03, 0x04, 0x85, 0x16 };
            Key_Sentences(Data);
            Control_Color1(16);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x05, 0x03, 0x05, 0x03, 0x05, 0x88, 0x16 };
            Key_Sentences(Data);
            Control_Color1(17);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x06, 0x03, 0x06, 0x03, 0x06, 0x8b, 0x16 };
            Key_Sentences(Data);
            Control_Color1(18);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x07, 0x03, 0x07, 0x03, 0x07, 0x8e, 0x16 };
            Key_Sentences(Data);
            Control_Color1(19);
        }

        private void button24_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x71, 0x16 };
            Key_Sentences(Data);
            Control_Color2(1);
        }

        private void button25_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x74, 0x16 };
            Key_Sentences(Data);
            Control_Color2(2);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00, 0x74, 0x16 };
            Key_Sentences(Data);
            Control_Color2(3);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x77, 0x16 };
            Key_Sentences(Data);
            Control_Color2(4);
        }

        private void button28_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x7A, 0x16 };
            Key_Sentences(Data);
            Control_Color2(5);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x02, 0x00, 0x02, 0x00, 0x02, 0x00, 0x77, 0x16 };
            Key_Sentences(Data);
            Control_Color2(6);
        }

        private void button29_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01, 0x7A, 0x16 };
            Key_Sentences(Data);
            Control_Color2(7);
        }

        private void button30_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x02, 0x02, 0x02, 0x02, 0x02, 0x02, 0x7D, 0x16 };
            Key_Sentences(Data);
            Control_Color2(8);
        }

        private void button32_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x02, 0x03, 0x02, 0x03, 0x02, 0x03, 0x80, 0x16 };
            Key_Sentences(Data);
            Control_Color2(9);
        }

        private void button33_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x00, 0x03, 0x00, 0x03, 0x00, 0x7A, 0x16 };
            Key_Sentences(Data);
            Control_Color2(12);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x01, 0x03, 0x01, 0x03, 0x01, 0x7D, 0x16 };
            Key_Sentences(Data);
            Control_Color2(13);
        }

        private void button34_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x02, 0x03, 0x02, 0x03, 0x02, 0x80, 0x16 };
            Key_Sentences(Data);
            Control_Color2(14);
        }

        private void button35_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x03, 0x03, 0x03, 0x03, 0x03, 0x83, 0x16 };
            Key_Sentences(Data);
            Control_Color2(15);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x04, 0x03, 0x04, 0x03, 0x04, 0x86, 0x16 };
            Key_Sentences(Data);
            Control_Color2(16);
        }

        private void button36_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x05, 0x03, 0x05, 0x03, 0x05, 0x89, 0x16 };
            Key_Sentences(Data);
            Control_Color2(17);
        }

        private void button37_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x06, 0x03, 0x06, 0x03, 0x06, 0x8C, 0x16 };
            Key_Sentences(Data);
            Control_Color2(18);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x07, 0x03, 0x07, 0x03, 0x07, 0x8F, 0x16 };
            Key_Sentences(Data);
            Control_Color2(19);
        }

        private void button41_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x72, 0x16 };
            Key_Sentences(Data);
            Control_Color3(1);
        }

        private void button42_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x75, 0x16 };
            Key_Sentences(Data);
            Control_Color3(2);
        }

        private void button43_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00, 0x75, 0x16 };
            Key_Sentences(Data);
            Control_Color3(3);
        }

        private void button44_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x78, 0x16 };
            Key_Sentences(Data);
            Control_Color3(4);
        }

        private void button45_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x01, 0x02, 0x01, 0x02, 0x01, 0x02, 0x7B, 0x16 };
            Key_Sentences(Data);
            Control_Color3(5);
        }

        private void button48_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x02, 0x00, 0x02, 0x00, 0x02, 0x00, 0x78, 0x16 };
            Key_Sentences(Data);
            Control_Color3(6);
        }

        private void button46_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x02, 0x01, 0x02, 0x01, 0x02, 0x01, 0x7B, 0x16 };
            Key_Sentences(Data);
            Control_Color3(7);
        }

        private void button47_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x02, 0x02, 0x02, 0x02, 0x02, 0x02, 0x7E, 0x16 };
            Key_Sentences(Data);
            Control_Color3(8);
        }

        private void button49_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x02, 0x03, 0x02, 0x03, 0x02, 0x03, 0x81, 0x16 };
            Key_Sentences(Data);
            Control_Color3(9);
        }

        private void button50_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x00, 0x03, 0x00, 0x03, 0x00, 0x7B, 0x16 };
            Key_Sentences(Data);
            Control_Color3(12);
        }

        private void button39_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x01, 0x03, 0x01, 0x03, 0x01, 0x7E, 0x16 };
            Key_Sentences(Data);
            Control_Color3(13);
        }

        private void button51_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x02, 0x03, 0x02, 0x03, 0x02, 0x81, 0x16 };
            Key_Sentences(Data);
            Control_Color3(14);
        }

        private void button52_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x03, 0x03, 0x03, 0x03, 0x03, 0x84, 0x16 };
            Key_Sentences(Data);
            Control_Color3(15);
        }

        private void button38_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x04, 0x03, 0x04, 0x03, 0x04, 0x87, 0x16 };
            Key_Sentences(Data);
            Control_Color3(16);
        }

        private void button53_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x05, 0x03, 0x05, 0x03, 0x05, 0x8A, 0x16 };
            Key_Sentences(Data);
            Control_Color3(17);
        }

        private void button54_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x06, 0x03, 0x06, 0x03, 0x06, 0x8D, 0x16 };
            Key_Sentences(Data);
            Control_Color3(18);
        }

        private void button40_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x07, 0x03, 0x07, 0x03, 0x07, 0x90, 0x16 };
            Key_Sentences(Data);
            Control_Color3(19);
        }
        #endregion


        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox23_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void serialPort2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int byteNumber = serialPort2.BytesToRead;
                Delay(20);
                //延时等待数据接收完毕。
                while ((byteNumber < serialPort2.BytesToRead) && (serialPort2.BytesToRead < 4800))
                {
                    byteNumber = serialPort2.BytesToRead;
                    Delay(20);
                }
                int n = serialPort2.BytesToRead; //记录下缓冲区的字节个数 
                byte[] buf = new byte[n]; //声明一个临时数组存储当前来的串口数据 
                serialPort2.Read(buf, 0, n); //读取缓冲数据到buf中，同时将这串数据从缓冲区移除 
                
                if(buf[0]==0x81&&buf[1]==0x00&&buf[2]==0x80&&buf[3]==0x00&&buf[4]==0x4d&n==128)
                {
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                //textBox13.Text = "";
                //textBox14.Text = "";
                //textBox15.Text = "";
                //textBox16.Text = "";
                //textBox17.Text = "";
                //textBox18.Text = "";
                textBox19.Text = "";
                textBox20.Text = "";
                textBox21.Text = "";
                textBox22.Text = "";
                textBox23.Text = "";
                textBox24.Text = "";
                textBox25.Text = "";
                textBox26.Text = "";
                textBox27.Text = "";
                textBox28.Text = "";
                textBox29.Text = "";
                textBox30.Text = "";
                textBox31.Text = "";
                textBox32.Text = "";
                textBox33.Text = "";
                textBox34.Text = "";
                textBox35.Text = "";
                byte[] mark1 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark1[h] = buf[5+h];
                }
                int rate = BitConverter.ToInt32(mark1.ToArray(), 0);
                 rate = rate / 10000;
                string str1 = Convert.ToString(rate);
                textBox35.AppendText((str1.Length == 1 ? "0" + str1 : str1));//频率
                //if (buf[9]  == 0) { textBox5.AppendText("UA量限：" + "380V" + "\r\n"); } else if (buf[9] == 1) { textBox5.AppendText("UA量限：" + "220V" + "\r\n"); } else if (buf[9] == 2) { textBox5.AppendText("UA量限：" + "100V" + "\r\n"); } else if (buf[9] == 3) { textBox5.AppendText("UA量限：" + "57V" + "\r\n"); } else if (buf[9] == 4) { textBox5.AppendText("UA量限：" + "30V" + "\r\n"); } else if (buf[9] == 5) { textBox5.AppendText("UA量限：" + "600V" + "\r\n"); }
                //if (buf[10] == 0) { textBox5.AppendText("UB量限：" + "380V" + "\r\n"); } else if (buf[10] == 1) { textBox5.AppendText("UB量限：" + "220V" + "\r\n"); } else if (buf[10] == 2) { textBox5.AppendText("UB量限：" + "100V" + "\r\n"); } else if (buf[10] == 3) { textBox5.AppendText("UB量限：" + "57V" + "\r\n"); } else if (buf[10] == 4) { textBox5.AppendText("UB量限：" + "30V" + "\r\n"); } else if (buf[10] == 5) { textBox5.AppendText("UB量限：" + "600V" + "\r\n"); }
                //if (buf[11] == 0) { textBox5.AppendText("UC量限：" + "380V" + "\r\n"); } else if (buf[11] == 1) { textBox5.AppendText("UC量限：" + "220V" + "\r\n"); } else if (buf[11] == 2) { textBox5.AppendText("UC量限：" + "100V" + "\r\n"); } else if (buf[11] == 3) { textBox5.AppendText("UC量限：" + "57V" + "\r\n"); } else if (buf[11] == 4) { textBox5.AppendText("UC量限：" + "30V" + "\r\n"); } else if (buf[11] == 5) { textBox5.AppendText("UC量限：" + "600V" + "\r\n"); }
                //if (buf[12] == 0) { textBox5.AppendText("IA量限：" + "20A" + "\r\n"); } else if (buf[12] == 1) { textBox5.AppendText("IA量限：" + "5A" + "\r\n"); } else if (buf[12] == 2) { textBox5.AppendText("IA量限：" + "1A" + "\r\n"); } else if (buf[12] == 3) { textBox5.AppendText("IA量限：" + "0.2A" + "\r\n"); } else if (buf[12] == 4) { textBox5.AppendText("IA量限：" + "10A" + "\r\n"); } else if (buf[12] == 5) { textBox5.AppendText("IA量限：" + "60A" + "\r\n"); }
                //if (buf[13] == 0) { textBox5.AppendText("IB量限：" + "20A" + "\r\n"); } else if (buf[13] == 1) { textBox5.AppendText("IB量限：" + "5A" + "\r\n"); } else if (buf[13] == 2) { textBox5.AppendText("IB量限：" + "1A" + "\r\n"); } else if (buf[13] == 3) { textBox5.AppendText("IB量限：" + "0.2A" + "\r\n"); } else if (buf[13] == 4) { textBox5.AppendText("IB量限：" + "10A" + "\r\n"); } else if (buf[13] == 5) { textBox5.AppendText("IB量限：" + "00A" + "\r\n"); }
                //if (buf[14] == 0) { textBox5.AppendText("IC量限：" + "20A" + "\r\n"); } else if (buf[14] == 1) { textBox5.AppendText("IC量限：" + "5A" + "\r\n"); } else if (buf[14] == 2) { textBox5.AppendText("IC量限：" + "1A" + "\r\n"); } else if (buf[14] == 3) { textBox5.AppendText("IC量限：" + "0.2A" + "\r\n"); } else if (buf[14] == 4) { textBox5.AppendText("IC量限：" + "10A" + "\r\n"); } else if (buf[14] == 5) { textBox5.AppendText("IC量限：" + "60A" + "\r\n"); }
                byte[] mark2 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark2[h] = buf[15+h];
                }
                double tensionA = BitConverter.ToInt32(mark2.ToArray(), 0);
                if (buf[9] == 0 || buf[9] == 1 || buf[9] == 2 || buf[9] == 5) tensionA = tensionA / 1000;
                else if (buf[9] == 3 || buf[9] == 4) tensionA = tensionA / 10000;
                string str2 = Convert.ToString(tensionA);
                textBox7.AppendText((str2.Length == 1 ? "0" + str2 : str2) );//UA
                byte[] mark3 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark3[h] = buf[19 + h];
                }
                double tensionB = BitConverter.ToInt32(mark3.ToArray(), 0);
                if (buf[10] == 0 || buf[10] == 1 || buf[10] == 2 || buf[10] == 5) tensionB = tensionB / 1000;
                else if (buf[10] == 3 || buf[10] == 4) tensionB = tensionB / 10000;
                string str3 = Convert.ToString(tensionB);
                textBox8.AppendText((str3.Length == 1 ? "0" + str3 : str3) );//UB
                byte[] mark4 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark4[h] = buf[23 + h];
                }
                double tensionC = BitConverter.ToInt32(mark4.ToArray(), 0);
                if (buf[11] == 0 || buf[11] == 1 || buf[11] == 2 || buf[11] == 5) tensionC = tensionC / 1000;
                else if (buf[11] == 3 || buf[11] == 4) tensionC = tensionC / 10000;
                string str4 = Convert.ToString(tensionC);
                textBox9.AppendText((str4.Length == 1 ? "0" + str4 : str4) );//UC
                byte[] mark5 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark5[h] = buf[27 + h];
                }
                double galvanicA = BitConverter.ToInt32(mark5.ToArray(), 0);
                if (buf[12] == 0 || buf[12] == 4 || buf[12] == 5) galvanicA = galvanicA / 10000;
                else if (buf[12] == 1 || buf[12] == 2) galvanicA = galvanicA / 100000;
                else if (buf[12] == 3) galvanicA = galvanicA / 1000000;
                string str5 = Convert.ToString(galvanicA);
                textBox10.AppendText((str5.Length == 1 ? "0" + str5 : str5) );//IA
                byte[] mark6 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark6[h] = buf[31 + h];
                }
                double galvanicB = BitConverter.ToInt32(mark6.ToArray(), 0);
                if (buf[12] == 0 || buf[12] == 4 || buf[12] == 5) galvanicB = galvanicB / 10000;
                else if (buf[12] == 1 || buf[12] == 2) galvanicB = galvanicB / 100000;
                else if (buf[12] == 3) galvanicB = galvanicB/ 1000000;
                string str6 = Convert.ToString(galvanicB);
                textBox11.AppendText((str6.Length == 1 ? "0" + str6 : str6) );//IB
                byte[] mark7 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark7[h] = buf[35 + h];
                }
                double galvanicC = BitConverter.ToInt32(mark7.ToArray(), 0);
                if (buf[12] == 0 || buf[12] == 4 || buf[12] == 5) galvanicC = galvanicC / 10000;
                else if (buf[12] == 1 || buf[12] == 2) galvanicC= galvanicC/ 100000;
                else if (buf[12] == 3) galvanicC = galvanicC / 1000000;
                string str7 = Convert.ToString(galvanicC);
                textBox12.AppendText((str7.Length == 1 ? "0" + str7 : str7));//IC
               // byte[] mark8 = new byte[4];
               // for (int h = 0; h < 4; h++)
               // {
               //     mark8[h] = buf[39 + h];
               // }
               // double phaseUA = BitConverter.ToInt32(mark8.ToArray(), 0);
               // phaseUA = phaseUA / 1000;
               // string str8 = Convert.ToString(phaseUA);
               // if (phaseUA>0)
               //{
               // textBox13.AppendText((str8.Length == 1 ? "0" + str8 : str8));//UA角度
               //}
               // else textBox13.AppendText("0");
               // byte[] mark9 = new byte[4];
               // for (int h = 0; h < 4; h++)
               // {
               //     mark9[h] = buf[43 + h];
               // }
               // double phaseUB = BitConverter.ToInt32(mark9.ToArray(), 0);
               // phaseUB = phaseUB / 1000;
               // string str9 = Convert.ToString(phaseUB);
               // if (phaseUB > 0)
               // {
               //     textBox14.AppendText((str9.Length == 1 ? "0" + str9 : str9));//UB角度
               // }
               // else textBox14.AppendText("0");
               // byte[] mark10 = new byte[4];
               // for (int h = 0; h < 4; h++)
               // {
               //     mark10[h] = buf[47 + h];
               // }
               // double phaseUC = BitConverter.ToInt32(mark10.ToArray(), 0);
               // phaseUC = phaseUC / 1000;
               // string str10 = Convert.ToString(phaseUC);
               // if (phaseUC > 0)
               // {
               //     textBox15.AppendText((str10.Length == 1 ? "0" + str10 : str10));//UC角度
               // }
               // else textBox15.AppendText("0");
                //byte[] mark11 = new byte[4];
                //for (int h = 0; h < 4; h++)
                //{
                //    mark11[h] = buf[51 + h];
                //}
                //double phaseIA = BitConverter.ToInt32(mark11.ToArray(), 0);
                //phaseIA = phaseIA / 1000;
                //string str11 = Convert.ToString(phaseIA);
                //if (phaseIA > 0)
                //{
                //    textBox16.AppendText((str11.Length == 1 ? "0" + str11 : str11));//IA角度
                //}
                //else textBox16.AppendText("0");
                //byte[] mark12 = new byte[4];
                //for (int h = 0; h < 4; h++)
                //{
                //    mark12[h] = buf[55 + h];
                //}
                //double phaseIB = BitConverter.ToInt32(mark12.ToArray(), 0);
                //phaseIB = phaseIB / 1000;
                //string str12 = Convert.ToString(phaseIB);
                //if (phaseIB > 0)
                //{
                //    textBox17.AppendText((str12.Length == 1 ? "0" + str12 : str12));//IB角度
                //}
                //else textBox17.AppendText("0");
                //byte[] mark13 = new byte[4];
                //for (int h = 0; h < 4; h++)
                //{
                //    mark13[h] = buf[59 + h];
                //}
                //double phaseIC = BitConverter.ToInt32(mark13.ToArray(), 0);
                //phaseIC = phaseIC / 1000;
                //string str13 = Convert.ToString(phaseIC);
                //if (phaseIC > 0)
                //{
                //    textBox18.AppendText((str13.Length == 1 ? "0" + str13 : str13));//IC角度
                //}
                //else textBox18.AppendText("0");
                byte[] mark14 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark14[h] = buf[63 + h];
                }
                double activeA = BitConverter.ToInt32(mark14.ToArray(), 0);
                if (buf[9] == 4 && buf[12] == 3) activeA = activeA / 100000;
                else if (buf[12] == 3) activeA = activeA / 10000;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 2) activeA = activeA / 10000;
                else if (buf[12] == 2) activeA = activeA / 1000;
                else if ((buf[9] == 2 || buf[9] == 3 || buf[9] == 4) && buf[12] == 1) activeA = activeA / 1000;
                else if (buf[12] == 1) activeA = activeA / 100;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 4) activeA = activeA / 1000;
                else if (buf[12] == 4) activeA = activeA / 100;
                else if (buf[9] == 4 && buf[12] == 0) activeA = activeA / 1000;
                else if (buf[12] == 0) activeA = activeA / 100;
                else if (buf[12] == 5) activeA = activeA / 100;
                string str14 = Convert.ToString(activeA);
                textBox19.AppendText((str14.Length == 1 ? "0" + str14 : str14));//A相有功功率测量值
                byte[] mark15 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark15[h] = buf[67 + h];
                }
                double activeB = BitConverter.ToInt32(mark15.ToArray(), 0);
                if (buf[10] == 4 && buf[13] == 3) activeB = activeB / 100000;
                else if (buf[13] == 3) activeB = activeB / 10000;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 2) activeB = activeB / 10000;
                else if (buf[13] == 2) activeB = activeB / 1000;
                else if ((buf[10] == 2 || buf[10] == 3 || buf[10] == 4) && buf[13] == 1) activeB = activeB / 1000;
                else if (buf[13] == 1) activeB = activeB / 100;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 4) activeB = activeB / 1000;
                else if (buf[13] == 4) activeB = activeB / 100;
                else if (buf[10] == 4 && buf[13] == 0) activeB = activeB / 1000;
                else if (buf[13] == 0) activeB = activeB / 100;
                else if (buf[13] == 5) activeB = activeB / 100;
                string str15 = Convert.ToString(activeB);
                textBox20.AppendText((str15.Length == 1 ? "0" + str15 : str15));//B相有功功率测量值
                byte[] mark16 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark16[h] = buf[71 + h];
                }
                double activeC = BitConverter.ToInt32(mark16.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) activeC = activeC / 100000;
                else if (buf[14] == 3) activeC = activeC / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) activeC = activeC / 10000;
                else if (buf[14] == 2) activeC = activeC / 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) activeC = activeC / 1000;
                else if (buf[14] == 1) activeC = activeC / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) activeC = activeC / 1000;
                else if (buf[14] == 4) activeC = activeC / 100;
                else if (buf[11] == 4 && buf[14] == 0) activeC = activeC / 1000;
                else if (buf[14] == 0) activeC = activeC / 100;
                else if (buf[14] == 5) activeC = activeC / 100;
                string str16 = Convert.ToString(activeC);
                textBox21.AppendText((str16.Length == 1 ? "0" + str16 : str16));//C相有功功率测量值
                byte[] mark26 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark26[h] = buf[75 + h];
                }
                double active = BitConverter.ToInt32(mark26.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) active = active / 100000;
                else if (buf[14] == 3) active = active / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) active = active / 10000;
                else if (buf[14] == 2) active = active/ 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) active = active/ 1000;
                else if (buf[14] == 1) active = active / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) active = active/ 1000;
                else if (buf[14] == 4) active = active / 100;
                else if (buf[11] == 4 && buf[14] == 0) active = active / 1000;
                else if (buf[14] == 0) active= active / 100;
                else if (buf[14] == 5) active = active / 100;
                string str26 = Convert.ToString(active);
                textBox31.AppendText((str26.Length == 1 ? "0" + str26 : str26));//总有功功率测量值
                byte[] mark17 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark17[h] = buf[79 + h];
                }
                double reactiveA = BitConverter.ToInt32(mark17.ToArray(), 0);
                if (buf[9] == 4 && buf[12] == 3) reactiveA = reactiveA / 100000;
                else if (buf[12] == 3) reactiveA = reactiveA / 10000;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 2) reactiveA = reactiveA / 10000;
                else if (buf[12] == 2) reactiveA = reactiveA / 1000;
                else if ((buf[9] == 2 || buf[9] == 3 || buf[9] == 4) && buf[12] == 1) reactiveA = reactiveA / 1000;
                else if (buf[12] == 1) reactiveA = reactiveA / 100;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 4) reactiveA = reactiveA / 1000;
                else if (buf[12] == 4) reactiveA = reactiveA / 100;
                else if (buf[9] == 4 && buf[12] == 0) reactiveA = reactiveA / 1000;
                else if (buf[12] == 0) reactiveA = reactiveA / 100;
                else if (buf[12] == 5) reactiveA = reactiveA / 100;
                string str17 = Convert.ToString(reactiveA);
                textBox22.AppendText((str17.Length == 1 ? "0" + str17 : str17));//A相无功功率测量值
                byte[] mark18 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark18[h] = buf[83 + h];
                }
                double reactiveB = BitConverter.ToInt32(mark18.ToArray(), 0);
                if (buf[10] == 4 && buf[13] == 3) reactiveB = reactiveB / 100000;
                else if (buf[13] == 3) reactiveB = reactiveB / 10000;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 2) reactiveB = reactiveB / 10000;
                else if (buf[13] == 2) reactiveB = reactiveB / 1000;
                else if ((buf[10] == 2 || buf[10] == 3 || buf[10] == 4) && buf[13] == 1) reactiveB = reactiveB / 1000;
                else if (buf[13] == 1) reactiveB = reactiveB / 100;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 4) reactiveB = reactiveB / 1000;
                else if (buf[13] == 4) reactiveB = reactiveB / 100;
                else if (buf[10] == 4 && buf[13] == 0) reactiveB = reactiveB / 1000;
                else if (buf[13] == 0) reactiveB = reactiveB / 100;
                else if (buf[13] == 5) reactiveB = reactiveB / 100;
                string str18 = Convert.ToString(reactiveB);
                textBox23.AppendText( (str18.Length == 1 ? "0" + str18 : str18));//B相无功功率测量值
                byte[] mark19 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark19[h] = buf[87 + h];
                }
                double reactiveC = BitConverter.ToInt32(mark19.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) reactiveC = reactiveC / 100000;
                else if (buf[14] == 3) reactiveC = reactiveC / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) reactiveC = reactiveC / 10000;
                else if (buf[14] == 2) reactiveC = reactiveC / 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) reactiveC = reactiveC / 1000;
                else if (buf[14] == 1) reactiveC = reactiveC / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) reactiveC = reactiveC / 1000;
                else if (buf[14] == 4) reactiveC = reactiveC / 100;
                else if (buf[11] == 4 && buf[14] == 0) reactiveC = reactiveC / 1000;
                else if (buf[14] == 0) reactiveC = reactiveC / 100;
                else if (buf[14] == 5) reactiveC = reactiveC / 100;
                string str19 = Convert.ToString(reactiveC);
                textBox24.AppendText((str19.Length == 1 ? "0" + str19 : str19));//C相无功功率测量值

                byte[] mark27 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark27[h] = buf[91 + h];
                }
                double reactive = BitConverter.ToInt32(mark27.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) reactive = reactive / 100000;
                else if (buf[14] == 3) reactive = reactive / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) reactive = reactive / 10000;
                else if (buf[14] == 2) reactive = reactive / 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) reactive= reactive / 1000;
                else if (buf[14] == 1) reactive = reactive / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) reactive = reactive / 1000;
                else if (buf[14] == 4) reactive = reactive / 100;
                else if (buf[11] == 4 && buf[14] == 0) reactive = reactive / 1000;
                else if (buf[14] == 0) reactive= reactive/ 100;
                else if (buf[14] == 5) reactive = reactive/ 100;
                string str27= Convert.ToString(reactive);
                textBox32.AppendText((str27.Length == 1 ? "0" + str27 : str27));//总无功功率测量值

                byte[] mark20 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark20[h] = buf[95 + h];
                }
                double apparentA = BitConverter.ToInt32(mark20.ToArray(), 0);
                if (buf[9] == 4 && buf[12] == 3) apparentA = apparentA / 100000;
                else if (buf[12] == 3) apparentA = apparentA / 10000;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 2) apparentA = apparentA / 10000;
                else if (buf[12] == 2) apparentA = apparentA / 1000;
                else if ((buf[9] == 2 || buf[9] == 3 || buf[9] == 4) && buf[12] == 1) apparentA = apparentA / 1000;
                else if (buf[12] == 1) apparentA = apparentA / 100;
                else if ((buf[9] == 3 || buf[9] == 4) && buf[12] == 4) apparentA = apparentA / 1000;
                else if (buf[12] == 4) apparentA = apparentA / 100;
                else if (buf[9] == 4 && buf[12] == 0) apparentA = apparentA / 1000;
                else if (buf[12] == 0) apparentA = apparentA / 100;
                else if (buf[12] == 5) apparentA = apparentA / 100;
                string str20 = Convert.ToString(apparentA);
                textBox25.AppendText( (str20.Length == 1 ? "0" + str20 : str20));//A相视在功率测量值
                byte[] mark21 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark21[h] = buf[99 + h];
                }
                double apparentB = BitConverter.ToInt32(mark21.ToArray(), 0);
                if (buf[10] == 4 && buf[13] == 3) apparentB = apparentB / 100000;
                else if (buf[13] == 3) apparentB = apparentB / 10000;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 2) apparentB = apparentB / 10000;
                else if (buf[13] == 2) apparentB = apparentB / 1000;
                else if ((buf[10] == 2 || buf[10] == 3 || buf[10] == 4) && buf[13] == 1) apparentB = apparentB / 1000;
                else if (buf[13] == 1) apparentB = apparentB / 100;
                else if ((buf[10] == 3 || buf[10] == 4) && buf[13] == 4) apparentB = apparentB / 1000;
                else if (buf[13] == 4) apparentB = apparentB / 100;
                else if (buf[10] == 4 && buf[13] == 0) apparentB = apparentB / 1000;
                else if (buf[13] == 0) apparentB = apparentB / 100;
                else if (buf[13] == 5) apparentB = apparentB / 100;
                string str21 = Convert.ToString(apparentB);
                textBox26.AppendText((str21.Length == 1 ? "0" + str21 : str21));//B相视在功率测量值
                byte[] mark22 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark22[h] = buf[103 + h];
                }
                double apparentC = BitConverter.ToInt32(mark22.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) apparentC = apparentC / 100000;
                else if (buf[14] == 3) apparentC = apparentC / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) apparentC = apparentC / 10000;
                else if (buf[14] == 2) apparentC = apparentC / 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) apparentC = apparentC / 1000;
                else if (buf[14] == 1) apparentC = apparentC / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) apparentC = apparentC / 1000;
                else if (buf[14] == 4) apparentC = apparentC / 100;
                else if (buf[11] == 4 && buf[14] == 0) apparentC = apparentC / 1000;
                else if (buf[14] == 0) apparentC = apparentC / 100;
                else if (buf[14] == 5) apparentC = apparentC / 100;
                string str22 = Convert.ToString(apparentC);
                textBox27.AppendText((str22.Length == 1 ? "0" + str22 : str22));//C相视在功率测量值

                byte[] mark28= new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark28[h] = buf[107 + h];
                }
                double apparent = BitConverter.ToInt32(mark28.ToArray(), 0);
                if (buf[11] == 4 && buf[14] == 3) apparent = apparent / 100000;
                else if (buf[14] == 3) apparent = apparent / 10000;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 2) apparent = apparent/ 10000;
                else if (buf[14] == 2) apparent= apparent / 1000;
                else if ((buf[11] == 2 || buf[11] == 3 || buf[11] == 4) && buf[14] == 1) apparent = apparent / 1000;
                else if (buf[14] == 1) apparent = apparent / 100;
                else if ((buf[11] == 3 || buf[11] == 4) && buf[14] == 4) apparent = apparent / 1000;
                else if (buf[14] == 4) apparent = apparent / 100;
                else if (buf[11] == 4 && buf[14] == 0) apparent = apparent / 1000;
                else if (buf[14] == 0) apparent= apparent / 100;
                else if (buf[14] == 5) apparent= apparent / 100;
                string str28= Convert.ToString(apparent);
                textBox33.AppendText((str28.Length == 1 ? "0" + str28 : str28));//总视在功率测量值

                byte[] mark23 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark23[h] = buf[111 + h];
                }
                double powerA = BitConverter.ToInt32(mark23.ToArray(), 0);
                powerA = powerA / 100000;
                string str23 = Convert.ToString(powerA);
                textBox28.AppendText((str23.Length == 1 ? "0" + str23 : str23));//A相功率因数
                byte[] mark24 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark24[h] = buf[115 + h];
                }
                double powerB = BitConverter.ToInt32(mark24.ToArray(), 0);
                powerB = powerB / 100000;
                string str24 = Convert.ToString(powerB);
                textBox29.AppendText((str24.Length == 1 ? "0" + str24 : str24));//B相功率因数
                byte[] mark25 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark25[h] = buf[119 + h];
                }
                double powerC = BitConverter.ToInt32(mark25.ToArray(), 0);
                powerC = powerC / 100000;
                string str25 = Convert.ToString(powerC);
                textBox30.AppendText((str25.Length == 1 ? "0" + str25 : str25));//C相功率因数
                byte[] mark29 = new byte[4];
                for (int h = 0; h < 4; h++)
                {
                    mark29[h] = buf[123 + h];
                }
                double power = BitConverter.ToInt32(mark29.ToArray(), 0);
                power = power / 100000;
                string str29 = Convert.ToString(power);
                textBox34.AppendText((str29.Length == 1 ? "0" + str29 : str29));//总功率因数
              }
            }
            catch
            {

            }
        }

        private void button68_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort2.IsOpen)
                {
                    serialPort2.Close();    //关闭串口
                    button68.Text = "打开串口";
                    button68.BackColor = Color.ForestGreen;
                    comboBox25.Enabled = true;
                    comboBox24.Enabled = true;
                    comboBox23.Enabled = true;
                    comboBox22.Enabled = true;
                    comboBox21.Enabled = true;
                    button62_Click(button62, new EventArgs());
                }
                else
                {
                    serialPort2.PortName = comboBox24.Text;//设置串口号
                    serialPort2.BaudRate = Convert.ToInt32(comboBox25.Text, 10);//十进制数据转换，设置波特率
                    serialPort2.Open();     //打开串口
                    comboBox25.Enabled = false;
                    comboBox24.Enabled = false;
                    comboBox23.Enabled = false;
                    comboBox22.Enabled = false;
                    comboBox21.Enabled = false;
                    button68.Text = "关闭串口";
                    button68.BackColor = Color.Firebrick;
                }
            }
            catch
            {
                MessageBox.Show("端口错误,请检查串口", "错误");
            }
            if (comboBox23.Text.Trim() == "0")    //设置停止位
            {
                serialPort2.StopBits = StopBits.None;
            }
            else if (comboBox23.Text.Trim() == "1.5")
            {
                serialPort2.StopBits = StopBits.OnePointFive;
            }
            else if (comboBox23.Text.Trim() == "2")
            {
                serialPort2.StopBits = StopBits.Two;
            }
            else
            {
                serialPort2.StopBits = StopBits.One;
            }

            serialPort2.DataBits = Convert.ToInt16(comboBox22.Text.Trim());    //设置数据位

            if (comboBox21.Text.Trim() == "奇校验")    //设置校验
            {
                serialPort2.Parity = Parity.Odd;
            }
            else if (comboBox21.Text.Trim() == "偶校验")
            {
                serialPort2.Parity = Parity.Even;
            }
            else
            {
                serialPort2.Parity = Parity.None;
            }
        }

        private void button67_Click(object sender, EventArgs e)
        {
            //textBox5.Text = "";//清屏
        }

        private void button56_Click(object sender, EventArgs e)
        {
            timer2.Stop();
            byte[] Data = new byte[6] { 0x81, 0x00, 0x06, 0x00, 0x54, 0x52 };
            for (int i = 0; i < 3; i++)
            {
                button62_Click(button62, new EventArgs());
                Delay(100);
                Key2_Sentences(Data,6);
            }
            button57.Enabled = true;//关闭电源按钮可用
            button56.Enabled = false;//开启电源按钮不可用
            timer2.Start();
        }

        private void button57_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[6] { 0x81, 0x00, 0x06, 0x00, 0x4f, 0x49 };
            for (int i = 0; i < 5; i++)
            {
                Key2_Sentences(Data, 6);
                Delay(100);
            }
            button56.Enabled = true;//开启电源按钮可用
            button57.Enabled = false;//关闭电源按钮不可用
        }

        private void button58_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[6] { 0x81, 0x00, 0x06, 0x00, 0x4d, 0x4b };
            Key2_Sentences(Data, 6);
        }

        private void button55_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[6] { 0x81, 0x00, 0x06, 0x00, 0x52, 0x54 };
            Key2_Sentences(Data, 6);
            button57.Enabled = true;//关闭电源按钮可用
            button56.Enabled = true;//开启电源按钮可用
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button62_Click(object sender, EventArgs e)
        {
            double a = 0, b = 0;
            int d = 0, f = 0;
            a = Convert.ToDouble(comboBox15.Text);
            b = Convert.ToDouble(comboBox16.Text);
            if (a == 380) d = 0x00; if (a == 220) d = 0x01; if (a == 100) d = 0x02; if (a == 57.7) d = 0x03; if (a == 600) d = 0x05;
            if (b == 20) f = 0x00; if (b == 5) f = 0x01; if (b == 1) f = 0x02; if (b == 0.2) f = 0x03; if (b == 0.05) f = 0x04; if (b == 100) f = 0x05;
            byte[] Data = new byte[12];
            Data[0] = 0X81;
            Data[1] = 00;
            Data[2] = 0X0c;
            Data[3] = 00;
            Data[4] = 0X31;
            Data[5] = (byte)d;
            Data[6] = (byte)f;
            Data[7] = (byte)d;
            Data[8] = (byte)f;
            Data[9] = (byte)d;
            Data[10] = (byte)f;
            Int64 q = 0x0c ^ 0x31 ^ f ^ d;
            Data[11] = (byte)q;

            for (int i = 0; i < 5; i++)
            {
                Key2_Sentences(Data, 12);
                Delay(100);
            }
        }

        private void button60_Click(object sender, EventArgs e)
        {
            double a = 0, b = 0, h = 0, z = 0, f = 0, d = 0, g = 0, u = 0;
            Int64 m = 0, x = 0, y = 0, o = 0, p = 0, l = 0, k = 0, v = 0, r = 0, t = 0, j = 0, n = 0;
            d = Convert.ToDouble(textBox41.Text);//十进制数据转换
            f = Convert.ToDouble(textBox46.Text);//十进制数据转换

            a = Convert.ToDouble(comboBox15.Text);
            b = Convert.ToDouble(comboBox16.Text);

            r = (int)(d * 100);
            l = (int)(f * 100);

            if (a == 100 || a == 220 || a == 380 || a == 600) { r = r * 10;}
            else if ( a == 57.7) { r = r * 100; t = t * 100;}

            if (b == 100 || b == 20) { l = l * 100;}
            else if (b == 5 || b == 1) { l = l * 1000;}
            else if (b == 0.2) { l = l * 10000;}
           
            string str = System.Convert.ToString(r, 2);
            string str1 = System.Convert.ToString(l, 2);

            n = Convert.ToInt32(str, 2);
            m = Convert.ToInt32(str1, 2);

            byte[] src1 = new byte[4];
            src1[0] = (byte)((n >> 24) & 0xFF);
            src1[1] = (byte)((n >> 16) & 0xFF);
            src1[2] = (byte)((n >> 8) & 0xFF);
            src1[3] = (byte)(n & 0xFF);

            byte[] src2 = new byte[4];
            src2[0] = (byte)((m >> 24) & 0xFF);
            src2[1] = (byte)((m >> 16) & 0xFF);
            src2[2] = (byte)((m >> 8) & 0xFF);
            src2[3] = (byte)(m & 0xFF);

            byte[] Data = new byte[30];
            Data[5] = src1[3];
            Data[6] = src1[2];
            Data[7] = src1[1];
            Data[8] = src1[0];

            Data[9] = src1[3];
            Data[10] = src1[2];
            Data[11] = src1[1];
            Data[12] = src1[0];

            Data[13] = src1[3];
            Data[14] = src1[2];
            Data[15] = src1[1];
            Data[16] = src1[0];

            Data[17] = src2[3];
            Data[18] = src2[2];
            Data[19] = src2[1];
            Data[20] = src2[0];

            Data[21] = src2[3];
            Data[22] = src2[2];
            Data[23] = src2[1];
            Data[24] = src2[0];

            Data[25] = src2[3];
            Data[26] = src2[2];
            Data[27] = src2[1];
            Data[28] = src2[0];

            Data[0] = 0X81;
            Data[1] = 00;
            Data[2] = 0X1E;
            Data[3] = 00;
            Data[4] = 0X32;

            Int64 q = 0x1e ^ 0x32 ^ src1[0] ^ src1[1] ^ src1[2] ^ src1[3] ^ src2[0] ^ src2[1] ^ src2[2] ^ src2[3];
            Data[29] = (byte)q;// Data[2] ^ Data[4] ^ Data[5] ^ Data[6] ^ Data[7] ^ Data[8] ^ Data[17] ^ Data[18] ^ Data[19] ^ Data[20];

            for (int i = 0; i < 5; i++)
            {
                Key2_Sentences(Data, 30);
                Delay(100);
            }
        }

        private void comboBox30_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox28_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox27_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox26_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button73_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort3.IsOpen)
                {
                    serialPort3.Close();    //关闭串口
                    button73.Text = "打开串口";
                    button73.BackColor = Color.ForestGreen;
                    comboBox30.Enabled = true;
                    comboBox29.Enabled = true;
                    comboBox28.Enabled = true;
                    comboBox27.Enabled = true;
                    comboBox26.Enabled = true;
                }
                else
                {
                    serialPort3.PortName = comboBox30.Text;//设置串口号
                    serialPort3.BaudRate = Convert.ToInt32(comboBox29.Text, 10);//十进制数据转换，设置波特率
                    serialPort3.Open();     //打开串口
                    comboBox30.Enabled = false;
                    comboBox29.Enabled = false;
                    comboBox28.Enabled = false;
                    comboBox27.Enabled = false;
                    comboBox26.Enabled = false;
                    button73.Text = "关闭串口";
                    button73.BackColor = Color.Firebrick;
                }
            }
            catch
            {
                MessageBox.Show("端口错误,请检查串口", "错误");
            }

            if (comboBox28.Text.Trim() == "0")    //设置停止位
            {
                serialPort3.StopBits = StopBits.None;
            }
            else if (comboBox28.Text.Trim() == "1.5")
            {
                serialPort3.StopBits = StopBits.OnePointFive;
            }
            else if (comboBox28.Text.Trim() == "2")
            {
                serialPort3.StopBits = StopBits.Two;
            }
            else
            {
                serialPort3.StopBits = StopBits.One;
            }

            serialPort3.DataBits = Convert.ToInt16(comboBox27.Text.Trim());    //设置数据位

            if (comboBox26.Text.Trim() == "偶校验")    //设置校验
            {
                serialPort3.Parity = Parity.Odd;
            }
            else if (comboBox26.Text.Trim() == "奇校验")
            {
                serialPort3.Parity = Parity.Even;
            }
            else
            {
                serialPort3.Parity = Parity.None;
            }
        }

        private void button72_Click(object sender, EventArgs e)
        {
            textBox6.Text = "";//清屏
        }

        private void button71_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[20] { 0x68, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x68, 0x03, 0x08, 0x34, 0x33, 0x33, 0x3a, 0x33, 0x33, 0x33, 0x33, 0x7c, 0x16 };
            Key3_Sentences(Data,20);
        }

        private void button70_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[20] { 0x68, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x68, 0x03, 0x08, 0x35, 0x33, 0x33, 0x3a, 0x33, 0x33, 0x33, 0x33, 0x7d, 0x16 };
            Key3_Sentences(Data, 20);
        }

        private void button69_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[20] { 0x68, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x68, 0x03, 0x08, 0x36, 0x33, 0x33, 0x3a, 0x33, 0x33, 0x33, 0x33, 0x7e, 0x16 };
            Key3_Sentences(Data, 20);
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private string function0(Int32 data1)
        {
            string str = Convert.ToString(data1, 16).ToUpper();
            return str;
        }

        private string function1(Int32 data1)
        {
            Int32 x;
            x = data1 & 0x0f;
            string str0 = System.Convert.ToString(x, 16);
            return str0;
        }

        private string function2(Int32 data1)
        {
            Int32 y;
            y = (data1 >> 4) & 0x000f;
            string str0 = System.Convert.ToString(y, 16);
            return str0;
        }

        private string function3(Int32 data1)
        {
            string str0 = System.Convert.ToString(data1, 10);
            return str0;
        }

        private DateTime current_time = new DateTime();

        //public static class CommonRes
        //{
        //    public static SerialPort serialPort3 = new SerialPort();
        //}


        private void serialPort3_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            int ZONE = 0;
            int Value = 0;
            int dataCnt = 0;
            Int32 DWDI = 0x00000000;

            UInt32 DataID = 0;
            const UInt32 readCTWorkCfg = 0x07000001;                           //获取CT互感器工况信息
            const UInt32 readCTVerCfg = 0x07000002;                           //读取CT互感器版本
            const UInt32 readCoreVerCfg = 0x07000003;                           //读取算法核心板版本
            const UInt32 writeCTfileCfg = 0x0f000001;                           //写CT互感器文件
            const UInt32 writeIDCfg = 0x07020005;                           //设置巡检仪核心算法板ID
            const UInt32 readIDCfg = 0x07020006;                           //读取巡检仪核心算法板ID
            const UInt32 writeCTIDCfg = 0x07020003;                           //设置CT互感器ID 
            const UInt32 readCTIDCfg = 0x07020004;                           //读取CT互感器ID 

            /*struct CT_INFO
            {   
                uint   loopState;          //电流回路状态
                Int32  temperature;        //TA专用模块环境温度
                UInt32 fMax;               //相频率最大值
                UInt32 fMin;               //相频率最小值
                UInt32 Ie;                 //相工频电流有效值
                UInt32 Imp_rate_1;         //第一组相回路阻抗频率
                UInt32 Imp_1;              //第一组相回路阻抗模值
                UInt32 Imp_angle_1;        //第一组相回路阻抗角度	

                UInt32 Imp_rate_2;         //第二组相回路阻抗频率
                UInt32 Imp_2;              //第二组相回路阻抗模值
                UInt32 Imp_angle_2;        //第二组相回路阻抗角度

                UInt32 Imp_rate_3;         //第三组相回路阻抗频率
                UInt32 Imp_3;              //第三组相回路阻抗模值
                UInt32 Imp_angle_3;        //第三组相回路阻抗角度	
                uint   Distortion_50Hz;    //畸变率
                uint   KL_Num;             //开路正常数量
                uint   DL_Num;             //短路正常数量
                UInt32 KL_Slope1;          //斜率1
                UInt32 KL_Slope2;          //斜率2
            }*/

            try
            {
                int byteNumber = serialPort3.BytesToRead;
                Delay(20);
                //延时等待数据接收完毕。
                while ((byteNumber < serialPort3.BytesToRead) && (serialPort3.BytesToRead < 4800))
                {
                    byteNumber = serialPort3.BytesToRead;
                    Delay(20);
                }
                int datalen = serialPort3.BytesToRead; //记录下缓冲区的字节个数 
                byte[] databuf = new byte[datalen]; //声明一个临时数组存储当前来的串口数据 
                serialPort3.Read(databuf, 0, datalen); //读取缓冲数据到buf中，同时将这串数据从缓冲区移除

                //设置文字显示
                //foreach (byte Member in databuf)
                //{
                //    string str = Member.ToString("X2");
                //    textBox6.AppendText(str + " ");
                //}
                //textBox6.AppendText("\r\n");

                for (int i = 10; i < (datalen - 2); i++)
                {
                    databuf[i] -= 0x33;
                }

                byte[] buffer = new byte[(datalen - 16)];

                for (int i = 0; i < (datalen - 16); i++)
                {
                    buffer[i] = databuf[i + 14];
                }

                byte[] bufferh = new byte[14];      //
                for (int h = 0; h < 14; h++)
                {
                    bufferh[h] = databuf[h];
                }

                hlxj.String_Decrypt(buffer, (ushort)(datalen - 16));

                //设置文字显示
                //foreach (byte Member in buffer)
                //{
                //    string str = Member.ToString("X2");
                //    textBox6.AppendText(str + " ");
                //}
                //textBox6.AppendText("\r\n");

                DataID = (UInt32)(databuf[10] + (databuf[11] << 8) + (databuf[12] << 16) + (databuf[13] << 24));     //数据标识
                byte Cnt = 0;
                byte Ant = 0;
                switch (DataID)
                {
                    case readCTWorkCfg:

                        #region     //A相工况信息解析
                        switch (buffer[Cnt++])
                        {
                            case 0:
                                textBox6.AppendText("A相CT状态：" + "正常" + "\r\n");wenduArray[0]="正常"; break;
                            case 1:
                                textBox6.AppendText("A相CT状态：" + "短路" + "\r\n");wenduArray[0]="短路"; break;
                            case 2:
                                textBox6.AppendText("A相CT状态：" + "开路" + "\r\n");wenduArray[0]="开路"; break;
                            case 6:
                                textBox6.AppendText("A相CT状态：" + "串接二极管" + "\r\n");wenduArray[0]="串接二极管"; break;
                        }
                        string A_temperature = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_temperature = A_temperature.Insert(3, ".");
                        textBox6.AppendText("A相温度：" + srt_A_temperature + "\r\n");
                        wenduArray[++Ant] = srt_A_temperature;
                        Cnt += 2;

                        string A_fMax = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_fMax = A_fMax.Insert(3, ".");
                        textBox6.AppendText("A相频率最大值：" + srt_A_fMax + "\r\n");
                        wenduArray[++Ant] = srt_A_fMax;
                        Cnt += 3;

                        string A_fMin = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_fMin = A_fMin.Insert(3, ".");
                        textBox6.AppendText("A相频率最小值：" + srt_A_fMin + "\r\n");
                        wenduArray[++Ant] =srt_A_fMin;
                        Cnt += 3;

                        string A_Ie = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Ie = A_Ie.Insert(3, ".");
                        textBox6.AppendText("A相工频电流值：" + srt_A_Ie + "\r\n");
                        switch (standard_stream)
                        {
                            case 0: textBox75.AppendText(srt_A_Ie); break;
                            case 1: textBox76.AppendText(srt_A_Ie); break;
                            case 2: textBox77.AppendText(srt_A_Ie); break;
                            case 3: textBox78.AppendText(srt_A_Ie); break;
                            case 4: textBox79.AppendText(srt_A_Ie); break;
                            case 5: textBox80.AppendText(srt_A_Ie); break;
                        }
                        wenduArray[++Ant] =srt_A_Ie;
                        Cnt += 3;
                        //第一组
                        string A_Imp_rate_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_rate_1 = A_Imp_rate_1.Insert(3, ".");
                        textBox6.AppendText("A相第一组阻抗频率：" + srt_A_Imp_rate_1 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_rate_1;
                        Cnt += 2;

                        string A_Imp_1 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_1 = A_Imp_1.Insert(6, ".");
                        textBox6.AppendText("A相第一组阻抗模值：" + srt_A_Imp_1 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_1;
                        Cnt += 4;

                        string A_Imp_angle_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_IImp_angle_1 = A_Imp_angle_1.Insert(3, ".");
                        textBox6.AppendText("A相第一组阻抗角度：" + srt_A_IImp_angle_1 + "\r\n");
                        wenduArray[++Ant] =srt_A_IImp_angle_1;
                        Cnt += 2;
                        //第二组
                        string A_Imp_rate_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_rate_2 = A_Imp_rate_2.Insert(3, ".");
                        textBox6.AppendText("A相第二组阻抗频率：" + srt_A_Imp_rate_2 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_rate_2;
                        Cnt += 2;

                        string A_Imp_2 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_2 = A_Imp_2.Insert(6, ".");
                        textBox6.AppendText("A相第二组阻抗模值：" + srt_A_Imp_2 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_2;
                        Cnt += 4;

                        string A_Imp_angle_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_IImp_angle_2 = A_Imp_angle_2.Insert(3, ".");
                        textBox6.AppendText("A相第二组阻抗角度：" + srt_A_IImp_angle_2 + "\r\n");
                        wenduArray[++Ant] =srt_A_IImp_angle_2;
                        Cnt += 2;
                        //第三组
                        string A_Imp_rate_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_rate_3 = A_Imp_rate_3.Insert(3, ".");
                        textBox6.AppendText("A相第三组阻抗频率：" + srt_A_Imp_rate_3 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_rate_3;
                        Cnt += 2;

                        string A_Imp_3 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_Imp_3 = A_Imp_3.Insert(6, ".");
                        textBox6.AppendText("A相第三组阻抗模值：" + srt_A_Imp_3 + "\r\n");
                        wenduArray[++Ant] =srt_A_Imp_3;
                        Cnt += 4;

                        string A_Imp_angle_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_IImp_angle_3 = A_Imp_angle_3.Insert(3, ".");
                        textBox6.AppendText("A相第三组阻抗角度：" + srt_A_IImp_angle_3 + "\r\n");
                        wenduArray[++Ant] =srt_A_IImp_angle_3;
                        Cnt += 2;
#if Debug
                        string A_Distortion_50Hz = HexToString(buffer[Cnt]);
                        //string srt_A_Distortion_50Hz = A_Distortion_50Hz.Insert(0, ".");
                        textBox6.AppendText("A相畸变率：" + A_Distortion_50Hz + "\r\n");
                        wenduArray[Ant++] =A_Distortion_50Hz;
                        Cnt += 1;

                        string A_KL_Num = HexToString(buffer[Cnt]);
                        //string srt_A_KL_Num = A_KL_Num.Insert(0, ".");
                        textBox6.AppendText("A相开路正常数量：" + A_KL_Num + "\r\n");
                        wenduArray[Ant++] =A_KL_Num;
                        Cnt += 1;

                        string A_DL_Num = HexToString(buffer[Cnt]);
                        //string srt_A_DL_Num = A_DL_Num.Insert(0, ".");
                        textBox6.AppendText("A相短路正常数量：" + A_DL_Num + "\r\n");
                        wenduArray[Ant++] =A_DL_Num;
                        Cnt += 1;

                        string A_KL_Slope1 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_KL_Slope1 = A_KL_Slope1.Insert(2, ".");
                        textBox6.AppendText("A相斜率1：" + srt_A_KL_Slope1 + "\r\n");
                        wenduArray[Ant++] =srt_A_KL_Slope1;
                        Cnt += 3;

                        string A_KL_Slope2 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_A_KL_Slope2 = A_KL_Slope2.Insert(2, ".");
                        textBox6.AppendText("A相斜率2：" + srt_A_KL_Slope2 + "\r\n");
                        wenduArray[Ant++] = srt_A_KL_Slope2;
                        Cnt += 3;
#endif
                        textBox6.AppendText("\r\n");
                        #endregion

                        #region     //B相工况信息解析
                        switch (buffer[Cnt++])
                        {
                            case 0:
                                textBox6.AppendText("B相CT状态：" + "正常" + "\r\n");wenduArray[++Ant] ="正常"; break;
                            case 1:
                                textBox6.AppendText("B相CT状态：" + "短路" + "\r\n");wenduArray[++Ant] ="短路";break;
                            case 2:
                                textBox6.AppendText("B相CT状态：" + "开路" + "\r\n");wenduArray[++Ant] ="开路"; break;
                            case 6:
                                textBox6.AppendText("B相CT状态：" + "串接二极管" + "\r\n");wenduArray[++Ant] ="串接二极管"; break;
                        }
                        string B_temperature = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_temperature = B_temperature.Insert(3, ".");
                        textBox6.AppendText("B相温度：" + srt_B_temperature + "\r\n");
                        wenduArray[++Ant] =srt_B_temperature;
                        Cnt += 2;

                        string B_fMax = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_fMax = B_fMax.Insert(3, ".");
                        textBox6.AppendText("B相频率最大值：" + srt_B_fMax + "\r\n");
                       wenduArray[++Ant] =srt_B_fMax;
                        Cnt += 3;

                        string B_fMin = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_fMin = B_fMin.Insert(3, ".");
                        textBox6.AppendText("B相频率最小值：" + srt_B_fMin + "\r\n");
                        wenduArray[++Ant] =srt_B_fMin;
                        Cnt += 3;

                        string B_Ie = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Ie = B_Ie.Insert(3, ".");
                        textBox6.AppendText("B相工频电流值：" + srt_B_Ie + "\r\n");
                        switch (standard_stream)
                        {
                            case 0: textBox86.AppendText(srt_B_Ie); break;
                            case 1: textBox85.AppendText(srt_B_Ie); break;
                            case 2: textBox84.AppendText(srt_B_Ie); break;
                            case 3: textBox83.AppendText(srt_B_Ie); break;
                            case 4: textBox82.AppendText(srt_B_Ie); break;
                            case 5: textBox81.AppendText(srt_B_Ie); break;
                        }
                        wenduArray[++Ant] =srt_B_Ie;
                        Cnt += 3;
                        //第一组
                        string B_Imp_rate_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_rate_1 = B_Imp_rate_1.Insert(3, ".");
                        textBox6.AppendText("B相第一组阻抗频率：" + srt_B_Imp_rate_1 + "\r\n");
                       wenduArray[++Ant] =srt_B_Imp_rate_1;
                        Cnt += 2;

                        string B_Imp_1 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_1 = B_Imp_1.Insert(6, ".");
                        textBox6.AppendText("B相第一组阻抗模值：" + srt_B_Imp_1 + "\r\n");
                        wenduArray[++Ant] =srt_B_Imp_1;
                        Cnt += 4;

                        string B_Imp_angle_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_IImp_angle_1 = B_Imp_angle_1.Insert(3, ".");
                        textBox6.AppendText("B相第一组阻抗角度：" + srt_B_IImp_angle_1 + "\r\n");
                       wenduArray[++Ant] =srt_B_IImp_angle_1;
                        Cnt += 2;
                        //第二组
                        string B_Imp_rate_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_rate_2 = B_Imp_rate_2.Insert(3, ".");
                        textBox6.AppendText("B相第二组阻抗频率：" + srt_B_Imp_rate_2 + "\r\n");
                        wenduArray[++Ant] =srt_B_Imp_rate_2;
                        Cnt += 2;

                        string B_Imp_2 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_2 = B_Imp_2.Insert(6, ".");
                        textBox6.AppendText("B相第二组阻抗模值：" + srt_B_Imp_2 + "\r\n");
                        wenduArray[++Ant] =srt_B_Imp_2;
                        Cnt += 4;

                        string B_Imp_angle_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_IImp_angle_2 = B_Imp_angle_2.Insert(3, ".");
                        textBox6.AppendText("B相第二组阻抗角度：" + srt_B_IImp_angle_2 + "\r\n");
                        wenduArray[++Ant] =srt_B_IImp_angle_2;
                        Cnt += 2;
                        //第三组
                        string B_Imp_rate_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_rate_3 = B_Imp_rate_3.Insert(3, ".");
                        textBox6.AppendText("B相第三组阻抗频率：" + srt_B_Imp_rate_3 + "\r\n");
                        wenduArray[++Ant] =srt_B_Imp_rate_3;
                        Cnt += 2;

                        string B_Imp_3 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_Imp_3 = B_Imp_3.Insert(6, ".");
                        textBox6.AppendText("B相第三组阻抗模值：" + srt_B_Imp_3 + "\r\n");
                        wenduArray[++Ant] =srt_B_Imp_3;
                        Cnt += 4;

                        string B_Imp_angle_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_IImp_angle_3 = B_Imp_angle_3.Insert(3, ".");
                        textBox6.AppendText("B相第三组阻抗角度：" + srt_B_IImp_angle_3 + "\r\n");
                        wenduArray[++Ant] =srt_B_IImp_angle_3;
                        Cnt += 2;
#if Debug
                        string B_Distortion_50Hz = HexToString(buffer[Cnt]);
                        //string srt_B_Distortion_50Hz = B_Distortion_50Hz.Insert(0, ".");
                        textBox6.AppendText("B相畸变率：" + B_Distortion_50Hz + "\r\n");
                        wenduArray[Ant++] =B_Distortion_50Hz;
                        Cnt += 1;

                        string B_KL_Num = HexToString(buffer[Cnt]);
                        //string srt_B_KL_Num = B_KL_Num.Insert(0, ".");
                        textBox6.AppendText("B相开路正常数量：" + B_KL_Num + "\r\n");
                        wenduArray[Ant++] =B_KL_Num;
                        Cnt += 1;

                        string B_DL_Num = HexToString(buffer[Cnt]);
                        //string srt_B_DL_Num = B_DL_Num.Insert(0, ".");
                        textBox6.AppendText("B相短路正常数量：" + B_DL_Num + "\r\n");
                        wenduArray[Ant++] =B_DL_Num;
                        Cnt += 1;

                        string B_KL_Slope1 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_KL_Slope1 = B_KL_Slope1.Insert(2, ".");
                        textBox6.AppendText("B相斜率1：" + srt_B_KL_Slope1 + "\r\n");
                        wenduArray[Ant++] =srt_B_KL_Slope1;
                        Cnt += 3;

                        string B_KL_Slope2 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_B_KL_Slope2 = B_KL_Slope2.Insert(2, ".");
                        textBox6.AppendText("B相斜率2：" + srt_B_KL_Slope2 + "\r\n");
                        wenduArray[Ant++] = srt_B_KL_Slope2;
                        Cnt += 3;
#endif
                        textBox6.AppendText("\r\n");
                        #endregion

                        #region     //C相工况信息解析
                        switch (buffer[Cnt++])
                        {
                            case 0:
                                textBox6.AppendText("C相CT状态：" + "正常" + "\r\n");wenduArray[++Ant] ="正常"; break;
                            case 1:
                                textBox6.AppendText("C相CT状态：" + "短路" + "\r\n");wenduArray[++Ant] ="短路"; break;
                            case 2:
                                textBox6.AppendText("C相CT状态：" + "开路" + "\r\n");wenduArray[++Ant] ="开路"; break;
                            case 6:
                                textBox6.AppendText("C相CT状态：" + "串接二极管" + "\r\n");wenduArray[++Ant] ="串接二极管"; break;
                        }
                        string C_temperature = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_temperature = C_temperature.Insert(3, ".");
                        textBox6.AppendText("C相温度：" + srt_C_temperature + "\r\n");
                        wenduArray[++Ant] = srt_C_temperature;
                        Cnt += 2;

                        string C_fMax = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_fMax = C_fMax.Insert(3, ".");
                        textBox6.AppendText("C相频率最大值：" + srt_C_fMax + "\r\n");
                        wenduArray[++Ant] = srt_C_fMax;
                        Cnt += 3;

                        string C_fMin = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_fMin = C_fMin.Insert(3, ".");
                        textBox6.AppendText("C相频率最小值：" + srt_C_fMin + "\r\n");
                        wenduArray[++Ant] = srt_C_fMin;
                        Cnt += 3;

                        string C_Ie = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Ie = C_Ie.Insert(3, ".");
                        textBox6.AppendText("C相工频电流值：" + srt_C_Ie + "\r\n");
                        switch (standard_stream)
                        {
                            case 0: textBox92.AppendText(srt_C_Ie); break;
                            case 1: textBox91.AppendText(srt_C_Ie); break;
                            case 2: textBox90.AppendText(srt_C_Ie); break;
                            case 3: textBox101.AppendText(srt_C_Ie); break;
                            case 4: textBox88.AppendText(srt_C_Ie); break;
                            case 5: textBox87.AppendText(srt_C_Ie); break;
                        }
                        wenduArray[++Ant] =srt_C_Ie ;
                        Cnt += 3;
                        //第一组
                        string C_Imp_rate_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_rate_1 = C_Imp_rate_1.Insert(3, ".");
                        textBox6.AppendText("C相第一组阻抗频率：" + srt_C_Imp_rate_1 + "\r\n");
                        wenduArray[++Ant] =srt_C_Imp_rate_1 ;
                        Cnt += 2;

                        string C_Imp_1 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_1 = C_Imp_1.Insert(6, ".");
                        textBox6.AppendText("C相第一组阻抗模值：" + srt_C_Imp_1 + "\r\n");
                        wenduArray[++Ant] = srt_C_Imp_1;
                        Cnt += 4;

                        string C_Imp_angle_1 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_IImp_angle_1 = C_Imp_angle_1.Insert(3, ".");
                        textBox6.AppendText("C相第一组阻抗角度：" + srt_C_IImp_angle_1 + "\r\n");
                        wenduArray[++Ant] =srt_C_IImp_angle_1 ;
                        Cnt += 2;
                        //第二组
                        string C_Imp_rate_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_rate_2 = C_Imp_rate_2.Insert(3, ".");
                        textBox6.AppendText("C相第二组阻抗频率：" + srt_C_Imp_rate_2 + "\r\n");
                        wenduArray[++Ant] =srt_C_Imp_rate_2 ;
                        Cnt += 2;

                        string C_Imp_2 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_2 = C_Imp_2.Insert(6, ".");
                        textBox6.AppendText("C相第二组阻抗模值：" + srt_C_Imp_2 + "\r\n");
                        wenduArray[++Ant] =srt_C_Imp_2 ;
                        Cnt += 4;

                        string C_Imp_angle_2 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_IImp_angle_2 = C_Imp_angle_2.Insert(3, ".");
                        textBox6.AppendText("C相第二组阻抗角度：" + srt_C_IImp_angle_2 + "\r\n");
                        wenduArray[++Ant] = srt_C_IImp_angle_2;
                        Cnt += 2;
                        //第三组
                        string C_Imp_rate_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_rate_3 = C_Imp_rate_3.Insert(3, ".");
                        textBox6.AppendText("C相第三组阻抗频率：" + srt_C_Imp_rate_3 + "\r\n");
                        wenduArray[++Ant] = srt_C_Imp_rate_3;
                        Cnt += 2;

                        string C_Imp_3 = HexToString(buffer[Cnt + 3]) + HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_Imp_3 = C_Imp_3.Insert(6, ".");
                        textBox6.AppendText("C相第三组阻抗模值：" + srt_C_Imp_3 + "\r\n");
                        wenduArray[++Ant] =srt_C_Imp_3 ;
                        Cnt += 4;

                        string C_Imp_angle_3 = HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_IImp_angle_3 = C_Imp_angle_3.Insert(3, ".");
                        textBox6.AppendText("C相第三组阻抗角度：" + srt_C_IImp_angle_3 );
                        wenduArray[++Ant] =srt_C_IImp_angle_3 ;
                        Cnt += 2;
#if Debug
                        string C_Distortion_50Hz = HexToString(buffer[Cnt]);
                        //string srt_C_Distortion_50Hz = C_Distortion_50Hz.Insert(0, ".");
                        textBox6.AppendText("C相畸变率：" + C_Distortion_50Hz + "\r\n");
                        wenduArray[Ant++] =C_Distortion_50Hz ;
                        Cnt += 1;

                        string C_KL_Num = HexToString(buffer[Cnt]);
                        //string srt_C_KL_Num = C_KL_Num.Insert(0, ".");
                        textBox6.AppendText("C相开路正常数量：" + C_KL_Num + "\r\n");
                        wenduArray[Ant++] = C_KL_Num;
                        Cnt += 1;

                        string C_DL_Num = HexToString(buffer[Cnt]);
                        //string srt_C_DL_Num = C_DL_Num.Insert(0, ".");
                        textBox6.AppendText("C相短路正常数量：" + C_DL_Num + "\r\n");
                        wenduArray[Ant++] = C_DL_Num;
                        Cnt += 1;

                        string C_KL_Slope1 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_KL_Slope1 = C_KL_Slope1.Insert(2, ".");
                        textBox6.AppendText("C相斜率1：" + srt_C_KL_Slope1 + "\r\n");
                        wenduArray[Ant++] = srt_C_KL_Slope1;
                        Cnt += 3;

                        string C_KL_Slope2 = HexToString(buffer[Cnt + 2]) + HexToString(buffer[Cnt + 1]) + HexToString(buffer[Cnt]);
                        string srt_C_KL_Slope2 = C_KL_Slope2.Insert(2, ".");
                        textBox6.AppendText("C相斜率2：" + srt_C_KL_Slope2 + "\r\n");
                        wenduArray[Ant++] = srt_C_KL_Slope2;
                        Cnt += 3;
#endif
                        textBox6.AppendText("\r\n");
                        #endregion

                        break;
                    case readCTVerCfg:

                        #region      //A相CT互感器版本信息解析
                        string srt_A_module = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("A相互感器型号：" + srt_A_module + "\r\n");
                        Cnt += 24;

                        string srt_A_moduleID = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("A相互感器ID：" + srt_A_moduleID + "\r\n");
                        Cnt += 24;

                        string srt_A_FactoryID = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("A相厂商代码：" + srt_A_FactoryID + "\r\n");
                        Cnt += 4;

                        string srt_A_SoftVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("A相软件版本号：" + srt_A_SoftVersion + "\r\n");
                        Cnt += 4;

                        string srt_A_SoftVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("A相软件版本号：" + srt_A_SoftVerDate + "\r\n");
                        Cnt += 3;

                        string srt_A_HardVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("A相硬件版本号：" + srt_A_HardVersion + "\r\n");
                        Cnt += 4;

                        string srt_A_HardVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("A相硬件版本日期：" + srt_A_HardVerDate + "\r\n");
                        Cnt += 3;

                        string srt_A_ExtVerDate = ArrayTosString(buffer, Cnt, 8);
                        textBox6.AppendText("A相厂家扩展信息：" + srt_A_ExtVerDate + "\r\n");
                        Cnt += 8;
                        #endregion

                        #region    //B相CT互感器版本信息解析
                        string srt_B_module = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("B相互感器型号：" + srt_B_module + "\r\n");
                        Cnt += 24;

                        string srt_B_moduleID = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("B相互感器ID：" + srt_B_moduleID + "\r\n");
                        Cnt += 24;

                        string srt_B_FactoryID = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("B相厂商代码：" + srt_B_FactoryID + "\r\n");
                        Cnt += 4;

                        string srt_B_SoftVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("B相软件版本号：" + srt_B_SoftVersion + "\r\n");
                        Cnt += 4;

                        string srt_B_SoftVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("B相软件版本号：" + srt_B_SoftVerDate + "\r\n");
                        Cnt += 3;

                        string srt_B_HardVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("B相硬件版本号：" + srt_B_HardVersion + "\r\n");
                        Cnt += 4;

                        string srt_B_HardVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("B相硬件版本日期：" + srt_B_HardVerDate + "\r\n");
                        Cnt += 3;

                        string srt_B_ExtVerDate = ArrayTosString(buffer, Cnt, 8);
                        textBox6.AppendText("B相厂家扩展信息：" + srt_B_ExtVerDate + "\r\n");
                        Cnt += 8;
                        #endregion

                        #region        //C相CT互感器版本信息解析
                        string srt_C_module = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("C相互感器型号：" + srt_C_module + "\r\n");
                        Cnt += 24;

                        string srt_C_moduleID = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("C相互感器ID：" + srt_C_moduleID + "\r\n");
                        Cnt += 24;

                        string srt_C_FactoryID = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("C相厂商代码：" + srt_C_FactoryID + "\r\n");
                        Cnt += 4;

                        string srt_C_SoftVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("C相软件版本号：" + srt_C_SoftVersion + "\r\n");
                        Cnt += 4;

                        string srt_C_SoftVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("C相软件版本号：" + srt_C_SoftVerDate + "\r\n");
                        Cnt += 3;

                        string srt_C_HardVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("C相硬件版本号：" + srt_C_HardVersion + "\r\n");
                        Cnt += 4;

                        string srt_C_HardVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("C相硬件版本日期：" + srt_C_HardVerDate + "\r\n");
                        Cnt += 3;

                        string srt_C_ExtVerDate = ArrayTosString(buffer, Cnt, 8);
                        textBox6.AppendText("C相厂家扩展信息：" + srt_C_ExtVerDate + "\r\n");
                        Cnt += 8;
                        #endregion

                        break;
                    case readCoreVerCfg:

                        #region      //算法核心板版本信息解析
                        string srt_Arith_module = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("算法板互感器型号：" + srt_Arith_module + "\r\n");
                        Cnt += 24;

                        string srt_Arith_moduleID = ArrayTosString(buffer, Cnt, 24);
                        textBox6.AppendText("算法板互感器ID：" + srt_Arith_moduleID + "\r\n");
                        Cnt += 24;

                        string srt_Arith_FactoryID = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("算法板厂商代码：" + srt_Arith_FactoryID + "\r\n");
                        Cnt += 4;

                        string srt_Arith_SoftVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("算法板软件版本号：" + srt_Arith_SoftVersion + "\r\n");
                        Cnt += 4;

                        string srt_Arith_SoftVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("算法板软件版本号：" + srt_Arith_SoftVerDate + "\r\n");
                        Cnt += 3;

                        string srt_Arith_HardVersion = ArrayTosString(buffer, Cnt, 4);
                        textBox6.AppendText("算法板硬件版本号：" + srt_Arith_HardVersion + "\r\n");
                        Cnt += 4;

                        string srt_Arith_HardVerDate = ArrayTosString(buffer, Cnt, 3);
                        textBox6.AppendText("算法板硬件版本日期：" + srt_Arith_HardVerDate + "\r\n");
                        Cnt += 3;

                        string srt_Arith_ExtVerDate = ArrayTosString(buffer, Cnt, 8);
                        textBox6.AppendText("算法板厂家扩展信息：" + srt_Arith_ExtVerDate + "\r\n");
                        Cnt += 8;
                        #endregion

                        break;
                    case readIDCfg:

                        #region   //读取算法板ID解析
                        byte[] ModuleID = new byte[24];
                        byte[] ModuleInID = new byte[24];
                        for (int num = 0; num < 24; num++)
                        {
                            ModuleID[num] = buffer[num];
                            ModuleInID[num] = buffer[num + 24];
                        }
                        GetIDtext.Text = " ";                             //清空
                        GetInIDtext.Text = " ";                           //清空
                        DLT64507.rever_char(ref ModuleID, ModuleID.Length);
                        DLT64507.rever_char(ref ModuleInID, ModuleID.Length);
                        string str_ModuleID = BitConverter.ToString(ModuleID).Replace("-", "");
                        string str_ModuleInID = BitConverter.ToString(ModuleInID).Replace("-", "");
                        GetIDtext.Text = str_ModuleID;
                        GetInIDtext.Text = str_ModuleInID;
                        textBox38.AppendText("算法板ModuleID:" + str_ModuleID + "\r\n");
                        textBox38.AppendText("算法板ModuleInID:" + str_ModuleInID + "\r\n");
                        #endregion

                        break;
                    case readCTIDCfg:

                        #region    //读取互感器ID
                        byte[] A_CtID = new byte[24];
                        byte[] A_CtInID = new byte[24];
                        byte[] A_CtbarID = new byte[24];
                        byte[] B_CtID = new byte[24];
                        byte[] B_CtInID = new byte[24];
                        byte[] B_CtbarID = new byte[24];
                        byte[] C_CtID = new byte[24];
                        byte[] C_CtInID = new byte[24];
                        byte[] C_CtbarID = new byte[24];

                        for (int num = 0; num < 24; num++)
                        {
                            A_CtID[num] = buffer[num];
                            A_CtInID[num] = buffer[num + 24];
                            A_CtbarID[num] = buffer[num + 24*2];
                            B_CtID[num] = buffer[num + 24 * 3];
                            B_CtInID[num] = buffer[num + 24 * 4];
                            B_CtbarID[num] = buffer[num + 24 * 5];
                            C_CtID[num] = buffer[num + 24 * 6];
                            C_CtInID[num] = buffer[num + 24 * 7];
                            C_CtbarID[num] = buffer[num + 24 * 8];
                        }
                        //A相
                        textBox42.Text = " ";                          //清空
                        textBox37.Text = " ";                          //清空
                        textBox43.Text = " ";                          //清空
                        DLT64507.rever_char(ref A_CtID, A_CtID.Length);
                        DLT64507.rever_char(ref A_CtInID, A_CtInID.Length);
                        DLT64507.rever_char(ref A_CtbarID, A_CtbarID.Length);
                        string str_A_CtID = BitConverter.ToString(A_CtID).Replace("-", "");
                        string str_A_CtInID = BitConverter.ToString(A_CtInID).Replace("-", "");
                        string str_A_CtbarID = BitConverter.ToString(A_CtbarID).Replace("-", "");
                        textBox42.Text = str_A_CtID;
                        textBox37.Text = str_A_CtInID;
                        textBox43.Text = str_A_CtbarID;
                        textBox38.AppendText("A相ID" + "\r\n");
                        textBox38.AppendText("A_CtID:" + str_A_CtID + "\r\n");
                        textBox38.AppendText("A_CtInID:" + str_A_CtInID + "\r\n");
                        textBox38.AppendText("A_CtbarID:" + str_A_CtbarID + "\r\n");
                        //B相
                        textBox36.Text = " ";                           //清空
                        textBox2.Text = " ";                           //清空
                        textBox44.Text = " ";                           //清空
                        DLT64507.rever_char(ref B_CtID, B_CtID.Length);
                        DLT64507.rever_char(ref B_CtInID, B_CtInID.Length);
                        DLT64507.rever_char(ref B_CtbarID, B_CtbarID.Length);
                        string str_B_CtID = BitConverter.ToString(B_CtID).Replace("-", "");
                        string str_B_CtInID = BitConverter.ToString(B_CtInID).Replace("-", "");
                        string str_B_CtbarID = BitConverter.ToString(B_CtbarID).Replace("-", "");
                        textBox36.Text = str_B_CtID;
                        textBox2.Text = str_B_CtInID;
                        textBox44.Text = str_B_CtbarID;
                        textBox38.AppendText("B相ID" + "\r\n");
                        textBox38.AppendText("B_CtID:" + str_B_CtID + "\r\n");
                        textBox38.AppendText("B_CtInID:" + str_B_CtInID + "\r\n");
                        textBox38.AppendText("B_CtInID:" + str_B_CtbarID + "\r\n");
                        //C相
                        textBox3.Text = " ";                           //清空
                        textBox4.Text = " ";                           //清空
                        textBox45.Text = " ";                           //清空
                        DLT64507.rever_char(ref C_CtID, C_CtID.Length);
                        DLT64507.rever_char(ref C_CtInID, C_CtInID.Length);
                        DLT64507.rever_char(ref C_CtbarID, C_CtbarID.Length);
                        string str_C_CtID = BitConverter.ToString(C_CtID).Replace("-", "");
                        string str_C_CtInID = BitConverter.ToString(C_CtInID).Replace("-", "");
                        string str_C_CtbarID = BitConverter.ToString(C_CtbarID).Replace("-", "");
                        textBox3.Text = str_C_CtID;
                        textBox4.Text = str_C_CtInID;
                        textBox45.Text = str_C_CtbarID;
                        textBox38.AppendText("C相ID" + "\r\n");
                        textBox38.AppendText("C_CtID:" + str_C_CtID + "\r\n");
                        textBox38.AppendText("C_CtInID:" + str_C_CtInID + "\r\n");
                        textBox38.AppendText("C_CtInID:" + str_C_CtbarID + "\r\n");
                        #endregion

                        break;

                    case 0x07020002:                      
                        #region//库参数配置
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_CUR25.Text = " ";
                        A_CUR25.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_CUR650.Text = " ";
                        A_CUR650.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_CUR2000.Text = " ";
                        A_CUR2000.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_5_D.Text = " ";
                        A_Amp_5_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_5_U.Text = " ";
                        A_Amp_5_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_10_D.Text = " ";
                        A_Amp_10_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_10_U.Text = " ";
                        A_Amp_10_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_15_D.Text = " ";
                        A_Amp_15_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_Amp_15_U.Text = " ";
                        A_Amp_15_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_KL_NUM.Text = " ";
                        A_KL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_DL_NUM.Text = " ";
                        A_DL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_FRELT_DOWN.Text = " ";
                        A_FRELT_DOWN.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        A_GPCUR_ITHD.Text = " ";
                        A_GPCUR_ITHD.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        A_KL_SCANFRE.Text = " ";
                        A_KL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        A_DL_SCANFRE.Text = " ";
                        A_DL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_KL_GAIN_K1.Text = " ";
                        A_KL_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_KL_GAIN_K2.Text = " ";
                        A_KL_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_ZC_GAIN_K1.Text = " ";
                        A_ZC_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_ZC_GAIN_K2.Text = " ";
                        A_ZC_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_LS1.Text = " ";
                        AIRON_CORE1_LS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_RS1.Text = " ";
                        AIRON_CORE1_RS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_LS2.Text = " ";
                        AIRON_CORE1_LS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_RS2.Text = " ";
                        AIRON_CORE1_RS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_LS3.Text = " ";
                        AIRON_CORE1_LS3.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        AIRON_CORE1_RS3.Text = " ";
                        AIRON_CORE1_RS3.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_CUR25.Text = " ";
                        B_CUR25.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_CUR650.Text = " ";
                        B_CUR650.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_CUR2000.Text = " ";
                        B_CUR2000.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_5_D.Text = " ";
                        B_Amp_5_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_5_U.Text = " ";
                        B_Amp_5_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_10_D.Text = " ";
                        B_Amp_10_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_10_U.Text = " ";
                        B_Amp_10_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_15_D.Text = " ";
                        B_Amp_15_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_Amp_15_U.Text = " ";
                        B_Amp_15_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_KL_NUM.Text = " ";
                        B_KL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_DL_NUM.Text = " ";
                        B_DL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_FRELT_DOWN.Text = " ";
                        B_FRELT_DOWN.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        B_GPCUR_ITHD.Text = " ";
                        B_GPCUR_ITHD.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        B_KL_SCANFRE.Text = " ";
                        B_KL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        B_DL_SCANFRE.Text = " ";
                        B_DL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_KL_GAIN_K1.Text = " ";
                        B_KL_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_KL_GAIN_K2.Text = " ";
                        B_KL_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_ZC_GAIN_K1.Text = " ";
                        B_ZC_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_ZC_GAIN_K2.Text = " ";
                        B_ZC_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                          Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_LS1.Text = " ";
                        BIRON_CORE1_LS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_RS1.Text = " ";
                        BIRON_CORE1_RS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_LS2.Text = " ";
                        BIRON_CORE1_LS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_RS2.Text = " ";
                        BIRON_CORE1_RS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_LS3.Text = " ";
                        BIRON_CORE1_LS3.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BIRON_CORE1_RS3.Text = " ";
                        BIRON_CORE1_RS3.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_CUR25.Text = " ";
                        C_CUR25.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_CUR650.Text = " ";
                        C_CUR650.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_CUR2000.Text = " ";
                        C_CUR2000.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_5_D.Text = " ";
                        C_Amp_5_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_5_U.Text = " ";
                        C_Amp_5_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_10_D.Text = " ";
                        C_Amp_10_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_10_U.Text = " ";
                        C_Amp_10_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_15_D.Text = " ";
                        C_Amp_15_D.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_Amp_15_U.Text = " ";
                        C_Amp_15_U.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_KL_NUM.Text = " ";
                        C_KL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_DL_NUM.Text = " ";
                        C_DL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_FRELT_DOWN.Text = " ";
                        C_FRELT_DOWN.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        C_GPCUR_ITHD.Text = " ";
                        C_GPCUR_ITHD.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        C_KL_SCANFRE.Text = " ";
                        C_KL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        C_DL_SCANFRE.Text = " ";
                        C_DL_SCANFRE.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_KL_GAIN_K1.Text = " ";
                        C_KL_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_KL_GAIN_K2.Text = " ";
                        C_KL_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_ZC_GAIN_K1.Text = " ";
                        C_ZC_GAIN_K1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_ZC_GAIN_K2.Text = " ";
                        C_ZC_GAIN_K2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                          Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_LS1.Text = " ";
                        CIRON_CORE1_LS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_RS1.Text = " ";
                        CIRON_CORE1_RS1.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_LS2.Text = " ";
                        CIRON_CORE1_LS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_RS2.Text = " ";
                        CIRON_CORE1_RS2.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_LS3.Text = " ";
                        CIRON_CORE1_LS3.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CIRON_CORE1_RS3.Text = " ";
                        CIRON_CORE1_RS3.Text = ByteOperate.inttostr(Value).Replace("-", "");
 #endregion
                        break;
                    case 0x07020008:
                        #region//精度校准参数
                        Value = buffer[dataCnt++];
                        A_VPP.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_KL_LIMIT.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_POWER_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        A_POWER_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_POWER_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_POWER_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_HIGHT_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        A_HIGHT_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_HIGHT_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_HIGHT_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_DL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        A_KL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        B_VPP.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_KL_LIMIT.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_POWER_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        B_POWER_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_POWER_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_POWER_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_HIGHT_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        B_HIGHT_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_HIGHT_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_HIGHT_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_DL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        B_KL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt++];
                        C_VPP.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_KL_LIMIT.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_POWER_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        C_POWER_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_POWER_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_POWER_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_HIGHT_L_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;

                        C_HIGHT_L_S.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_HIGHT_H_K.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_HIGHT_H_S.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_DL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        C_KL_value.Text = ByteOperate.inttostr(Value).Replace("-", "");
 #endregion
                        break;
                    case 0x0702000a:
                        #region
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_POWER_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox18.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if(Value<2100&&Value>2000) label26.BackColor = Color.ForestGreen;
                        else label26.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_POWER_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox17.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label27.BackColor = Color.ForestGreen;
                        else label27.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_POWER_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox16.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label28.BackColor = Color.ForestGreen;
                        else label28.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_POWER_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label29.BackColor = Color.ForestGreen;
                        else label29.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_POWER_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox14.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label30.BackColor = Color.ForestGreen;
                        else label30.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_POWER_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox13.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label31.BackColor = Color.ForestGreen;
                        else label31.BackColor = Color.Red;
                         #endregion
                        break;
                    case 0x0702000c:
                        #region
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_HIGHT_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox50.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label250.BackColor = Color.ForestGreen;
                        else label250.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        A_HIGHT_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox49.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label249.BackColor = Color.ForestGreen;
                        else label249.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_HIGHT_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox48.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label248.BackColor = Color.ForestGreen;
                        else label248.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        B_HIGHT_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox47.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label247.BackColor = Color.ForestGreen;
                        else label247.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_HIGHT_H1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox40.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label246.BackColor = Color.ForestGreen;
                        else label246.BackColor = Color.Red;

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        C_HIGHT_H5.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        textBox39.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (Value < 2100 && Value > 2000) label245.BackColor = Color.ForestGreen;
                        else label245.BackColor = Color.Red;
                         #endregion
                        break;
                    case 0x0702000f:
                        #region//抄读校准数据
                        //A相工频电流畸变率(HH)
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        A_aberration.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button9.BackColor == Color.ForestGreen) textBox74.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[0] = Convert.ToString(Value);
                        //A相开路扫频统计正常数(HH)
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        ADL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相短路扫频统计正常数(HH)
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        AKL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相Vpp_5K（XX.XXXX）
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        AVpp_5K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) AKDOWN05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi==1) AKUP05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button7.BackColor == Color.ForestGreen && biaozhi==1) textBox68.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox65.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相Vpp_10K（XX.XXXX）
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        AVpp_10K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) AKDOWN10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) AKUP10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相Vpp_15K（XX.XX）
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        AVpp_15K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) AKDOWN15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) AKUP15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相频率最小值
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        AMin_frequency.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox71.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相斜率1
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Arake1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button65.BackColor == Color.ForestGreen) textBox53.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox53.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相斜率2
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Arake2.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button65.BackColor == Color.ForestGreen) textBox56.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox56.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //A相Th_value
                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        ATh_value.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        //B相工频电流畸变率(HH)
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        B_aberration.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button9.BackColor == Color.ForestGreen) textBox73.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //HH
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        BDL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //HH
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        BKL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        BVpp_5K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) BKDOWN05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) BKUP05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button7.BackColor == Color.ForestGreen && biaozhi == 1) textBox67.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox64.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        BVpp_10K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) BKDOWN10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) BKUP10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        BVpp_15K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) BKDOWN15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) BKUP15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        BMin_frequency.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox70.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Brake1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button89.BackColor == Color.ForestGreen) textBox52.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox52.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Brake2.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button89.BackColor == Color.ForestGreen) textBox55.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox55.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        BTh_value.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        //C相工频电流畸变率(HH)
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        C_aberration.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button9.BackColor == Color.ForestGreen) textBox72.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //HH
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        CDL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        //HH
                        Value = buffer[dataCnt];
                        dataCnt += 1;
                        CKL_NUM.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        CVpp_5K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) CKDOWN05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) CKUP05.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button7.BackColor == Color.ForestGreen && biaozhi == 1) textBox66.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox63.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        CVpp_10K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) CKDOWN10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) CKUP10.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16) + (buffer[dataCnt + 3] << 24);
                        dataCnt += 4;
                        CVpp_15K.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button64.BackColor == Color.ForestGreen) CKDOWN15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        else if (button7.BackColor == Color.ForestGreen && biaozhi == 1) CKUP15.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        CMin_frequency.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen && biaozhi == 1) textBox69.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Crake1.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button94.BackColor == Color.ForestGreen) textBox51.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox51.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8) + (buffer[dataCnt + 2] << 16);
                        dataCnt += 3;
                        Crake2.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button94.BackColor == Color.ForestGreen) textBox54.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        if (button6.BackColor == Color.ForestGreen) textBox54.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);

                        Value = buffer[dataCnt] + (buffer[dataCnt + 1] << 8);
                        dataCnt += 2;
                        CTh_value.Text = ByteOperate.inttostr(Value).Replace("-", "");
                        shiduArray[++ZONE] = Convert.ToString(Value);
                        KEEP_Click(KEEP, new EventArgs());
                         #endregion
                        break;
                }

                if (databuf[0] == 0x68 && databuf[7] == 0x68 && databuf[8] == 0xC3 && datalen == 14)
                {
                    if ((databuf[11] & 0x20) == 0x20) textBox6.AppendText("无效数据");
                    else if ((databuf[11] & 0x02) == 0x02) textBox6.AppendText("无请求数据");
                    else if ((databuf[11] & 0x01) == 0x01) textBox6.AppendText("其它错误");
                    textBox6.AppendText("\r\n");
                }
                Array.Clear(buffer, 0, buffer.Length);
                Array.Clear(databuf, 0, databuf.Length);
                if (DataID == 0x07000001 || DataID == 0x07000002 || DataID == 0x07000003)
                {
                 if (checkBox1.Checked)
                 {
                    //显示时间
                    current_time = System.DateTime.Now;     //获取当前时间
                    textBox6.AppendText(current_time.ToString("时间" + "  " + "yyyy-MM-dd HH:mm:ss"));
                    textBox6.AppendText("\r\n" + "\r\n");
                 }
                 else
                 {
                    textBox6.AppendText("\r\n"); //不显示时间
                 }
                }
                if (DataID == 0x07020002 || DataID == 0x07020008 || DataID == 0x0702000a || DataID == 0x0702000c || DataID == 0x0702000f || DataID == 0x07020004 || DataID == 0x07020006)
                {
                    if (checkBox1.Checked)
                    {
                        //显示时间
                        current_time = System.DateTime.Now;     //获取当前时间
                        textBox38.AppendText(current_time.ToString("时间" + "  " + "yyyy-MM--dd HH:mm:ss"));
                        textBox38.AppendText("\r\n" + "\r\n");
                    }
                    else
                    {
                        textBox38.AppendText("\r\n"); //不显示时间
                    }
                }
            }
            catch
            {
                //MessageBox.Show("帧错误/解析函数错误");
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //定时时间到
            button71_Click(button71, new EventArgs());    //调用发送按钮回调函数
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                //自动发送功能选中,开始自动发送
                numericUpDown1.Enabled = false;     //失能时间选择
                timer1.Interval = (int)numericUpDown1.Value;     //定时器赋初值
                timer1.Start();     //启动定时器
            }
            else
            {
                //自动发送功能未选中,停止自动发送
                numericUpDown1.Enabled = true;     //使能时间选择
                timer1.Stop();     //停止定时器
            }
        }

        #region 空
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox27_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox29_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox38_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox40_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox41_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox42_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox43_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            button58_Click(button58, new EventArgs());
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }
        #endregion
         private string[] wenduArray = new string[42] ;

        private void button67_Click_1(object sender, EventArgs e)
        {
            dataGridView1.GridColor = Color.Blue;//设置网格颜色
            int index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = current_time.ToString("yyyy-MM--dd HH:mm:ss");
            this.dataGridView1.Rows[index].Cells[1].Value = wenduArray[0];
            this.dataGridView1.Rows[index].Cells[2].Value = wenduArray[1];
            this.dataGridView1.Rows[index].Cells[3].Value = wenduArray[2];
            this.dataGridView1.Rows[index].Cells[4].Value = wenduArray[3];
            this.dataGridView1.Rows[index].Cells[5].Value = wenduArray[4];
            this.dataGridView1.Rows[index].Cells[6].Value = wenduArray[5];
            this.dataGridView1.Rows[index].Cells[7].Value = wenduArray[6];
            this.dataGridView1.Rows[index].Cells[8].Value = wenduArray[7];
            this.dataGridView1.Rows[index].Cells[9].Value = wenduArray[8];
            this.dataGridView1.Rows[index].Cells[10].Value = wenduArray[9];
            this.dataGridView1.Rows[index].Cells[11].Value = wenduArray[10];
            this.dataGridView1.Rows[index].Cells[12].Value = wenduArray[11];
            this.dataGridView1.Rows[index].Cells[13].Value = wenduArray[12];
            this.dataGridView1.Rows[index].Cells[14].Value = wenduArray[13];
            this.dataGridView1.Rows[index].Cells[15].Value = wenduArray[14];
            this.dataGridView1.Rows[index].Cells[16].Value = wenduArray[15];
            this.dataGridView1.Rows[index].Cells[17].Value = wenduArray[16];
            this.dataGridView1.Rows[index].Cells[18].Value = wenduArray[17];
            this.dataGridView1.Rows[index].Cells[19].Value = wenduArray[18];
            this.dataGridView1.Rows[index].Cells[20].Value = wenduArray[19];
            this.dataGridView1.Rows[index].Cells[21].Value = wenduArray[20];
            this.dataGridView1.Rows[index].Cells[22].Value = wenduArray[21];
            this.dataGridView1.Rows[index].Cells[23].Value = wenduArray[22];
            this.dataGridView1.Rows[index].Cells[24].Value = wenduArray[23];
            this.dataGridView1.Rows[index].Cells[25].Value = wenduArray[24];
            this.dataGridView1.Rows[index].Cells[26].Value = wenduArray[25];
            this.dataGridView1.Rows[index].Cells[27].Value = wenduArray[26];
            this.dataGridView1.Rows[index].Cells[28].Value = wenduArray[27];
            this.dataGridView1.Rows[index].Cells[29].Value = wenduArray[28];
            this.dataGridView1.Rows[index].Cells[30].Value = wenduArray[29];
            this.dataGridView1.Rows[index].Cells[31].Value = wenduArray[30];
            this.dataGridView1.Rows[index].Cells[32].Value = wenduArray[31];
            this.dataGridView1.Rows[index].Cells[33].Value = wenduArray[32];
            this.dataGridView1.Rows[index].Cells[34].Value = wenduArray[33];
            this.dataGridView1.Rows[index].Cells[35].Value = wenduArray[34];
            this.dataGridView1.Rows[index].Cells[36].Value = wenduArray[35];
            this.dataGridView1.Rows[index].Cells[37].Value = wenduArray[36];
            this.dataGridView1.Rows[index].Cells[38].Value = wenduArray[37];
            this.dataGridView1.Rows[index].Cells[39].Value = wenduArray[38];
            this.dataGridView1.Rows[index].Cells[40].Value = wenduArray[39];
            this.dataGridView1.Rows[index].Cells[41].Value = wenduArray[40];
            this.dataGridView1.Rows[index].Cells[42].Value = wenduArray[41];
        }

        private void button77_Click(object sender, EventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Excel文件|*.xlsx|Excel(2003文件)|*.xls";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = fileDialog.FileName;
                Excel.Application application = new Excel.Application();
                Excel.Workbooks workbooks = application.Workbooks;
                Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = workbook.Worksheets[1];
                int colIndex = 0;
                worksheet.Rows[1].RowHeight = 20; //第一行行高为60（单位：磅）
                for (int i = 1; i < 44; i++)
                {
                    worksheet.Columns[i].ColumnWidth = 25;
                }
                //导出DataGridView中的标题
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    if (dataGridView1.Columns[i].Visible)//做同于不导出隐藏列
                    {
                        colIndex++;
                        worksheet.Cells[1, colIndex] = dataGridView1.Columns[i].HeaderText;
                    }
                }
                //导出DataGridView中的数据
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    colIndex = 0;
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        if (dataGridView1.Columns[j].Visible)
                        {
                            colIndex++;
                            worksheet.Cells[i + 2, colIndex] = "'" + dataGridView1.Rows[i].Cells[j].Value;
                        }
                    }
                }
                //保存文件
                workbook.SaveAs(fileDialog.FileName);
                application.Quit();
                MessageBox.Show("导出成功");
            }
        }
        #region 空
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void groupBox13_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {

        }   

        private void button17_Click_1(object sender, EventArgs e)
        {

        }

        private void button75_Click(object sender, EventArgs e)
        {

        }

        private void SetIDtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void SetInIDtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void GetIDtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void GetInIDtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
        #endregion

        private void button78_Click(object sender, EventArgs e)
        {
            comboBox24.Items.Clear();
            //获取电脑当前可用串口并添加到选项列表中
            comboBox24.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());

            comboBox1.Items.Clear();
            //获取电脑当前可用串口并添加到选项列表中
            comboBox1.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());

            comboBox30.Items.Clear();
            //获取电脑当前可用串口并添加到选项列表中
            comboBox30.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
        }

        private void groupBox16_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox15_Enter(object sender, EventArgs e)
        {

        }

        private void SetIDAtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void button74_Click_2(object sender, EventArgs e)
        {
            #region 算法板ID
            byte i = 0;
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte textlen = 0;
            byte[] textbuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            textlen = (byte)ByteOperate.GetBytes(ref textbuf, SetIDtext.Text);
            DLT64507.rever_char(ref textbuf, textlen);
            for (i = 0; i < textlen; i++)
            {
                databuf[datalen++] = textbuf[i];
            }
            textlen = (byte)ByteOperate.GetBytes(ref textbuf, SetInIDtext.Text);
            DLT64507.rever_char(ref textbuf, textlen);
            for (i = 0; i < textlen; i++)
            {
                databuf[datalen++] = textbuf[i];
            }
            DZFJDLL.String_Encrypt(databuf, datalen);
            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020005, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void button17_Click_2(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020006, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void button76_Click_1(object sender, EventArgs e)
        {
            #region 互感器ID
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte dataCnt = 0;
            byte[] datatmp = new byte[256];
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetIDAtext.Text);    //A相第一组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetInIDAtext.Text);  //A相第二组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetbarIDAtext.Text);  //A相第三组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }

            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetIDBtext.Text);    //B相第一组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetInIDBtext.Text);  //B相第二组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetbarIDBtext.Text);  //B相第三组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }

            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetIDCtext.Text);    //C相第一组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetInIDCtext.Text);  //C相第二组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }
            dataCnt = (byte)ByteOperate.GetBytes(ref datatmp, SetbarIDCtext.Text);  //C相第二组
            DLT64507.rever_char(ref datatmp, dataCnt);
            for (i = 0; i < dataCnt; i++)
            {
                databuf[datalen++] = datatmp[i];
            }

            DZFJDLL.String_Encrypt(databuf, datalen);
            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020003, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void button75_Click_1(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020004, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void button79_Click(object sender, EventArgs e)
        {
            #region 库参数配置
            ushort Tx_len = 0;
            int Value = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte dataCnt = 0;
            byte[] datatmp = new byte[256];
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Value = Convert.ToInt32(A_CUR25.Text);                                 //将文本框的文本转化成int型
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_CUR650.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_CUR2000.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_Amp_5_D.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_Amp_5_U.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_Amp_10_D.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_Amp_10_U.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_Amp_15_D.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_Amp_15_U.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_KL_NUM.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_DL_NUM.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_FRELT_DOWN.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_GPCUR_ITHD.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(A_KL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(A_DL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(A_KL_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_KL_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_ZC_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_ZC_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(AIRON_CORE1_LS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(AIRON_CORE1_RS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(AIRON_CORE1_LS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(AIRON_CORE1_RS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(AIRON_CORE1_LS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(AIRON_CORE1_RS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_CUR25.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_CUR650.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_CUR2000.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_Amp_5_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_Amp_5_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_Amp_10_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_Amp_10_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_Amp_15_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_Amp_15_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_KL_NUM.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_DL_NUM.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_FRELT_DOWN.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_GPCUR_ITHD.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(B_KL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(B_DL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(B_KL_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_KL_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_ZC_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_ZC_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(BIRON_CORE1_LS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(BIRON_CORE1_RS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(BIRON_CORE1_LS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(BIRON_CORE1_RS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(BIRON_CORE1_LS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(BIRON_CORE1_RS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);


            Value = Convert.ToInt32(C_CUR25.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_CUR650.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_CUR2000.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_Amp_5_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_Amp_5_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_Amp_10_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_Amp_10_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_Amp_15_D.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_Amp_15_U.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_KL_NUM.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_DL_NUM.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_FRELT_DOWN.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_GPCUR_ITHD.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(C_KL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(C_DL_SCANFRE.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(C_KL_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_KL_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_ZC_GAIN_K1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_ZC_GAIN_K2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(CIRON_CORE1_LS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(CIRON_CORE1_RS1.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(CIRON_CORE1_LS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(CIRON_CORE1_RS2.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(CIRON_CORE1_LS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(CIRON_CORE1_RS3.Text);    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            DZFJDLL.String_Encrypt(databuf, datalen);
            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020001, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void button85_Click(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020002, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void button86_Click(object sender, EventArgs e)
        {
            #region 精度校准参数
            ushort Tx_len = 0;
            int Value = 0;
            byte[] Tx_buf = new byte[256];
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };
            double xy=0;
            Value = Convert.ToInt32(A_VPP.Text);                                 //将文本框的文本转化成int型
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(A_KL_LIMIT.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = int.Parse(A_POWER_L_K.Text);// Convert.ToDouble(A_POWER_L_K.Text); 
            databuf[datalen++] = (byte)(Value);
            databuf[datalen++] = (byte)(Value >> 8);

            Value = int.Parse(A_POWER_L_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
            if (Value < 0)
            {
                int nk = System.Math.Abs(Value);
                databuf[datalen++] = (byte)nk;
                databuf[datalen++] = (byte)((nk >> 8) | 0x80);
            }
            else
            {
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);
            }

            Value = int.Parse(A_POWER_H_K.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = int.Parse(A_POWER_H_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
            if (Value < 0)
            {
                int nk = System.Math.Abs(Value);
                databuf[datalen++] = (byte)nk;
                databuf[datalen++] = (byte)((nk >> 8) | 0x80);
            }
            else
            {
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);
            }
            
            // else
            //{
            //    Value = Convert.ToInt32(A_POWER_L_K.Text) ;     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(A_POWER_L_S.Text);     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(A_POWER_H_K.Text);     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(A_POWER_H_S.Text); 
            //    //Value = ByteOperate.GetInt32(datatmp, true);
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);
            //}

            Value = Convert.ToInt32(A_HIGHT_L_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_HIGHT_L_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(A_HIGHT_H_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_HIGHT_H_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(A_DL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(A_KL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_VPP.Text);                                 //将文本框的文本转化成int型
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(B_KL_LIMIT.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);
           
                Value = int.Parse(B_POWER_L_K.Text); //xy = Convert.ToDouble(B_POWER_L_K.Text);
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);

                Value = int.Parse(B_POWER_L_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
                if (Value < 0)
                {
                    int nk = System.Math.Abs(Value);
                    databuf[datalen++] = (byte)nk;
                    databuf[datalen++] = (byte)((nk >> 8) | 0x80);
                }
                else
                {
                    databuf[datalen++] = (byte)Value;
                    databuf[datalen++] = (byte)(Value >> 8);
                }

                Value = int.Parse(B_POWER_H_K.Text);
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);

                Value = int.Parse(B_POWER_H_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
                if (Value < 0)
                {
                    int nk = System.Math.Abs(Value);
                    databuf[datalen++] = (byte)nk;
                    databuf[datalen++] = (byte)((nk >> 8) | 0x80);
                }
                else
                {
                    databuf[datalen++] = (byte)Value;
                    databuf[datalen++] = (byte)(Value >> 8);
                }
              
            
            //else
            //{
            //    Value = Convert.ToInt32(B_POWER_L_K.Text) ;     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(B_POWER_L_S.Text);     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(B_POWER_H_K.Text);     //A相第一组
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);

            //    Value = Convert.ToInt32(B_POWER_H_S.Text); 
            //    //Value = ByteOperate.GetInt32(datatmp, true);
            //    databuf[datalen++] = (byte)Value;
            //    databuf[datalen++] = (byte)(Value >> 8);
            //}

            Value = Convert.ToInt32(B_HIGHT_L_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_HIGHT_L_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(B_HIGHT_H_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_HIGHT_H_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(B_DL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(B_KL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_VPP.Text);                                 //将文本框的文本转化成int型
            databuf[datalen++] = (byte)Value;

            Value = Convert.ToInt32(C_KL_LIMIT.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

          
            Value = int.Parse(C_POWER_L_K.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = int.Parse(C_POWER_L_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
            if (Value < 0)
            {
                int nk = System.Math.Abs(Value);
                databuf[datalen++] = (byte)nk;
                databuf[datalen++] = (byte)((nk >> 8) | 0x80);
            }
            else
            {
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);
            }

            Value = int.Parse(C_POWER_H_K.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = int.Parse(C_POWER_H_S.Text); //xy = Convert.ToDouble(A_POWER_H_K.Text);Value = (int)xy;    //A相第一组
            if (Value < 0)
            {
                int nk = System.Math.Abs(Value);
                databuf[datalen++] = (byte)nk;
                databuf[datalen++] = (byte)((nk >> 8) | 0x80);
            }
            else
            {
                databuf[datalen++] = (byte)Value;
                databuf[datalen++] = (byte)(Value >> 8);
            }
           
            //else
            //{
            //Value = Convert.ToInt32(C_POWER_L_K.Text);     //A相第一组
            //databuf[datalen++] = (byte)Value;
            //databuf[datalen++] = (byte)(Value >> 8);

            //Value = Convert.ToInt32(C_POWER_L_S.Text);     //A相第一组
            //databuf[datalen++] = (byte)Value;
            //databuf[datalen++] = (byte)(Value >> 8);

            //Value = Convert.ToInt32(C_POWER_H_K.Text);     //A相第一组
            //databuf[datalen++] = (byte)Value;
            //databuf[datalen++] = (byte)(Value >> 8);

            //Value = Convert.ToInt32(C_POWER_H_S.Text);
            ////Value = ByteOperate.GetInt32(datatmp, true);
            //databuf[datalen++] = (byte)Value;
            //databuf[datalen++] = (byte)(Value >> 8);
            //}

            Value = Convert.ToInt32(C_HIGHT_L_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_HIGHT_L_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            Value = Convert.ToInt32(C_HIGHT_H_K.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_HIGHT_H_S.Text);    //A相第一组
            //Value = ByteOperate.GetInt32(datatmp, true);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);

            Value = Convert.ToInt32(C_DL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            Value = Convert.ToInt32(C_KL_value.Text);
            databuf[datalen++] = (byte)Value;
            databuf[datalen++] = (byte)(Value >> 8);
            databuf[datalen++] = (byte)(Value >> 16);

            DZFJDLL.String_Encrypt(databuf, datalen);
            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020007, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void button87_Click(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020008, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void button81_Click(object sender, EventArgs e)
        {
            #region 设置短路调教
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x07020009, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void button80_Click(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x0702000a, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void S_Hight_Base_Click(object sender, EventArgs e)
        {
            #region 设置开路调教
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x0702000b, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
            #endregion
        }

        private void R_Hight_Base_Click(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x0702000c, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {

        }

    

        private void Reading_Click(object sender, EventArgs e)
        {
            ushort Tx_len = 0;
            byte[] Tx_buf = new byte[256];
            byte i = 0;
            byte datalen = 0;
            byte[] databuf = new byte[256];
            byte[] CT_addr = new byte[6] { 0x01, 0x00, 0x00, 0x00, 0x00, 0x00 };

            Tx_len = DLT64507.DLT645_write(0, 0x03, CT_addr, 0x0702000f, databuf, datalen, Tx_buf);
            if (serialPort3.IsOpen)//判断串口是否打开，如果打开执行下一步操作
            {
                //设置文字显示
                foreach (byte Member in Tx_buf)
                {
                    string str = Convert.ToString(Member, 16).ToUpper();
                    textBox38.AppendText((str.Length == 1 ? "0" + str : str) + " ");
                }
                textBox38.AppendText("\r\n");
                serialPort3.Write(Tx_buf, 0, Tx_len);
            }
        }

        private void Crake1_TextChanged(object sender, EventArgs e)
        {

        }
        private string[] shiduArray = new string[31];

        private void KEEP_Click(object sender, EventArgs e)
        {
            int xduf = 0,ybuf = 0;
            dataGridView2.GridColor = Color.Blue;//设置网格颜色

            int index = this.dataGridView2.Rows.Add();
            this.dataGridView2.Rows[index].Cells[xduf].Value = current_time.ToString("yyyy-MM--dd HH:mm:ss");
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];

            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];

            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];

            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];

            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];

            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
            this.dataGridView2.Rows[index].Cells[++xduf].Value = shiduArray[++ybuf];
        }

        private void button93_Click(object sender, EventArgs e)
        {
             SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Excel文件|*.xlsx|Excel(2003文件)|*.xls";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = fileDialog.FileName;
                Excel.Application application = new Excel.Application();
                Excel.Workbooks workbooks = application.Workbooks;
                Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = workbook.Worksheets[1];
                int colIndex = 0;
                worksheet.Rows[1].RowHeight = 20; //第一行行高为60（单位：磅）
                for (int i = 1; i < 32; i++)
                {
                    worksheet.Columns[i].ColumnWidth = 25;
                }
                //导出DataGridView中的标题
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    if (dataGridView2.Columns[i].Visible)//做同于不导出隐藏列
                    {
                        colIndex++;
                        worksheet.Cells[1, colIndex] = dataGridView2.Columns[i].HeaderText;
                    }
                }
                //导出DataGridView中的数据
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    colIndex = 0;
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        if (dataGridView2.Columns[j].Visible)
                        {
                            colIndex++;
                            worksheet.Cells[i + 2, colIndex] = "'" + dataGridView2.Rows[i].Cells[j].Value;
                        }
                    }
                }
                //保存文件
                workbook.SaveAs(fileDialog.FileName);
                application.Quit();
                MessageBox.Show("导出成功");
            }
        }

        private void groupBox26_Enter(object sender, EventArgs e)
        {

        }

        private void textBox44_TextChanged(object sender, EventArgs e)
        {

        }

        private void button84_Click(object sender, EventArgs e)
        {
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textwans.Text = this.saveFileDialog1.FileName;
                address3 = this.saveFileDialog1.FileName;
            }
        }

        private void button83_Click(object sender, EventArgs e)
        {
            #region 合并Bin文件
            long i;
            string ad;
            BinaryWriter bw;
            ad = textoffset.Text;
            long address4 = long.Parse(ad, System.Globalization.NumberStyles.HexNumber);    //获取偏移地址
            for (i = 0; i < address4; i++)
            {
                if (i > SIZE2)
                    write[i] = 0;
                else
                    write[i] = read2[i];
            }
            for (i = 0; i < (SIZE1 + Version_num + num); i++)
            {
                write[(address4 + i)] = read1[i];
            }

            // 保存文件
            try
            {
                bw = new BinaryWriter(new FileStream(address3, FileMode.Create));

            }
            catch (IOException EX)
            {
                textanswer.Text = "新建BIN文件失败";
                return;
            }
            try
            {
                for (i = 0; i < (address4 + SIZE1 + num + Version_num); i++)
                {
                    bw.Write(write[i]);
                }
                textanswer.Text = "BIN文件合并成功";
            }
            catch (IOException EX)
            {
                textanswer.Text = "合并BIN文件失败";
                return;
            }

            bw.Close();
            #endregion
        }

        private void BOOT_Click(object sender, EventArgs e)
        {
            BinaryReader br;
            int i;
            int result = 0;
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)    //获取文件地址
            {
                textBOOT.Text = this.openFileDialog1.FileName;
                address1 = this.openFileDialog1.FileName;

                result = 1;
            }
            else
                textanswer.Text = "未添加BOOT文件";
            if (result == 1)
            {
                try
                {
                    FileStream fileStream = new FileStream(address1, FileMode.Open);

                    FileInfo mydir = new FileInfo(address1);    //获取文件大小
                    SIZE2 = mydir.Length;
                    br = new BinaryReader(fileStream);
                }
                catch (IOException EX)
                {
                    textanswer.Text = ("获取BOOT文件失败");
                    return;
                }
                try
                {
                    for (i = 0; i < SIZE2; i++)
                    {
                        read2[i] = br.ReadByte();
                    }
                }
                catch (IOException EX)
                {
                    textanswer.Text = ("BOOT文件读取失败");
                    return;
                }

                br.Close();
            }
            else return;

        }

        private void APP_Click(object sender, EventArgs e)
        {
            BinaryReader br;
            int i;
            int result = 0;
            if (this.openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textAPP.Text = this.openFileDialog2.FileName;
                address2 = this.openFileDialog2.FileName;
                result = 1;
            }
            else
                textanswer.Text = "未添加APP文件";

            if (result == 1)
            {
                try
                {
                    FileStream fileStream = new FileStream(address2, FileMode.Open);
                    br = new BinaryReader(fileStream);
                    FileInfo mydir = new FileInfo(address2);    //获取文件大小
                    SIZE1 = mydir.Length;
                    br = new BinaryReader(fileStream);
                }
                catch (IOException EX)
                {
                    textanswer.Text = ("获取APP文件失败");
                    return;
                }
                try
                {
                    for (i = 0; i < SIZE1; i++)
                    {
                        read1[i] = br.ReadByte();
                    }

                }
                catch (IOException EX)
                {
                    textanswer.Text = ("APP文件读取失败");
                    return;
                }

                br.Close();
                //  textanswer.Text = Convert.ToString(read1[SIZE1 - 4]);
            }
            else return;
        }

        private void tailing_Click(object sender, EventArgs e)
        {
            int i;
            UInt32 crc = 0x00000000;
            UInt32 j;
            byte datalen = 0;
            byte[] tail = new byte[256];
            BinaryWriter bw;
            //判断是否进行过加尾
            if ((read1[SIZE1 - 5] == 0x00) && (read1[SIZE1 - 6] == 0x00) && (read1[SIZE1 - 7] == 0x00) && (read1[SIZE1 - 8] == 0x00) && (read1[SIZE1 - 9] == 0x00) && ((read1[SIZE1 - 10] == 0x30) || (read1[SIZE1 - 10] == 0x31)))
            {
                Version_num = 0;
                textanswer.Text = "BIN文件已进行加尾处理";
                return;
            }
            else
            {

                for (i = 0; i < texttailing.Text.Length; i++)
                {
                    tail[i] = (byte)texttailing.Text[i];
                }
                if (texttailing.Text.Length > 12)
                {
                    textanswer.Text = ("数据标识越界");
                    return;
                }

                if (SIZE1 % 4 != 0)             //补零，形成4的倍数
                {

                    num = 4 - (SIZE1 % 4);
                    for (i = 0; i < num; i++)
                        read1[SIZE1 + i] = 0;
                }
                for (i = 0; i < 12; i++)                  //写入型号
                {
                    read1[SIZE1 + num + i] = tail[i];

                }

                crc = crc ^ 0xffffffff;                         //CRC校验
                for (j = 0; j < (SIZE1 + 12 + num); j++)
                {
                    crc = (crc >> 8) ^ crc32_table[(crc & 0xff) ^ read1[j]];
                }
                crc = crc ^ 0xffffffff;
                CR[0] = (byte)((crc >> 24) & 0xff);
                CR[1] = (byte)((crc >> 16) & 0xff);
                CR[2] = (byte)((crc >> 8) & 0xff);
                CR[3] = (byte)(crc & 0xff);
                for (i = 0; i < 4; i++)
                {
                    read1[SIZE1 + 12 + num + i] = CR[i];
                }
                Version_num = 16;

                //写文件
                try
                {
                    bw = new BinaryWriter(new FileStream(address3, FileMode.Create));
                }
                catch (IOException EX)
                {
                    textanswer.Text = "生成BIN文件失败";
                    return;
                }
                try
                {
                    for (i = 0; i < (SIZE1 + num + Version_num); i++)
                    {
                        bw.Write(read1[i]);
                    }
                    textanswer.Text = "BIN文件加尾成功";
                }
                catch (IOException EX)
                {
                    textanswer.Text = "BIN文件加尾失败";
                    return;
                }

                bw.Close();
            }
        }

        private void buttonfile_Click(object sender, EventArgs e)
        {
            BinaryReader br;
            int i;
            int result = 0;
            string site;
            if (this.openFileDialog3.ShowDialog() == DialogResult.OK)    //获取文件地址
            {
                textfile.Text = this.openFileDialog1.FileName;
                site = this.openFileDialog1.FileName;

                result = 1;
            }
            else return;

            if (result == 1)
            {
                try
                {
                    FileStream fileStream = new FileStream(site, FileMode.Open);

                    FileInfo mydir = new FileInfo(site);    //获取文件大小
                    SIZE2 = mydir.Length;
                    br = new BinaryReader(fileStream);
                }
                catch (IOException EX)
                {
                    return;
                }
                try
                {
                    for (i = 0; i < SIZE2; i++)
                    {
                        read3[i] = br.ReadByte();
                    }
                }
                catch (IOException EX)
                {
                    return;
                }

                br.Close();
            }
            else return;
        }

        private void Upgrade_Click(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                //自动发送功能选中,开始自动发送
                numericUpDown2.Enabled = false;     //失能时间选择
                timer3.Interval = (int)numericUpDown2.Value;     //定时器赋初值
                timer3.Start();     //启动定时器
            }
            else
            {
                //自动发送功能未选中,停止自动发送
                numericUpDown2.Enabled = true;     //使能时间选择
                timer3.Stop();     //停止定时器
            }
        }

        #region 空
        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void tabPage11_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            Reading_Click(Reading, new EventArgs());
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label101_Click(object sender, EventArgs e)
        {

        }

        private void SetIDBtext_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox16_Enter_1(object sender, EventArgs e)
        {

        }

        private void label205_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox37_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox44_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox45_TextChanged(object sender, EventArgs e)
        {

        }

        private void AIRON_CORE1_RS1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox40_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox39_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void groupBox14_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        #endregion
        #region 阻抗相配置
        private void button59_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x0C, 0x03, 0x00C, 0x03, 0x0C, 0x9D, 0x16 };
            Key_Sentences(Data);
            Control_Color1(10);
        }

        private void button61_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x0B, 0x03, 0x0B, 0x03, 0x0B, 0x9A, 0x16 };
            Key_Sentences(Data);
            Control_Color1(11);
        }

        private void button63_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x08, 0x03, 0x08, 0x03, 0x08, 0x91, 0x16 };
            Key_Sentences(Data);
            Control_Color1(20);
        }

        private void button64_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x09, 0x03, 0x09, 0x03, 0x09, 0x94, 0x16 };
            Key_Sentences(Data);
            Control_Color1(21);
        }

        private void button65_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x00, 0x02, 0x06, 0x03, 0x0A, 0x03, 0x0A, 0x03, 0x0A, 0x97, 0x16 };
            Key_Sentences(Data);
            Control_Color1(22);
        } 

        private void button66_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x00C, 0x03, 0x0C, 0x03, 0x0C, 0x9E, 0x16 };
            Key_Sentences(Data);
            Control_Color2(10);
        }

        private void button83_Click_1(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x0B, 0x03, 0x0B, 0x03, 0x0B, 0x9B, 0x16 };
            Key_Sentences(Data);
            Control_Color2(11);
        }

        private void button84_Click_1(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x08, 0x03, 0x08, 0x03, 0x08, 0x92, 0x16 };
            Key_Sentences(Data);
            Control_Color2(20);
        }

        private void button88_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x09, 0x03, 0x09, 0x03, 0x09, 0x95, 0x16 };
            Key_Sentences(Data);
            Control_Color2(21);
        }

        private void button89_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x01, 0x02, 0x06, 0x03, 0x0A, 0x03, 0x0A, 0x03, 0x0A, 0x98, 0x16 };
            Key_Sentences(Data);
            Control_Color2(22);
        }

        private void button90_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x0C, 0x03, 0xBC, 0x03, 0x0C, 0x9F, 0x16 };
            Key_Sentences(Data);
            Control_Color3(10);
        }

        private void button91_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x0B, 0x03, 0x0B, 0x03, 0x0B, 0x9C, 0x16 };
            Key_Sentences(Data);
            Control_Color3(11);
        }

        private void button92_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x08, 0x03, 0x08, 0x03, 0x08, 0x93, 0x16 };
            Key_Sentences(Data);
            Control_Color3(20);
        }

        private void button93_Click_1(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x09, 0x03, 0x09, 0x03, 0x09, 0x96, 0x16 };
            Key_Sentences(Data);
            Control_Color3(21);
        }

        private void button94_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[12] { 0x68, 0x02, 0x02, 0x06, 0x03, 0x0A, 0x03, 0x0A, 0x03, 0x0A, 0x99, 0x16 };
            Key_Sentences(Data);
            Control_Color3(22);
        }
        #endregion

        int standard_position = 0, standard_stream=0,standard_sum=0;

        private void button82_Click(object sender, EventArgs e)
        {
            label282.Text = "等待中...";
            textBox41.Text ="";
            textBox46.Text = "";
            label26.BackColor = Color.White; label27.BackColor = Color.White; label28.BackColor = Color.White; label29.BackColor = Color.White; label30.BackColor = Color.White; label31.BackColor = Color.White;
            label245.BackColor = Color.White; label246.BackColor = Color.White; label247.BackColor = Color.White; label248.BackColor = Color.White; label249.BackColor = Color.White; label250.BackColor = Color.White;
            textBox41.AppendText("0");
            textBox46.AppendText("0");
            button60_Click(button60, new EventArgs());
            Delay(500);
            button6_Click(button6, new EventArgs());//A相二次开路
            Delay(500);
            button29_Click(button29, new EventArgs());//B相二次开路
            Delay(500);
            button46_Click(button46, new EventArgs());//C相二次开路
            Delay(500);
           
            button81_Click(button81, new EventArgs());               //设置短路调教
            Delay(800);
            S_Hight_Base_Click(S_Hight_Base, new EventArgs());      //设置开路调教
            Delay(800);
            //button80_Click(button80, new EventArgs());//读取工频基值
            //Delay(1000); 
            //R_Hight_Base_Click(R_Hight_Base, new EventArgs());//读取高频基值
            //Delay(1000);
            standard_position = 1;
            timer4.Interval = 60000;
            timer4.Start();     //启动定时器
       }

        #region 空
        private void label59_Click(object sender, EventArgs e)
       {

       }

       private void label244_Click(object sender, EventArgs e)
       {

       }

       private void label249_Click(object sender, EventArgs e)
       {

       }

       private void label27_Click(object sender, EventArgs e)
       {

       }

       private void label28_Click(object sender, EventArgs e)
       {

       }

       private void label30_Click(object sender, EventArgs e)
       {

       }

       private void label250_Click(object sender, EventArgs e)
       {

       }

       private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
       {

       }

       private void tabPage14_Click(object sender, EventArgs e)
       {

       }

       private void textBox17_TextChanged_1(object sender, EventArgs e)
       {

       }

       private void textBox50_TextChanged_1(object sender, EventArgs e)
       {

       }

       private void textBox49_TextChanged(object sender, EventArgs e)
       {

       }
        #endregion

      // double[] x_xy = new double[3] { 0.1, 0.3, 0.6 };
       double[] current1 = new double[3];
       double[] value1 = new double[3];
       double[] current2 = new double[3];
       double[] value2 = new double[3];
       double[] current3 = new double[3];
       double[] value3 = new double[3];
       double[] position1 = { 0.1, 0.3, 0.6 };
       double[] position2 = { 2, 4, 5 };

       private double Sum_Average(double[] d)
      {
        double z = 0;
	    for(int i=0;i<3;i++)
	   {
		  z = z + d[i];
	   }
	      z = z/3;
	    return z;
      }

       private double X_Y_By(double[] m, double[] n)
      {
        double z = 0;
        for (int i = 0; i < 3; i++)
	   {
	 	z = z + m[i]*n[i];
	   }
	   return z;
      }

       private double Squre_sum(double[] c)
     {
        double z = 0;
	    for(int i=0;i<3;i++)
	  {
	 	z = z + c[i]*c[i];
	  }
	    return z;
     }

         double Q = 0;				//拟合直线的斜率
         double R = 0;				//拟合直线的截距
         int K = 0;
         int B = 0;
         double x_sum_average = 0;   	//数组 X[N] 个元素求和 并求平均值
         double y_sum_average = 0;   	//数组 Y[N] 个元素求和 并求平均值
         double x_square_sum = 0;   	//数组 X[N] 个个元素的平均值
         double x_multiply_y = 0;   	//数组 X[N]和Y[N]对应元素的乘机
       private void Line_Fit(double[] X, double[] Y)
     {
	    x_sum_average= Sum_Average(X);
	    y_sum_average= Sum_Average(Y);
	    x_square_sum = Squre_sum(X);
	    x_multiply_y = X_Y_By(X,Y);

	    Q = (( x_multiply_y - 3 * x_sum_average * y_sum_average)/( x_square_sum - 3 * x_sum_average*x_sum_average ));
        R = (y_sum_average - Q * x_sum_average);
        K = (int)(Q * 10000);
        B = (int)(R * 10000);
        if (Y == value1) { A_POWER_H_K.AppendText(Convert.ToString(K)); A_POWER_H_S.AppendText(Convert.ToString(B));}
        else if (Y == current1)
        {
            A_POWER_L_K.Text = ""; A_POWER_L_S.Text = ""; A_POWER_H_K.Text = ""; A_POWER_H_S.Text = "";
            B_POWER_L_K.Text = ""; B_POWER_L_S.Text = ""; B_POWER_H_K.Text = ""; B_POWER_H_S.Text = "";
            C_POWER_L_K.Text = ""; C_POWER_L_S.Text = ""; C_POWER_H_K.Text = ""; C_POWER_H_S.Text = "";
            Delay(200);
            A_POWER_L_K.Text = Convert.ToString(K); A_POWER_L_S.Text = Convert.ToString(B);
        }
        else if (Y == value2) { B_POWER_H_K.Text = (Convert.ToString(K)); B_POWER_H_S.Text = (Convert.ToString(B)); }
        else if (Y == current2) { B_POWER_L_K.Text = (Convert.ToString(K)); B_POWER_L_S.Text = (Convert.ToString(B)); }
        else if (Y == value3) 
        { 
            C_POWER_H_K.Text = (Convert.ToString(K)); 
            C_POWER_H_S.Text = (Convert.ToString(B));
            Delay(200);
            button86_Click(button86, new EventArgs());
            Delay(2000);
        }
        else if (Y == current3) { C_POWER_L_K.Text = (Convert.ToString(K)); C_POWER_L_S.Text = (Convert.ToString(B)); }
     }

       Int32 sum1 = 0, sum2 = 0, sum3 = 0, SUM1 = 0, SUM2 = 0, SUM3 = 0;
       Int32 num1 = 0, num2 = 0, num3 = 0, NUM1 = 0, NUM2 = 0, NUM3 = 0;
       int dengdai = 0;
       private void timer4_Tick(object sender, EventArgs e)
       {
           switch (standard_position)
           {
               case 1://校准工频偏置
                   {
                       timer4.Stop();     //关闭定时器
                       button80_Click(button80, new EventArgs());//读取工频基值
                       Delay(2000);
                       R_Hight_Base_Click(R_Hight_Base, new EventArgs());//读取高频基值
                       Delay(2000);
                       standard_position = 0;
                       label282.Text = "";
                   } break;
               case 2://60K
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs()); 
                       Delay(2000);

                       A_Amp_5_D.Text = (Convert.ToString(Convert.ToInt32(AKDOWN05.Text)*10*1.5));//注入5K 10K 15K
                       A_Amp_10_D.Text = (Convert.ToString(Convert.ToInt32(AKDOWN10.Text)*10*1.5));
                       A_Amp_15_D.Text = (Convert.ToString(Convert.ToInt32(AKDOWN15.Text)*10*1.5));

                       B_Amp_5_D.Text = (Convert.ToString(Convert.ToInt32(BKDOWN05.Text)*10*1.5));
                       B_Amp_10_D.Text = (Convert.ToString(Convert.ToInt32(BKDOWN10.Text)*10*1.5));
                       B_Amp_15_D.Text = (Convert.ToString(Convert.ToInt32(BKDOWN15.Text)*10*1.5));

                       C_Amp_5_D.Text = (Convert.ToString(Convert.ToInt32(CKDOWN05.Text)*10*1.5));
                       C_Amp_10_D.Text = (Convert.ToString(Convert.ToInt32(CKDOWN10.Text)*10*1.5));
                       C_Amp_15_D.Text = (Convert.ToString(Convert.ToInt32(CKDOWN15.Text)*10*1.5));

                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       standard_position = 0;
                       AKDOWN05.Text = A_Amp_5_D.Text;
                       BKDOWN05.Text = B_Amp_5_D.Text;
                       CKDOWN05.Text = C_Amp_5_D.Text;
                       AKDOWN10.Text = A_Amp_10_D.Text;
                       BKDOWN10.Text = B_Amp_10_D.Text;
                       CKDOWN10.Text = C_Amp_10_D.Text;
                       AKDOWN15.Text = A_Amp_15_D.Text;
                       BKDOWN15.Text = B_Amp_15_D.Text;
                       CKDOWN15.Text = C_Amp_15_D.Text;
                       label282.Text = "";
                   } break;
               case 3://标准工频电流
                   { 
                       timer4.Stop();     //关闭定时器
                       button71_Click(button71, new EventArgs());//读取参数
                       Delay(2000);
                       standard_position = 0;
                       standard_stream = standard_stream + 1;//更改电流值
                       if (standard_stream < 6)
                       {
                           button95_Click(button95, new EventArgs());//重新扫频
                           Delay(2000);
                       }
                       else if (standard_stream == 6 && standard_sum == 0)//赋值
                       {
                           if (textBox75.Text != "" && textBox76.Text != "" && textBox77.Text != "" && textBox78.Text != "" && textBox79.Text != "" && textBox80.Text != "" &&
                               textBox86.Text != "" && textBox85.Text != "" && textBox84.Text != "" && textBox83.Text != "" && textBox82.Text != "" && textBox81.Text != "" &&
                               textBox92.Text != "" && textBox91.Text != "" && textBox90.Text != "" && textBox101.Text != "" && textBox88.Text != "" && textBox87.Text != "")
                           {
                           current1[0] = (double)Convert.ToDouble(textBox75.Text);
                           current1[1] = (double)Convert.ToDouble(textBox76.Text);
                           current1[2] = (double)Convert.ToDouble(textBox77.Text);
                           value1[0] = (double)Convert.ToDouble(textBox78.Text);
                           value1[1] = (double)Convert.ToDouble(textBox79.Text);
                           value1[2] = (double)Convert.ToDouble(textBox80.Text);
                           current2[0] = (double)Convert.ToDouble(textBox86.Text);
                           current2[1] = (double)Convert.ToDouble(textBox85.Text);
                           current2[2] = (double)Convert.ToDouble(textBox84.Text);
                           value2[0] = (double)Convert.ToDouble(textBox83.Text);
                           value2[1] = (double)Convert.ToDouble(textBox82.Text);
                           value2[2] = (double)Convert.ToDouble(textBox81.Text);
                           current3[0] = (double)Convert.ToDouble(textBox92.Text);
                           current3[1] = (double)Convert.ToDouble(textBox91.Text);
                           current3[2] = (double)Convert.ToDouble(textBox90.Text);
                           value3[0] = (double)Convert.ToDouble(textBox101.Text);
                           value3[1] = (double)Convert.ToDouble(textBox88.Text);
                           value3[2] = (double)Convert.ToDouble(textBox87.Text);
                           Delay(500);
                           
                           Line_Fit(position1, current1);//耦合算法
                           Line_Fit(position2, value1);
                           Line_Fit(position1, current2);
                           Line_Fit(position2, value2);
                           Line_Fit(position1, current3);
                           Line_Fit(position2, value3);
                           Delay(1000);
                           
                           textBox100.Text = A_POWER_L_K.Text; textBox99.Text = A_POWER_L_S.Text; textBox107.Text = A_POWER_H_K.Text; textBox106.Text = A_POWER_H_S.Text;
                           textBox98.Text  = B_POWER_L_K.Text; textBox97.Text = B_POWER_L_S.Text; textBox105.Text = B_POWER_H_K.Text; textBox104.Text = B_POWER_H_S.Text;
                           textBox96.Text  = C_POWER_L_K.Text; textBox95.Text = C_POWER_L_S.Text; textBox103.Text = C_POWER_H_K.Text; textBox102.Text = C_POWER_H_S.Text;
                           standard_stream = 0;
                           if (standard_sum == 0)
                           {
                               button95_Click(button95, new EventArgs());
                               Delay(2000);
                               standard_sum = 1;
                           }
                           }
                           else
                           {
                               button98_Click(button98,new EventArgs());
                               Delay(1000);
                           }
                       }
                       else if (standard_sum == 1)
                       {
                           standard_sum = 0;
                           standard_stream = 0;

                           int acm = 1;
                           double a = (0.1 - current1[0]) / 0.1;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm=1; }
                           else acm = 0;
                           a = (0.3 - current1[1]) / 0.3;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (0.6 - current1[2]) / 0.6;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;

                           a = (2 - value1[0]) / 2;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (4 - value1[1]) / 4;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (5 - value1[2]) / 5;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;

                           a = (0.1 - current2[0]) / 0.1;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (0.3 - current2[1]) / 0.3;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (0.6 - current2[2]) / 0.6;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;

                           a = (2 - value2[0]) / 2;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (4 - value2[1]) / 4;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (5 - value2[2]) / 5;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;

                           a = (0.1 - current3[0]) / 0.1;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (0.3 - current3[1]) / 0.3;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (0.6 - current3[2]) / 0.6;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;

                           a = (2 - value3[0]) / 2;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (4 - value3[1]) / 4;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           a = (5 - value3[2]) / 5;
                           if (a < 0.2 && a > -0.2 && acm == 1) { acm = 1; }
                           else acm = 0;
                           if(acm==1) label278.BackColor = Color.ForestGreen;
                           else if(acm==0)label278.BackColor = Color.Red;
                           label282.Text = "";
                       }
                   } break;
               case 4://扫频开路校准
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(2000);
                       A_KL_SCANFRE.Text = ""; B_KL_SCANFRE.Text = ""; C_KL_SCANFRE.Text = "";
                       A_KL_SCANFRE.Text = ATh_value.Text;
                       B_KL_SCANFRE.Text = BTh_value.Text;
                       C_KL_SCANFRE.Text = CTh_value.Text;
                       Delay(2000);
                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       textBox93.Text="扫频结束";
                       textBox62.Text = A_KL_SCANFRE.Text;
                       textBox61.Text = B_KL_SCANFRE.Text;
                       textBox60.Text = B_KL_SCANFRE.Text;
                       label282.Text = "";
                   } break;
               case 5://扫频短路校准
                   { 
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(2000);
                       A_DL_SCANFRE.Text = ""; B_DL_SCANFRE.Text = ""; C_DL_SCANFRE.Text = "";

                       Int32 s = Convert.ToInt32(ATh_value.Text);
                       s = s / 100;
                       A_DL_SCANFRE.AppendText(Convert.ToString(s));
                       s = Convert.ToInt32(BTh_value.Text);
                       s = s / 100;
                       B_DL_SCANFRE.AppendText(Convert.ToString(s));
                       s = Convert.ToInt32(CTh_value.Text);
                       s = s / 100; ;
                       C_DL_SCANFRE.AppendText(Convert.ToString(s));
                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       textBox94.Text="扫频结束";
                       textBox59.Text = A_DL_SCANFRE.Text;
                       textBox58.Text = B_DL_SCANFRE.Text;
                       textBox57.Text = C_DL_SCANFRE.Text;
                       label282.Text = "";
                   } break;
               case 6://短路
                     { 
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(2000);
                       Int32 h = Convert.ToInt32(AKUP05.Text);//设置5K 10K 15K
                       A_Amp_5_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(AKUP10.Text);
                       A_Amp_10_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(AKUP15.Text);
                       A_Amp_15_U.Text = (System.Convert.ToString(h, 16));

                       h = Convert.ToInt32(BKUP05.Text);
                       B_Amp_5_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(BKUP10.Text);
                       B_Amp_10_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(BKUP15.Text);
                       B_Amp_15_U.Text = (System.Convert.ToString(h,16));

                       h = Convert.ToInt32(CKUP05.Text);
                       C_Amp_5_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(CKUP10.Text);
                       C_Amp_10_U.Text = (System.Convert.ToString(h, 16));
                       h = Convert.ToInt32(CKUP15.Text);
                       C_Amp_15_U.Text = (System.Convert.ToString(h, 16));

                       h = (Int32)(Convert.ToInt32(AVpp_5K.Text) );//读取Ve
                       A_DL_value.Text = (Convert.ToString(h, 16));
                       h = (Int32)(Convert.ToInt32(BVpp_5K.Text) );
                       B_DL_value.Text = (Convert.ToString(h, 16));
                       h = (Int32)(Convert.ToInt32(CVpp_5K.Text) );
                       C_DL_value.Text = (Convert.ToString(h, 16));
                       //h = Convert.ToInt32(AVpp_5K.Text);
                       //A_DL_value.Text = (Convert.ToString(h,16));//读取Ve
                       //h = Convert.ToInt32(BVpp_5K.Text);
                       //B_DL_value.Text = (Convert.ToString(h,16));
                       //h = Convert.ToInt32(CVpp_5K.Text);
                       //C_DL_value.Text = (Convert.ToString(h,16));

                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       button86_Click(button86, new EventArgs());
                       Delay(2000);
                       standard_position = 0;
                       biaozhi = 0;
                       AKUP05.Text = A_Amp_5_U.Text;
                       AKUP10.Text = A_Amp_10_U.Text;
                       AKUP15.Text = A_Amp_15_U.Text;

                       BKUP05.Text = B_Amp_5_U.Text;
                       BKUP10.Text = B_Amp_10_U.Text;
                       BKUP15.Text = B_Amp_15_U.Text;

                       CKUP05.Text = C_Amp_5_U.Text;
                       CKUP10.Text = C_Amp_10_U.Text;
                       CKUP15.Text = C_Amp_15_U.Text;

                       textBox68.Text = A_DL_value.Text;
                       textBox67.Text = B_DL_value.Text;
                       textBox66.Text = C_DL_value.Text;
                       label282.Text = "";
                   } break;
               case 7://开路
                   {
                           timer4.Stop();     //关闭定时器
                           Reading_Click(Reading, new EventArgs());
                           Delay(3000);
                           A_FRELT_DOWN.Text = ""; B_FRELT_DOWN.Text = ""; C_FRELT_DOWN.Text = "";//频率
                           Int32 h = Convert.ToInt32(textBox71.Text) ;
                           A_FRELT_DOWN.Text = (Convert.ToString(h + 12288, 16));
                           h = Convert.ToInt32(textBox70.Text);
                           B_FRELT_DOWN.Text = (Convert.ToString(h + 12288, 16));
                           h = Convert.ToInt32(textBox69.Text);
                           C_FRELT_DOWN.Text = (Convert.ToString(h + 12288, 16));

                           button79_Click(button79, new EventArgs());
                           Delay(2000);
                           button86_Click(button86, new EventArgs());
                           Delay(2000);
                           standard_position = 0;
                           biaozhi = 0;

                           textBox71.Text = A_FRELT_DOWN.Text;
                           textBox70.Text = B_FRELT_DOWN.Text;
                           textBox69.Text = C_FRELT_DOWN.Text;
                           label282.Text = "";
                   }break;
               case 8:
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(2000);
                       standard_position = 0;
                       label282.Text = "";
                   }break;
               case 9://70K
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(3000);

                       sum1 = Convert.ToInt32(textBox53.Text);
                       sum2 = Convert.ToInt32(textBox52.Text);
                       sum3 = Convert.ToInt32(textBox51.Text);
                       num1 = Convert.ToInt32(textBox56.Text);
                       num2 = Convert.ToInt32(textBox55.Text);
                       num3 = Convert.ToInt32(textBox54.Text);

                       if (sum1 > num1) sum1 = sum1;
                       else sum1 = num1;

                       if (sum2 > num2) sum2 = sum2;
                       else sum2 = num2;

                       if (sum3 > num3) sum3 = sum3;
                       else sum3 = num3;
                      
                       //button79_Click(button79, new EventArgs());
                       //Delay(1000);
                       standard_position = 0;

                       button6_Click(button8, new EventArgs());
                       Delay(500);
                       button29_Click(button29, new EventArgs());
                       Delay(500);
                       button46_Click(button46, new EventArgs());
                       Delay(500);
                       dispel();
                       //  Reading_Click(Reading, new EventArgs());
                       standard_position = 10;
                       timer4.Interval = 80000;
                       timer4.Start();     //启动定时器
                       label282.Text = "";
                   }break;
               case 10://开路测斜率
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(3000);

                       SUM1 = Convert.ToInt32(textBox53.Text);
                       SUM2 = Convert.ToInt32(textBox52.Text);
                       SUM3 = Convert.ToInt32(textBox51.Text);
                       NUM1 = Convert.ToInt32(textBox56.Text);
                       NUM2 = Convert.ToInt32(textBox55.Text);
                       NUM3 = Convert.ToInt32(textBox54.Text);

                       if (SUM1 > NUM1) SUM1 = SUM1;
                       else SUM1 = NUM1;

                       if (SUM2 > NUM2) SUM2 = SUM2;
                       else SUM2 = NUM2;

                       if (SUM3 > NUM3) SUM3 = SUM3;
                       else SUM3 = NUM3;
                       Int32 K1 = (Int32)(sum1 * 0.7 + SUM1 * 0.3);
                       Int32 K2 = (Int32)(sum2 * 0.7 + SUM2 * 0.3);
                       Int32 K3 = (Int32)(sum3 * 0.7 + SUM3 * 0.3);
                       A_KL_GAIN_K1.Text = Convert.ToString(K1 / 100); A_KL_GAIN_K2.Text = Convert.ToString(K1 / 100);
                       B_KL_GAIN_K1.Text = Convert.ToString(K2 / 100); B_KL_GAIN_K2.Text = Convert.ToString(K2 / 100);
                       C_KL_GAIN_K1.Text = Convert.ToString(K3 / 100); C_KL_GAIN_K2.Text = Convert.ToString(K3 / 100);

                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       standard_position = 0;
                       textBox53.Text = A_KL_GAIN_K1.Text;
                       textBox52.Text = B_KL_GAIN_K1.Text;
                       textBox51.Text = C_KL_GAIN_K1.Text;
                       textBox56.Text = A_KL_GAIN_K1.Text;
                       textBox55.Text = B_KL_GAIN_K1.Text;
                       textBox54.Text = C_KL_GAIN_K1.Text;
                       label282.Text = "";
                   } break;
               case 11://50K
                   {
                       timer4.Stop();     //关闭定时器
                       Reading_Click(Reading, new EventArgs());
                       Delay(3000);

                       Int32 s = Convert.ToInt32(AVpp_5K.Text);
                       A_KL_value.Text = (Convert.ToString(s));//读取Ve
                       s = Convert.ToInt32(BVpp_5K.Text);
                       B_KL_value.Text = (Convert.ToString(s));
                       s = Convert.ToInt32(CVpp_5K.Text);
                       C_KL_value.Text = (Convert.ToString(s));
                       button79_Click(button79, new EventArgs());
                       Delay(2000);
                       button86_Click(button86, new EventArgs());
                       Delay(2000);
                       standard_position = 0;
                       biaozhi = 0;
                       textBox65.Text = A_KL_value.Text;
                       textBox64.Text = B_KL_value.Text;
                       textBox63.Text = C_KL_value.Text;
                       label282.Text = "";
                   } break; 
           }
               
       }

       private void groupBox4_Enter(object sender, EventArgs e)
       {

       }

       private void button95_Click(object sender, EventArgs e)
       {
           label282.Text = "等待中...";
           label278.BackColor = Color.White;
           if (standard_position == 0 && standard_stream == 0)//文本清空
          {
              textBox75.Text = ""; textBox76.Text = ""; textBox77.Text = ""; textBox78.Text = ""; textBox79.Text = ""; textBox80.Text = "";
              textBox81.Text = ""; textBox82.Text = ""; textBox83.Text = ""; textBox84.Text = ""; textBox85.Text = ""; textBox86.Text = "";
              textBox87.Text = ""; textBox88.Text = ""; textBox101.Text = ""; textBox90.Text = ""; textBox91.Text = ""; textBox92.Text = "";
          }
             button8_Click(button8, new EventArgs());//A相二次正常
             Delay(500);
             button31_Click(button31, new EventArgs());//B相二次正常
             Delay(500);
             button48_Click(button48, new EventArgs());//C相二次正常
             Delay(500);
             switch (standard_stream)//设电流参数
           {
               case 0: textBox41.Text = "0"; textBox46.Text = "0.1" ; break;
               case 1: textBox41.Text = "0"; textBox46.Text = "0.3"; break;
               case 2: textBox41.Text = "0"; textBox46.Text = "0.6"; break;
               case 3: textBox41.Text = "0"; textBox46.Text = "2"; break;
               case 4: textBox41.Text = "0"; textBox46.Text = "4"; break;
               case 5: textBox41.Text = "0"; textBox46.Text = "5"; break;
           }
                 Delay(200);
                 button60_Click(button60, new EventArgs());//开启电源
                 Delay(1000);
                 button60_Click(button60, new EventArgs());//开启电源
                 Delay(1000);
             standard_position = 3;
             timer4.Interval = 50000;
             timer4.Start();     //启动定时器
       }

       private void C_KDOWN15_TextChanged(object sender, EventArgs e)
       {

       }

       private void dispel()
       { 
           AKDOWN05.Text = "";  AKDOWN10.Text = "";  AKDOWN15.Text = "";
           BKDOWN05.Text = "";  BKDOWN10.Text = "";  BKDOWN15.Text = "";
           CKDOWN05.Text = "";  CKDOWN10.Text = "";  CKDOWN15.Text = "";
           AKUP05.Text = ""; AKUP10.Text = ""; AKUP15.Text = "";
           BKUP05.Text = ""; BKUP10.Text = ""; BKUP15.Text = "";
           CKUP05.Text = ""; CKUP10.Text = ""; CKUP15.Text = "";
           textBox51.Text = ""; textBox52.Text = ""; textBox53.Text = ""; textBox54.Text = ""; textBox55.Text = ""; textBox56.Text = "";
           textBox57.Text = ""; textBox58.Text = ""; textBox59.Text = ""; textBox60.Text = ""; textBox61.Text = ""; textBox62.Text = "";
           textBox63.Text = ""; textBox64.Text = ""; textBox65.Text = ""; textBox66.Text = ""; textBox67.Text = ""; textBox68.Text = "";
           textBox70.Text = ""; textBox71.Text = ""; textBox72.Text = ""; textBox73.Text = ""; textBox74.Text = ""; textBox69.Text = "";
           textBox93.Text = ""; textBox94.Text = "";
       }

       private void KDOWN_Click(object sender, EventArgs e)
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");
           button60_Click(button60, new EventArgs());
           Delay(500);
           button64_Click(button64, new EventArgs());
           Delay(500);
           button88_Click(button88, new EventArgs());
           Delay(500);
           button93_Click_1(button93, new EventArgs());
           Delay(500);
           dispel();
         //  Reading_Click(Reading, new EventArgs());
           standard_position = 2;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }
       int biaozhi = 0;
       private void KUP_Click(object sender, EventArgs e)//二次短路
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");
           button60_Click(button60, new EventArgs());//开启电源
           Delay(500);
           button7_Click(button7, new EventArgs());//A相短路
           Delay(500);
           button30_Click(button30, new EventArgs());//B相短路
           Delay(500);
           button47_Click(button47, new EventArgs());//C相短路
           Delay(500);
           dispel();
          // Reading_Click(Reading, new EventArgs());    
           standard_position = 6;
           biaozhi = 1;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }

       private void button96_Click(object sender, EventArgs e)//二次开路
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");
           button60_Click(button60, new EventArgs());
           Delay(500);
           button6_Click(button8, new EventArgs());
           Delay(500);
           button29_Click(button29, new EventArgs());
           Delay(500);
           button46_Click(button46, new EventArgs());
           Delay(500);
           dispel();
          // Reading_Click(Reading, new EventArgs()); 
           standard_position = 7;
           biaozhi = 1;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }

       private void diode_Click(object sender, EventArgs e)//串接二极管
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("3");
           button60_Click(button60, new EventArgs());
           Delay(500);
           button9_Click(button9, new EventArgs());
           Delay(500);
           button32_Click(button32, new EventArgs());
           Delay(500);
           button49_Click(button49, new EventArgs());
           dispel();
         //  Reading_Click(Reading, new EventArgs()); 
           standard_position = 8;
           timer4.Interval = 50000;
           timer4.Start();     //启动定时器
       }

       #region 空
       private void groupBox9_Enter(object sender, EventArgs e)
       {

       }

       private void AKUP10_TextChanged(object sender, EventArgs e)
       {

       }

       private void AKUP15_TextChanged(object sender, EventArgs e)
       {

       }

       private void label262_Click(object sender, EventArgs e)
       {

       }

       private void groupBox6_Enter(object sender, EventArgs e)
       {

       }

       private void label275_Click(object sender, EventArgs e)
       {

       }

        private void AVpp_5K_TextChanged(object sender, EventArgs e)
        {

        }

        private void A_Amp_5_D_TextChanged(object sender, EventArgs e)
        {

        }

        private void label277_Click(object sender, EventArgs e)
       {

       }

       private void textBox92_TextChanged(object sender, EventArgs e)
       {

       }

       private void A_POWER_H_S_TextChanged(object sender, EventArgs e)
       {

       }
       #endregion

       private void button96_Click_1(object sender, EventArgs e)//开路扫频
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");    
           dispel();
           button60_Click(button60, new EventArgs());
           Delay(500);
           button6_Click(button8, new EventArgs());
           Delay(500);
           button29_Click(button29, new EventArgs());
           Delay(500);
           button46_Click(button46, new EventArgs());
           Delay(500);
           standard_position = 4;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }

       private void button97_Click(object sender, EventArgs e)//短路扫频
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("1");
           dispel();
           button60_Click(button60, new EventArgs());
           Delay(500);
           button7_Click(button7, new EventArgs());
           Delay(500);
           button30_Click(button30, new EventArgs());
           Delay(500);
           button47_Click(button47, new EventArgs());
           Delay(500);
           standard_position = 5;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }

       private void textBox79_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox80_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox75_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox87_TextChanged(object sender, EventArgs e)
       {

       }

       private void A_DL_value_TextChanged(object sender, EventArgs e)
       {

       }

       private void button98_Click(object sender, EventArgs e)//校准工频电流复位
       {
           A_POWER_L_K.Text = "10000"; B_POWER_L_K.Text = "10000"; B_POWER_L_K.Text = "10000";
           A_POWER_L_S.Text = "0"; B_POWER_L_S.Text = "0"; B_POWER_L_S.Text = "0";
           A_POWER_H_K.Text = "10000"; B_POWER_H_K.Text = "10000"; B_POWER_H_K.Text = "10000";
           A_POWER_H_S.Text = "0"; B_POWER_H_S.Text = "0"; B_POWER_H_S.Text = "0";
           A_KL_value.Text = "593750"; B_KL_value.Text = "593750"; B_KL_value.Text = "593750";
           A_DL_value.Text = "008437"; B_DL_value.Text = "008437"; B_DL_value.Text = "008437";
           textBox100.Text = ""; textBox99.Text = ""; textBox106.Text = "";
           textBox107.Text = ""; textBox105.Text = ""; textBox103.Text = "";
           textBox98.Text = ""; textBox97.Text = ""; textBox104.Text = "";
           textBox96.Text = ""; textBox95.Text = ""; textBox102.Text = "";
           standard_position = 0;
           standard_stream = 0;
           standard_sum = 0;
           timer4.Stop();     //关闭定时器
           button86_Click(button86,new EventArgs());
           Delay(500);
       }

       private void button100_Click(object sender, EventArgs e)
       {
           A_Amp_5_D.Text = "84"; B_Amp_5_D.Text = "84"; C_Amp_5_D.Text = "84";
           A_Amp_5_U.Text = "61250"; B_Amp_5_U.Text = "61250"; C_Amp_5_U.Text = "61250";
           A_Amp_10_D.Text = "47"; B_Amp_10_D.Text = "47"; C_Amp_10_D.Text = "47";
           A_Amp_10_U.Text = "59375"; B_Amp_10_U.Text = "59375"; C_Amp_10_U.Text = "59375";
           A_Amp_15_D.Text = "234"; B_Amp_15_D.Text = "234"; C_Amp_15_D.Text = "234";
           A_Amp_15_U.Text = "56250"; B_Amp_15_U.Text = "56250"; C_Amp_15_U.Text = "56250";
           A_KL_SCANFRE.Text = "4"; B_KL_SCANFRE.Text = "4"; C_KL_SCANFRE.Text = "4";
           A_DL_SCANFRE.Text = "82"; B_DL_SCANFRE.Text = "82"; C_DL_SCANFRE.Text = "82";
           A_KL_GAIN_K1.Text = "99"; B_KL_GAIN_K1.Text = "99"; C_KL_GAIN_K1.Text = "99";
           A_KL_GAIN_K2.Text = "99"; B_KL_GAIN_K2.Text = "99"; C_KL_GAIN_K2.Text = "99";
           A_FRELT_DOWN.Text = "17975"; B_FRELT_DOWN.Text = "17975"; C_FRELT_DOWN.Text = "17975";
           standard_position = 0;
           timer4.Stop();     //关闭定时器
           button79_Click(button79, new EventArgs());
           Delay(500);
       }

       private void A_FRELT_DOWN_TextChanged(object sender, EventArgs e)
       {

       }

       private void button99_Click(object sender, EventArgs e)
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");
           button60_Click(button60, new EventArgs());
           Delay(500);
           button65_Click(button65, new EventArgs());
           Delay(500);
           button89_Click(button89, new EventArgs());
           Delay(500);
           button94_Click(button94, new EventArgs());
           dispel();
           //  Reading_Click(Reading, new EventArgs());
           standard_position = 9;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }

       private void textBox56_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox60_TextChanged(object sender, EventArgs e)
       {

       }

       private void button101_Click(object sender, EventArgs e)//50K
       {
           label282.Text = "等待中...";
           textBox41.Text = "";
           textBox46.Text = "";
           textBox41.AppendText("0");
           textBox46.AppendText("0");
           button60_Click(button60, new EventArgs());
           Delay(500);
           button63_Click(button53, new EventArgs());
           Delay(500);
           button84_Click_1(button84, new EventArgs());
           Delay(500);
           button92_Click(button92, new EventArgs());
           dispel();
           //  Reading_Click(Reading, new EventArgs());
           standard_position = 11;
           timer4.Interval = 80000;
           timer4.Start();     //启动定时器
       }
       #region 空
       private void label265_Click(object sender, EventArgs e)
       {

       }

       private void textBox86_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox97_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox76_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox88_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox84_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox98_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox100_TextChanged(object sender, EventArgs e)
       {

       }

       private void label273_Click(object sender, EventArgs e)
       {

       }

       private void textBox106_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox78_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox107_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox101_TextChanged(object sender, EventArgs e)
       {

       }

       private void textBox85_TextChanged(object sender, EventArgs e)
       {

       }
       #endregion

       private void label282_Click(object sender, EventArgs e)
       {

       }

       private void label274_Click(object sender, EventArgs e)
       {

       }

       private void button102_Click(object sender, EventArgs e)
       {
           A_POWER_H1.Text = "2048"; A_POWER_H5.Text = "2048"; B_POWER_H1.Text = "2048"; B_POWER_H5.Text = "2048"; C_POWER_H1.Text = "2048"; C_POWER_H5.Text = "2048";
           A_HIGHT_H1.Text = "2048"; A_HIGHT_H5.Text = "2048"; B_HIGHT_H1.Text = "2048"; B_HIGHT_H5.Text = "2048"; C_HIGHT_H1.Text = "2048"; C_HIGHT_H5.Text = "2048";
           textBox13.Text = "2048"; textBox14.Text = "2048"; textBox15.Text = "2048"; textBox16.Text = "2048"; textBox17.Text = "2048"; textBox18.Text = "2048";
           textBox39.Text = "2048"; textBox40.Text = "2048"; textBox47.Text = "2048"; textBox48.Text = "2048"; textBox49.Text = "2048"; textBox50.Text = "2048";
       }

       private void textBox40_TextChanged_2(object sender, EventArgs e)
       {

       }

       private void button103_Click(object sender, EventArgs e)/*按模块类别查找函数*/
       {
           StringBuilder buf = new StringBuilder(1024);//指定的buf大小必须大于传入的字符长度
           buf.Append("4258");
           int outdata = RS232dll.search_Moduletype(buf);
           string strout = outdata.ToString();
           textBox6.AppendText(strout + "\r\n");
       }

       private void button104_Click(object sender, EventArgs e)/*按模块编号查找函数*/
       {
           StringBuilder buf = new StringBuilder(1024);//指定的buf大小必须大于传入的字符长度
           buf.Append("02811827CC2AC7D83184");
           int outdata = RS232dll.search_Serialnumber(buf);
           string strout = outdata.ToString();
           textBox6.AppendText(strout + "\r\n");
       }

       private void button105_Click(object sender, EventArgs e)/*按厂家查找函数*/
       {
           StringBuilder buf = new StringBuilder(1024);//指定的buf大小必须大于传入的字符长度
           buf.Append("444b");
           int outdata = RS232dll.search_Manufacturer(buf);
           string strout = outdata.ToString();
           textBox6.AppendText(strout + "\r\n");
       }

       private void button106_Click(object sender, EventArgs e)/*按设备类别查找函数*/
       {
           StringBuilder buf = new StringBuilder(1024);//指定的buf大小必须大于传入的字符长度
           buf.Append("40");
           int outdata = RS232dll.search_DeviceType(buf);
           string strout = outdata.ToString();
           textBox6.AppendText(strout + "\r\n");

       }

       private void button107_Click(object sender, EventArgs e)/*按邮编查找函数*/
       {
           StringBuilder buf = new StringBuilder(1024);//指定的buf大小必须大于传入的字符长度
           buf.Append("420000");
           int outdata = RS232dll.search_AreaCode(buf);
           string strout = outdata.ToString();
           textBox6.AppendText(strout + "\r\n");
       }
    }
}
