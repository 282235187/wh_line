using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace synthetic
{
    public class DLT64507
    {
        //  [DllImport("hlxj.dll")]
        //      public static extern ushort DLT645_write(byte feNum, byte ctrl, byte[] addr, uint id, byte[] recv, byte recvLen, byte[] p);
        public static ushort DLT645_write(byte feNum, byte ctrl, byte[] addr, uint id, byte[] recv, byte recvLen, byte[] p)
        {
            //   ushort i = 0;
            ushort cnt = 0;

            for (ushort i = 0; i < feNum; i++)
            {
                p[cnt++] = 0xfe;                             //根据feNum值填fe数量到buffer
            }

            p[cnt++] = 0x68;                                  //填帧头

            for (ushort i = 0; i < 6; i++)
            {
                p[cnt++] = addr[i];                          //填地址
            }
            p[cnt++] = 0x68;                                 //填帧头
            p[cnt++] = ctrl;                                 //填控制码
            p[cnt++] = (byte)(8 + recvLen);                          //填数据域长度id+usrID+data

            for (ushort i = 0; i < 4; i++)                               //填数据标识
            {
                p[cnt++] = (byte)(id >> 8 * i);
            }

            for (ushort i = 0; i < 4; i++)
            {
                p[cnt++] = 0x00;                         //操作者代码
            }

            // memcpy(&p[cnt], &recv[0], recvLen);              //取数据内容
            for (ushort i = 0; i < recvLen; i++)
            {
                p[cnt++] = recv[i];
            }

            for (ushort i = 0; i < 8 + recvLen; i++)
            {
                p[feNum + 10 + i] += 0x33;
            }

            p[cnt] = 0x00;                                   //cs
            for (ushort i = feNum; i < cnt; i++)
            {
                p[cnt] += p[i];                              //填校验码
            }
            cnt += 1;

            p[cnt++] = 0x16;                                 //填结束符
            return cnt;
        }




        public static void rever_char(ref byte[] c, int n)
        {
            byte temp = 0;
            int j = 0;

            for (int i = 0; i < n / 2; i++)
            {
                j = n - 1 - i;
                temp = c[i];
                c[i] = c[j];
                c[j] = temp;
            }
        }


        public static void DLT_Check(ref byte[] rcv, int n)
        {
            int i;
            for (i = 0; i < n; i++)
            {
                if (rcv[i] == 0x68 && rcv[i + 7] == 0x68)
                {
                    if ((rcv[i + 9] + 12) <= (n - i))
                    {

                    }
                }
            }

        }

    }

}
