using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace synthetic
{
    /// <summary>
    /// 
    /// </summary>
    public static class ByteOperate
    {
        /// <summary>
        /// 将整型数据转换成
        /// </summary>
        /// <param name="intvalue">正整数</param>
        /// <param name="count">数组长度</param>
        /// <param name="HL">true表示高位在前</param>
        /// <returns>字节数组</returns>
        public static byte[] GetBytes(int intvalue, int count, bool HL)
        {
            byte[] bytes = new byte[count];
            for (int i = 0; i < count; i++)
            {
                int index = i;
                if (HL)
                    index = count - i - 1;
                bytes[index] = Convert.ToByte(intvalue % 256);
                intvalue = intvalue / 256;
            }
            return bytes;
        }
        /// <summary>
        /// 查找指定数组中指定的数据
        /// </summary>
        /// <param name="bytes">数据</param>
        /// <param name="bt">查找的数据</param>
        /// <param name="Length">查找的长度</param>
        /// <returns>返回-1表示未找到指定的字节</returns>
        public static int FindByte(byte[] bytes, byte bt, int Length)
        {
            for (int i = 0; i < bytes.Length && i < Length; i++)
            {
                if (bytes[i] == bt) return i;
            }
            return -1;
        }
        /// <summary>
        /// 从数组中获取整型数据
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="HL">true表示高位在前</param>
        /// <returns>整型数据</returns>
        public static int GetInt32(byte[] bytes, bool HL)
        {

            int value = bytes[0];
            for (int i = 1; i < bytes.Length && i < 4; i++)
            {
                if (HL)
                    value = (value << 8) + bytes[i];
                else
                {
                    int c = bytes[i];
                    value = value + (c << (i * 8));
                }
            }
            return value;
        }
        /// <summary>
        /// BCD_TO_BYTE
        /// </summary>
        /// <param name="bt">待转换字节</param>
        /// <returns>转换后字节</returns>
        public static byte BCD_TO_BYTE(byte bt)
        {
            int a1 = bt >> 4;
            int a2 = bt & 0x0f;
            return Convert.ToByte(a1 * 10 + a2);
        }
        /// <summary>
        /// BYTE_TO_BCD
        /// </summary>
        /// <param name="bt">待转换字节</param>
        /// <returns>转换后字节</returns>
        public static byte BYTE_TO_BCD(byte bt)
        {
            return Convert.ToByte(((bt / 10) << 4) | (bt % 10));
        }
        /// <summary>
        /// 字节数组转为16进制字符串，16进制字符串每字节之间以空格分开
        /// </summary>
        /// <param name="bytes">待转换字节数组</param>
        /// <param name="outstring">输出字符串</param>
        /// <returns></returns>
        public static bool GetString(byte[] bytes, ref string outstring)
        {
            return GetString(bytes, ref outstring, " ");
        }
        /// <summary>
        /// 获取ip字符串
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string GetIPString(byte[] bytes)
        {
            string str = bytes[0].ToString();
            for (int i = 1; i < bytes.Length; i++)
            {
                str = str + "." + bytes[i].ToString();
            }
            return str;
        }
        /// <summary>
        /// 获取IP字节数组
        /// </summary>
        /// <param name="ipstring"></param>
        /// <returns></returns>
        public static byte[] GetIPBytes(string ipstring)
        {
            byte[] bytes = new byte[4];
            string[] str = ipstring.Split('.');
            for (int i = 0; i < 4; i++)
            {
                bytes[i] = Convert.ToByte(str[i]);
            }
            return bytes;
        }
        /// <summary>
        /// 字节数组转为16进制字符串，16进制字符串每字节之间以指定分隔符分开
        /// </summary>
        /// <param name="bytes">待转换字节数组</param>
        /// <param name="outstring">输出字符串</param>
        /// <param name="split">16进制字符串每字节之间的分隔符</param>
        /// <returns></returns>
        public static bool GetString(byte[] bytes, ref string outstring, string split)
        {
            try
            {
                outstring = "";
                for (int i = 0; i < bytes.Length; i++)
                {
                    outstring += bytes[i].ToString("X2") + split;
                }
                outstring = outstring.Trim();
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// 从指定字节数组中截取从指定索引开始的若干个数据
        /// </summary>
        /// <param name="source">指定数组</param>
        /// <param name="startindex">指定索引</param>
        /// <param name="bytescount">数据个数</param>
        /// <returns>截取的数组</returns>
        public static byte[] GetBytes(byte[] source, int startindex, int bytescount)
        {
            byte[] bytes = new byte[bytescount];
            Array.ConstrainedCopy(source, startindex, bytes, 0, bytescount);
            return bytes;
        }
        /// <summary>
        /// 把指定字符串转换成字节数组
        /// </summary>
        /// <param name="bytes">转换后的字节数组</param>
        /// <param name="str">原始16进制字符串</param>
        /// <returns>转换后的数组长度</returns>
        public static int GetBytes(ref byte[] bytes, string str)
        {
            byte[] bytestmp = new byte[(str.Length + 1) / 2];
            int len = 0;
            for (int i = 0; i < bytestmp.Length; i++)
            {
                try
                {
                    bytestmp[i] = Convert.ToByte(str.Substring(0, 2), 16);
                    str = str.Remove(0, 2);
                    len++;
                }
                catch
                {
                    break;
                }
            }
            if (len > 0)
                bytes = GetBytes(bytestmp, 0, len);
            return len;
        }
        /// <summary>
        /// 把指定字符串转换成字节数组
        /// </summary>
        /// <param name="bytes">转换后的字节数组</param>
        /// <param name="str">原始16进制字符串</param>
        /// <param name="split">需要过滤掉的字符</param>
        /// <returns>转换后的数组长度</returns>
        public static int GetBytes(ref byte[] bytes, string str, char split)
        {
            for (int i = str.Length - 1; i >= 0; i--)
            {
                if (str[i] == split)
                {
                    str = str.Remove(i, 1);
                }
            }
            return GetBytes(ref bytes, str);
        }
        /// <summary>
        /// 计算和校验
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="startindex">起始位</param>
        /// <param name="len">计算长度</param>
        /// <returns>和校验</returns>
        public static byte GetSumCheckCRC8(byte[] bytes, int startindex, int len)
        {
            byte cs = 0;
            for (int i = 0; i < len; i++)
            {
                cs = Convert.ToByte((cs + bytes[startindex + i]) % 256);
            }
            return cs;
        }
        /// <summary>
        /// 计算IntelHex8校验
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="startindex">起始位</param>
        /// <param name="len">计算长度</param>
        /// <returns>IntelHex8校验</returns>
        public static byte GetSumCheckIntelHex8(byte[] bytes, int startindex, int len)
        {
            uint cs = GetSumCheckCRC8(bytes, startindex, len);
            cs = ~cs;
            return Convert.ToByte((0x01 + cs) % 256);
        }
        /// <summary>
        /// 计算BCC校验
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="startindex">起始位</param>
        /// <param name="len">计算长度</param>
        /// <returns>BCC校验</returns>
        public static byte GetSumCheckBCC8(byte[] bytes, int startindex, int len)
        {
            byte cs = 0;
            for (int i = 0; i < len; i++)
            {
                cs ^= bytes[startindex + i];
            }
            return cs;
        }
        /// <summary>
        /// 获取ASCII字符串,遇到0结束
        /// </summary>
        /// <param name="bytes">字节数组</param>
        /// <param name="startindex">起始索引</param>
        /// <param name="len">最大长度</param>
        /// <returns>字符串</returns>
        public static string GetASCIIString(byte[] bytes, int startindex, int len)
        {
            string str = "";
            for (int i = startindex, j = 0; i < bytes.Length && j < len; i++, j++)
            {
                if (bytes[i] == 0) break;
                str += (char)bytes[i];
            }
            return str;
        }
        /// <summary>
        /// 拼接字节数组
        /// </summary>
        /// <param name="bts1">该数组在前</param>
        /// <param name="bts2">该数组在后</param>
        /// <returns>拼接后的数组</returns>
        public static byte[] BytesAdd(byte[] bts1, byte[] bts2)
        {
            byte[] bts = new byte[bts1.Length + bts2.Length];
            int index = 0;
            for (; index < bts1.Length; index++)
            {
                bts[index] = bts1[index];
            }
            for (; index - bts1.Length < bts2.Length; index++)
            {
                bts[index] = bts2[index - bts1.Length];
            }
            return bts;
        }

        ///<summary>
        /// 字符串转16进制字节数组
        ///</summary>
        ///<param name="hexString"></param>
        ///<returns></returns>
        private static byte[] strToToHexByte(string hexString)
        {
            hexString = hexString.Replace("", "");
            if ((hexString.Length % 2) != 0)
                hexString += "";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }

        ///<summary>
        /// 字节数组转16进制字符串
        ///</summary>
        ///<param name="bytes"></param>
        ///<returns></returns>
        public static string byteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }

        ///<summary>
        /// 从汉字转换到16进制
        ///</summary>
        ///<param name="s"></param>
        ///<param name="charset">编码,如"utf-8","gb2312"</param>
        ///<param name="fenge">是否每字符用逗号分隔</param>
        ///<returns></returns>
        public static string ToHex(string s, string charset, bool fenge)
        {
            if ((s.Length % 2) != 0)
            {
                s += "";//空格
                //throw new ArgumentException("s is not valid chinese string!"); 
            }
            System.Text.Encoding chs = System.Text.Encoding.GetEncoding(charset);
            byte[] bytes = chs.GetBytes(s);
            string str = "";
            for (int i = 0; i < bytes.Length; i++)
            {
                str += string.Format("{0:X}", bytes[i]);
                if (fenge && (i != bytes.Length - 1))
                {
                    str += string.Format("{0}", ",");
                }
            }
            return str.ToLower();
        }

        ///<summary>
        /// 从16进制转换成汉字
        ///</summary>
        ///<param name="hex"></param>
        ///<param name="charset">编码,如"utf-8","gb2312"</param>
        ///<returns></returns>
        public static string UnHex(string hex, string charset)
        {
            if (hex == null)
                throw new ArgumentNullException("hex");
            hex = hex.Replace(",", "");
            hex = hex.Replace("\n", "");
            hex = hex.Replace("\\", "");
            hex = hex.Replace("", "");
            if (hex.Length % 2 != 0)
            {
                hex += "20";//空格
            }
            // 需要将 hex 转换成 byte 数组。
            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                try
                {
                    // 每两个字符是一个 byte。
                    bytes[i] = byte.Parse(hex.Substring(i * 2, 2),
                    System.Globalization.NumberStyles.HexNumber);
                }
                catch
                {
                    // Rethrow an exception with custom message.
                    throw new ArgumentException("hex is not a valid hex number!", "hex");
                }
            }
            System.Text.Encoding chs = System.Text.Encoding.GetEncoding(charset);
            return chs.GetString(bytes);
        }

        public static string inttostr(int Value)
        {
            byte i = 0;
            string str = "";
            str += Value;
            return str;
        }

        public static int BcdtoInt(byte[] bytes, bool HL)
        {

            int value = bytes[0];
            for (int i = 1; i < bytes.Length && i < 4; i++)
            {
                if (HL)
                    value = value + (bytes[i] >> 4) * 10 + (bytes[i] & 0x0f);
                else
                {
                    int c = bytes[i];
                    value = value + (c << (i * 8));
                }
            }
            return value;
        }
    }
}
