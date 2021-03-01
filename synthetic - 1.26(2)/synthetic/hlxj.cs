using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace synthetic
{
    class hlxj
    {
        [System.Runtime.InteropServices.DllImport("hlxj.dll")]
        public static extern void String_Decrypt(byte[] pBuffer, ushort BufferLength);
    }
}
