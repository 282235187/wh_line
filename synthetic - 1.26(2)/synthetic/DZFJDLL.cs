using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace synthetic
{
    class DZFJDLL
    {
        [System.Runtime.InteropServices.DllImport("DZFJDLL.dll")]
        public static extern void String_Encrypt(byte[] pBuffer, ushort BufferLength);
    }
}
