using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace synthetic
{
    class RS232dll
    {
            [DllImport(@"RS232dll.dll", EntryPoint = "search_Moduletype", CallingConvention = CallingConvention.Cdecl)]
            public static extern int search_Moduletype(System.Text.StringBuilder msg);/*按模块类别查找函数*/

            [DllImport(@"RS232dll.dll", EntryPoint = "search_Serialnumber", CallingConvention = CallingConvention.Cdecl)]
            public static extern int search_Serialnumber(System.Text.StringBuilder msg);/*按模块编号查找函数*/

            [DllImport(@"RS232dll.dll", EntryPoint = "search_Manufacturer", CallingConvention = CallingConvention.Cdecl)]
            public static extern int search_Manufacturer(System.Text.StringBuilder msg);/*按厂家查找函数*/

            [DllImport(@"RS232dll.dll", EntryPoint = "search_DeviceType", CallingConvention = CallingConvention.Cdecl)]
            public static extern int search_DeviceType(System.Text.StringBuilder msg);/*按设备类别查找函数*/

            [DllImport(@"RS232dll.dll", EntryPoint = "search_AreaCode", CallingConvention = CallingConvention.Cdecl)]
            public static extern int search_AreaCode(System.Text.StringBuilder msg);/*按邮编查找函数*/
    }
}
