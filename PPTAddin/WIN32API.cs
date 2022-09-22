using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using System.Drawing;

namespace PPTAddin
{
    class Win32API
    {
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }
        #region  DLL import
        /// <summary>
        ///  Used to set the window
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="hWndInsertAfter"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="uFlags"></param>
        /// <returns></returns>
        /// 
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ScreenToClient(IntPtr hWnd, ref ThisAddIn.POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern bool GetClientRect(IntPtr hwnd, ref Rectangle lpRect);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rectangle lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT lpPoint);
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        /// <summary>
        ///  Install hook
        /// </summary>
        /// <param name="idHook">Hook type, mouse\keyboard\macro etc. 10 kinds</param>
        /// <param name="lpfn">Hooked functions, functions used to process intercepted messages, global functions</param>
        /// <param name="hInstance">The handle of the current process,
        ///  NULL: a thread created by the current process, the child process is located in the current process;
        ///  0 (IntPtr.Zero): the hook procedure is associated with all threads, that is, the global hook</param>
        /// <param name="threadId">Set the thread ID to be hooked. NULL is a global hook</param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr SetWindowsHookEx(IntPtr idHook, HookProc lpfn, IntPtr pInstance, uint threadId);

        /// <summary>
        ///  Uninstall hook
        /// </summary>
        /// <param name="idHook"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(IntPtr pHookHandle);

        /// <summary>
        ///  Pass hook
        ///  Used to continue to pass the intercepted message, otherwise other programs may not get the corresponding message
        /// </summary>
        /// <param name="idHook"></param>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(IntPtr pHookHandle, int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>
        ///  Convert current key information
        /// </summary>
        /// <param name="uVirtKey"></param>
        /// <param name="uScanCode"></param>
        /// <param name="lpbKeyState"></param>
        /// <param name="lpwTransKey"></param>
        /// <param name="fuState"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int ToAscii(UInt32 uVirtKey, UInt32 uScanCode, byte[] lpbKeyState, byte[] lpwTransKey, UInt32 fuState);

        /// <summary>
        ///  Get button status
        /// </summary>
        /// <param name="pbKeyState"></param>
        /// <returns>Non-zero means success</returns>
        [DllImport("user32.dll")]
        public static extern int GetKeyboardState(byte[] pbKeyState);

        [DllImport("user32.dll")]
        public static extern short GetKeyStates(int vKey);

        /// <summary>
        ///  Get the current thread ID
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        /// <summary>
        ///  Screen capture location
        /// </summary>
        /// <param name="hdcDest">Handle of target device </param>
        /// <param name="nXDest">X coordinate of the upper left corner of the target object</param>
        /// <param name="nYDest">Y coordinate of the upper left corner of the target object</param>
        /// <param name="nWidth">The width of the rectangle of the target object</param>
        /// <param name="nHeight">The height of the rectangle of the target object </param>
        /// <param name="hdcSrc">Handle of source device </param>
        /// <param name="nXSrc">X coordinate of the upper left corner of the source object </param>
        /// <param name="nYSrc">Y coordinate of the upper left corner of the source object </param>
        /// <param name="dwRop">Operation value of raster </param>
        /// <returns></returns>
        [DllImportAttribute("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight,
            IntPtr hdcSrc, int nXSrc, int nYSrc, System.Int32 dwRop);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lpszDriver">Driver name</param>
        /// <param name="lpszDevice">Equipment name</param>
        /// <param name="lpszOutput">Useless, can be set to "NULL"</param>
        /// <param name="lpInitData">Arbitrary printer data</param>
        /// <returns></returns>
        [DllImportAttribute("gdi32.dll")]
        private static extern IntPtr CreateDC(string lpszDriver, string lpszDevice, string lpszOutput, IntPtr lpInitData);
        #endregion  DLL import

        /// <summary>
        ///  Hook delegation statement
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    }
}


    
