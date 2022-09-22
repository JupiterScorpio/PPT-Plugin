using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace PPTAddin
{
    public class HookHelper
    {
        #region  Enumeration definition
        /// <summary>
        ///  Operation type
        /// </summary>
        public enum HookType
        {
            KeyOperation,//Keyboard operation
            MouseOperation//Mouse operation
        }

        /// <summary>
        ///  Bottom hook ID
        /// </summary>
        public enum WH_Codes : int
        {
            //Low-level keyboard hook
            WH_KEYBOARD_LL = 2,//2: Monitoring keyboard messages and thread hook; 13: Global mouse hook monitoring

            //Bottom mouse hook
            WH_MOUSE_LL = 7, //7: monitor mouse hook; 14: global keyboard hook monitor mouse information

            //nCode is 0
            HC_ACTION = 0
        }

        /// <summary>
        ///  Mouse button identification
        /// </summary>
        public enum WM_MOUSE : int
        {
            /// <summary>
            ///  Mouse start
            /// </summary>
            WM_MOUSEFIRST = 0x200,

            /// <summary>
            ///  Mouse movement
            /// </summary>
            WM_MOUSEMOVE = 0x200,

            /// <summary>
            ///  Left click
            /// </summary>
            WM_LBUTTONDOWN = 0x201,

            /// <summary>
            ///  Left button release
            /// </summary>
            WM_LBUTTONUP = 0x202,

            /// <summary>
            ///  Double left click
            /// </summary>
            WM_LBUTTONDBLCLK = 0x203,

            /// <summary>
            ///  Right click
            /// </summary>
            WM_RBUTTONDOWN = 0x204,

            /// <summary>
            ///  Right click to release
            /// </summary>
            WM_RBUTTONUP = 0x205,

            /// <summary>
            ///  Right double click
            /// </summary>
            WM_RBUTTONDBLCLK = 0x206,

            /// <summary>
            ///  Middle button press
            /// </summary>
            WM_MBUTTONDOWN = 0x207,

            /// <summary>
            ///  Middle button release
            /// </summary>
            WM_MBUTTONUP = 0x208,

            /// <summary>
            ///  Middle button double click
            /// </summary>
            WM_MBUTTONDBLCLK = 0x209,

            /// <summary>
            ///  Scroll wheel
            /// </summary>
            /// <remarks>This message is only supported after WINNT4.0</remarks>
            WM_MOUSEWHEEL = 0x020A,
            WM_LBUTTONCLK=0x20B
        }

        /// <summary>
        ///  Keyboard key identification
        /// </summary>
        public enum WM_KEYBOARD : int
        {
            /// <summary>
            ///  Non-system key press
            /// </summary>
            WM_KEYDOWN = 0x100,

            /// <summary>
            ///  Non-system key release
            /// </summary>
            WM_KEYUP = 0x101,

            /// <summary>
            ///  System button press
            /// </summary>
            WM_SYSKEYDOWN = 0x104,

            /// <summary>
            ///  System button release
            /// </summary>
            WM_SYSKEYUP = 0x105
        }

        /// <summary>
        ///  SetWindowPos flag enumeration
        /// </summary>
        /// <remarks>For details, please refer to the description of the SetWindowPos function in MSDN</remarks>
        public enum SetWindowPosFlags : int
        {
            /// <summary>
            /// 
            /// </summary>
            SWP_NOSIZE = 0x0001,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOMOVE = 0x0002,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOZORDER = 0x0004,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOREDRAW = 0x0008,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOACTIVATE = 0x0010,

            /// <summary>
            /// 
            /// </summary>
            SWP_FRAMECHANGED = 0x0020,

            /// <summary>
            /// 
            /// </summary>
            SWP_SHOWWINDOW = 0x0040,

            /// <summary>
            /// 
            /// </summary>
            SWP_HIDEWINDOW = 0x0080,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOCOPYBITS = 0x0100,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOOWNERZORDER = 0x0200,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOSENDCHANGING = 0x0400,

            /// <summary>
            /// 
            /// </summary>
            SWP_DRAWFRAME = 0x0020,

            /// <summary>
            /// 
            /// </summary>
            SWP_NOREPOSITION = 0x0200,

            /// <summary>
            /// 
            /// </summary>
            SWP_DEFERERASE = 0x2000,

            /// <summary>
            /// 
            /// </summary>
            SWP_ASYNCWINDOWPOS = 0x4000

        }

        #endregion  Enumeration definition

        #region  Structure definition

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        /// <summary>
        ///  Mouse hook event structure definition
        /// </summary>
        /// <remarks>For detailed description, please refer to the description of MSLLHOOKSTRUCT in MSDN</remarks>
        [StructLayout(LayoutKind.Sequential)]
        public struct MouseHookStruct
        {
            /// <summary>
            /// Specifies a POINT structure that contains the x- and y-coordinates of the cursor, in screen coordinates.
            /// </summary>
            public POINT Point;

            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public UInt32 ExtraInfo;
        }

        /// <summary>
        ///  Keyboard hook event structure definition
        /// </summary>
        /// <remarks>For details, please refer to the description of KBDLLHOOKSTRUCT in MSDN</remarks>
        [StructLayout(LayoutKind.Sequential)]
        public struct KeyboardHookStruct
        {
            /// <summary>
            /// Specifies a virtual-key code. The code must be a value in the range 1 to 254. 
            /// </summary>
            public UInt32 VKCode;

            /// <summary>
            /// Specifies a hardware scan code for the key.
            /// </summary>
            public UInt32 ScanCode;

            /// <summary>
            /// Specifies the extended-key flag, event-injected flag, context code, 
            /// and transition-state flag. This member is specified as follows. 
            /// An application can use the following values to test the keystroke flags. 
            /// </summary>
            public UInt32 Flags;

            /// <summary>
            /// Specifies the time stamp for this message. 
            /// </summary>
            public UInt32 Time;

            /// <summary>
            /// Specifies extra information associated with the message. 
            /// </summary>
            public UInt32 ExtraInfo;
        }

        #endregion  Structure definition
    }
}
