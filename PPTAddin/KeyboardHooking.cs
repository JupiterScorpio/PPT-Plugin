using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace PPTAddin
{
    class KeyboardHooking
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        //declare the mouse hook constant.
        //For other hook types, you can obtain these values from Winuser.h in the Microsoft SDK.

        private const int WH_KEYBOARD = 2; // mouse
        private const int HC_ACTION = 0;

        private const int WH_KEYBOARD_LL = 13; // keyboard
        private const int WM_KEYDOWN = 0x0100;

        public static void SetHook()
        {
            // Ignore this compiler warning, as SetWindowsHookEx doesn't work with ManagedThreadId
#pragma warning disable 618
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618

        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }

        //Note that the custom code goes in this method the rest of the class stays the same.
        //It will trap if BOTH keys are pressed down.
        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
            else
            {
                if (nCode == HC_ACTION)
                {
                    Keys keyData = (Keys)wParam;

                    // CTRL + SHIFT + 7
                    if ((BindingFunctions.IsKeyDown(Keys.CapsLock) == true)&&(ThisAddIn.nCtrlOShift==1))
                        //&& (BindingFunctions.IsKeyDown(Keys.ShiftKey) == true)
                        //&& (BindingFunctions.IsKeyDown(keyData) == true) && (keyData == Keys.D7))
                    {
                        // DO SOMETHING HERE
                        ThisAddIn.bkeypress = true;
                    }
                    if ((BindingFunctions.IsKeyDown(Keys.CapsLock) == false) && (ThisAddIn.nCtrlOShift == 1))
                        ThisAddIn.bkeypress = false;
                    // CTRL + 7
                    if ((BindingFunctions.IsKeyDown(Keys.ShiftKey) == true)&& (ThisAddIn.nCtrlOShift == 2))
                        //&& (BindingFunctions.IsKeyDown(keyData) == true) && (keyData == Keys.D7))
                    {
                        // DO SOMETHING HERE
                        ThisAddIn.bkeypress = true;
                    }
                    if ((BindingFunctions.IsKeyDown(Keys.ShiftKey) == false) && (ThisAddIn.nCtrlOShift == 2))
                        ThisAddIn.bkeypress = false;
                    if(BindingFunctions.IsKeyDown(Keys.Delete))
                    {
                        ThisAddIn.bkeypress = true;
                        ThisAddIn.OnShapeDel();
                    }

                }
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }
    }

    public class BindingFunctions
    {
        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }
    }
}
