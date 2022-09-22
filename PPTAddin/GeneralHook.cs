using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTAddin
{
    public class GeneralHook
    {
        #region  Define variables
        //Define mouse hook processing function
        private Win32API.HookProc MouseHookProcedure = null;
        //Define keyboard hook processing function
        private Win32API.HookProc KeyboardProcDelegate = null;

        //Define keyboard hook handle
        private IntPtr khook;
        //Define the mouse hook handle
        private IntPtr hHook = IntPtr.Zero;

        //Define mouse events
        public event MouseEventHandler OnMouseActivity;
        #endregion

        /// <summary>
        ///  Install hook
        /// </summary>
        public void InstallHook(HookHelper.HookType hookType)
        {
            if (hookType == HookHelper.HookType.KeyOperation)
            {
                if (khook == IntPtr.Zero)//Keyboard hook
                {
                    uint id = Win32API.GetCurrentThreadId();
                    this.KeyboardProcDelegate = new Win32API.HookProc(this.KeyboardProc);
                    khook = Win32API.SetWindowsHookEx((IntPtr)HookHelper.WH_Codes.WH_KEYBOARD_LL, this.KeyboardProcDelegate, IntPtr.Zero, id);
                }
            }
            else
            {
                if (hHook == IntPtr.Zero)//Mouse hook
                {
                    uint id = Win32API.GetCurrentThreadId();
                    this.MouseHookProcedure = new Win32API.HookProc(this.MouseHookProc);
                    //Hanging festival hooks here
                    hHook = Win32API.SetWindowsHookEx((IntPtr)HookHelper.WH_Codes.WH_MOUSE_LL, MouseHookProcedure, IntPtr.Zero, id);
                }
            }
        }

        /// <summary>
        ///  Uninstall mouse hook
        /// </summary>
        public void UnInstallHook(HookHelper.HookType hookType)
        {
            bool isSuccess = false;
            if (hookType == HookHelper.HookType.KeyOperation)//Keyboard hook
            {
                if (khook != IntPtr.Zero)
                {
                    isSuccess = Win32API.UnhookWindowsHookEx(khook);
                    this.khook = IntPtr.Zero;
                }
            }
            else
            {
                if (this.hHook != IntPtr.Zero)//Mouse hook
                {
                    isSuccess = Win32API.UnhookWindowsHookEx(hHook);
                    this.hHook = IntPtr.Zero;
                }
            }
            if (isSuccess)
            {
               // MessageBox.Show("Successfully uninstalled!");
            }
            else
            {
                //MessageBox.Show("Uninstallation failed!");
            }
        }

        /// <summary>
        ///  Mouse hook processing function
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < (int)HookHelper.WH_Codes.HC_ACTION)
            {
                return Win32API.CallNextHookEx(hHook, nCode, wParam, lParam);
            }

            if (OnMouseActivity != null)
            {
                //Marshall the data from callback.
                HookHelper.MouseHookStruct mouseHookStruct = (HookHelper.MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(HookHelper.MouseHookStruct));
                MouseButtons button = MouseButtons.None;
                short mouseDelta = 0;
                switch ((int)wParam)
                {
                    case (int)HookHelper.WM_MOUSE.WM_LBUTTONDOWN:
                        //case WM_LBUTTONUP: 
                        //case WM_LBUTTONDBLCLK:
                        button = MouseButtons.Left;
                        break;
                    case (int)HookHelper.WM_MOUSE.WM_RBUTTONDOWN:
                        //case WM_RBUTTONUP: 
                        //case WM_RBUTTONDBLCLK: 
                        button = MouseButtons.Right;
                        break;
                    case (int)HookHelper.WM_MOUSE.WM_MOUSEWHEEL:
                        //button = MouseButtons.Middle;//Scroll wheel
                        //(value >> 16) & 0xffff; retrieves the high-order word from the given 32-bit value
                        mouseDelta = (short)((mouseHookStruct.MouseData >> 16) & 0xffff);
                        break;
                }

                int clickCount = 0;//Hits
                if (button != MouseButtons.None)
                {
                    if ((int)wParam == (int)HookHelper.WM_MOUSE.WM_LBUTTONDBLCLK || (int)wParam == (int)HookHelper.WM_MOUSE.WM_RBUTTONDBLCLK)
                    {
                        clickCount = 2;//Double click
                    }
                    else
                    {
                        clickCount = 1;//Click on
                    }
                }

                //Mouse events pass data
                MouseEventArgs e = new MouseEventArgs(button, clickCount, mouseHookStruct.Point.X, mouseHookStruct.Point.Y, mouseDelta);

                //Rewrite event
                OnMouseActivity(this, e);
            }

            return Win32API.CallNextHookEx(hHook, nCode, wParam, lParam);
        }

        /// <summary>
        ///  Keyboard hook processing function
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private int KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode < (int)HookHelper.WH_Codes.HC_ACTION)
                {
                    return Win32API.CallNextHookEx((IntPtr)khook, nCode, wParam, lParam);
                }

                HookHelper.KeyboardHookStruct keyHookStruct = (HookHelper.KeyboardHookStruct)Marshal.PtrToStructure(lParam, typeof(HookHelper.KeyboardHookStruct));

                #region
                //if ((int)wParam == (int)Keys.C && ((int)lParam & (int)Keys.ControlKey) != 0 ||//Ctrl+C
                //    (int)wParam == (int)Keys.X && ((int)lParam & (int)Keys.ControlKey) != 0)//Ctrl+V
                //{
                //    MessageBox.Show("C||V");
                //}

                //if (lParam.ToInt32() > 0)//Capture keyboard press
                //{
                //    Keys keys = (Keys)wParam;
                //        MessageBox.Show("Keyboard Press");
                //}
                //if (lParam.ToInt32() < 0)//Capture keyboard up
                //{
                //        MessageBox.Show("Keyboard up");
                //}
                /**************** 
                                   //Global keyboard hook to determine whether to press a key 
                                   wParam = = 0x100 // keyboard press 
                                   wParam = = 0x101 // keyboard up 
                ****************/
                //return 0;//If it returns 1, the message is ended, the message is expired, and no more delivery. If it returns 0 or calls the CallNextHookEx function, the message will continue to pass down the hook.
                #endregion

            }
            catch
            {

            }

            return Win32API.CallNextHookEx(khook, nCode, wParam, lParam);
        }
    }
}
