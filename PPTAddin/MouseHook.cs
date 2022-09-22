using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PPTAddin
{
    public class MouseHook
    {
        #region Define variables
        //Define hook processing function
        private Win32API.HookProc MouseHookProcedure;
        //Define hook handle
        private IntPtr hHook = IntPtr.Zero;
        //Define mouse events
        public event MouseEventHandler OnMouseActivity;
        #endregion

        /// <summary>
        ///  Install the mouse hook
        /// </summary>
        public void InstallHook()
        {
            if (hHook == IntPtr.Zero)
            {
                uint id = Win32API.GetCurrentThreadId();
                this.MouseHookProcedure = new Win32API.HookProc(this.MouseHookProc);
                //Hanging festival hooks here
                hHook = Win32API.SetWindowsHookEx((IntPtr)HookHelper.WH_Codes.WH_MOUSE_LL, MouseHookProcedure, IntPtr.Zero, id);
            }
        }

        /// <summary>
        ///  Uninstall mouse hook
        /// </summary>
        public void UnInstallHook()
        {
            bool isSuccess = false;
            if (this.hHook != IntPtr.Zero)
            {
                isSuccess = Win32API.UnhookWindowsHookEx(hHook);
                this.hHook = IntPtr.Zero;
            }
            if (isSuccess)
            {
                //MessageBox.Show("Successfully uninstalled!");
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
                        button = MouseButtons.XButton1;
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
                    case (int)HookHelper.WM_MOUSE.WM_LBUTTONCLK:
                        button = MouseButtons.Left;
                        break;
                    case (int)HookHelper.WM_MOUSE.WM_LBUTTONUP:
                        button = MouseButtons.XButton2;
                        break;
                }

                int clickCount = 0;//Hits
                if (button != MouseButtons.None)
                {
                    if (wParam == (IntPtr)HookHelper.WM_MOUSE.WM_LBUTTONDBLCLK || wParam == (IntPtr)HookHelper.WM_MOUSE.WM_RBUTTONDBLCLK)
                    {
                        clickCount = 2;//Double click
                    }
                    else if((wParam == (IntPtr)HookHelper.WM_MOUSE.WM_LBUTTONCLK))
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
    }
}
