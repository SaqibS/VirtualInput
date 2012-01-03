namespace VirtualInput
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public static class VirtualMouse
    {
        private static int hMouseHook = 0;
        private static NativeMethods.HookProc MouseHookProcedure;

        public static event MouseEventHandler MouseActivity;

        public static void StartInterceptor()
        {
            if (hMouseHook == 0)
            {
                MouseHookProcedure = new NativeMethods.HookProc(MouseHookProc);
                hMouseHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, MouseHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
                if (hMouseHook == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    StopInterceptor();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        public static void StopInterceptor()
        {
            if (hMouseHook != 0)
            {
                int result = NativeMethods.UnhookWindowsHookEx(hMouseHook);
                hMouseHook = 0;
                if (result == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        private static int MouseHookProc(int nCode, int wParam, IntPtr lParam)
        {
            if ((nCode >= 0) && (MouseActivity != null))
            {
                NativeMethods.MouseLLHookStruct mouseHookStruct = (NativeMethods.MouseLLHookStruct)Marshal.PtrToStructure(lParam, typeof(NativeMethods.MouseLLHookStruct));
                MouseButtons button = MouseButtons.None;
                short mouseDelta = 0;
                switch (wParam)
                {
                    case NativeMethods.WM_LBUTTONDOWN:
                        button = MouseButtons.Left;
                        break;
                    case NativeMethods.WM_RBUTTONDOWN:
                        button = MouseButtons.Right;
                        break;
                    case NativeMethods.WM_MOUSEWHEEL:
                        mouseDelta = (short)((mouseHookStruct.mouseData >> 16) & 0xffff);
                        break;
                }

                int clickCount = 0;
                if (button != MouseButtons.None)
                {
                    if (wParam == NativeMethods.WM_LBUTTONDBLCLK || wParam == NativeMethods.WM_RBUTTONDBLCLK)
                    {
                        clickCount = 2;
                    }
                    else
                    {
                        clickCount = 1;
                    }
                }

                MouseEventArgs e = new MouseEventArgs(button, clickCount, mouseHookStruct.pt.x, mouseHookStruct.pt.y, mouseDelta);
                MouseActivity(null, e);
            }

            return NativeMethods.CallNextHookEx(hMouseHook, nCode, wParam, lParam);
        }
    }
}
