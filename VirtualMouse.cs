namespace VirtualInput
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    /// <summary>Enables applications to intercept mouse activity</summary>
    public static class VirtualMouse
    {
        private static int hHook = 0;
        private static NativeMethods.HookProc hookProc;

        /// <summary>Fires when the mouse is moved/clicked/scrolled</summary>
        public static event MouseEventHandler MouseActivity;

        /// <summary>Starts intercepting keystrokes</summary>
        public static void StartInterceptor()
        {
            if (hHook == 0)
            {
                hookProc = new NativeMethods.HookProc(MouseHookProc);
                IntPtr moduleHandle = NativeMethods.GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName);
                hHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, hookProc, moduleHandle, 0);
                if (hHook == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    StopInterceptor();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        /// <summary>Stops intercepting keystrokes</summary>
        public static void StopInterceptor()
        {
            if (hHook != 0)
            {
                int result = NativeMethods.UnhookWindowsHookEx(hHook);
                hHook = 0;
                if (result == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        /// <summary>Called by the Win32 functions when a mouse event occurs</summary>
        /// <param name="nCode">A code that the hook procedure uses to determine how to process the message. If nCode is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx.</param>
        /// <param name="wParam">The identifier of the mouse message.</param>
        /// <param name="lParam">A pointer to a MOUSEHOOKSTRUCT structure.</param>
        /// <returns>1 if the event was handled, otherwise the return value of CallNextHookEx.</returns>
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

            return NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam);
        }
    }
}
