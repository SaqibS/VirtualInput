namespace VirtualInput
{
    using System;
    using System.Runtime.InteropServices;

    internal static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        internal class POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class MouseHookStruct
        {
            public POINT pt;
            public int hwnd;
            public int wHitTestCode;
            public int dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class MouseLLHookStruct
        {
            public POINT pt;
            public int mouseData;
            public int flags;
            public int time;
            public int dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class KeyboardHookStruct
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public int dwExtraInfo;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        internal static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, int dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        internal static extern int UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        internal static extern int CallNextHookEx(int idHook, int nCode, int wParam, IntPtr lParam);

        internal delegate int HookProc(int nCode, int wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern int ToAscii(int uVirtKey, int uScanCode, byte[] lpbKeyState, byte[] lpwTransKey, int fuState);

        [DllImport("user32.dll")]
        internal static extern int GetKeyboardState(byte[] pbKeyState);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        internal static extern short GetKeyState(int vKey);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetModuleHandle(string name);

        internal const int WH_MOUSE_LL = 14;
        internal const int WH_KEYBOARD_LL = 13;

        internal const int WH_MOUSE = 7;
        internal const int WH_KEYBOARD = 2;

        internal const int WM_MOUSEMOVE = 0x200;
        internal const int WM_LBUTTONDOWN = 0x201;
        internal const int WM_RBUTTONDOWN = 0x204;
        internal const int WM_MBUTTONDOWN = 0x207;
        internal const int WM_LBUTTONUP = 0x202;
        internal const int WM_RBUTTONUP = 0x205;
        internal const int WM_MBUTTONUP = 0x208;
        internal const int WM_LBUTTONDBLCLK = 0x203;
        internal const int WM_RBUTTONDBLCLK = 0x206;
        internal const int WM_MBUTTONDBLCLK = 0x209;
        internal const int WM_MOUSEWHEEL = 0x020A;

        internal const int WM_KEYDOWN = 0x100;
        internal const int WM_KEYUP = 0x101;
        internal const int WM_SYSKEYDOWN = 0x104;
        internal const int WM_SYSKEYUP = 0x105;

        internal const byte VK_SHIFT = 0x10;
        internal const byte VK_CAPITAL = 0x14;
        internal const byte VK_NUMLOCK = 0x90;
    }
}
