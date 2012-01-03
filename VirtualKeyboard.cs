namespace VirtualInput
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public static class VirtualKeyboard
    {
        private static int hKeyboardHook = 0;
        private static NativeMethods.HookProc KeyboardHookProcedure;

        public static event KeyEventHandler KeyDown;
        public static event KeyPressEventHandler KeyPress;
        public static event KeyEventHandler KeyUp;

        public static void StartInterceptor()
        {
            if (hKeyboardHook == 0)
            {
                KeyboardHookProcedure = new NativeMethods.HookProc(KeyboardHookProc);
                hKeyboardHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
                if (hKeyboardHook == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    StopInterceptor();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        public static void StopInterceptor()
        {
            if (hKeyboardHook != 0)
            {
                int result = NativeMethods.UnhookWindowsHookEx(hKeyboardHook);
                hKeyboardHook = 0;
                if (result == 0)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new Win32Exception(errorCode);
                }
            }
        }

        private static int KeyboardHookProc(int nCode, Int32 wParam, IntPtr lParam)
        {
            bool handled = false;
            if ((nCode >= 0) && (KeyDown != null || KeyUp != null || KeyPress != null))
            {
                NativeMethods.KeyboardHookStruct MyKeyboardHookStruct = (NativeMethods.KeyboardHookStruct)Marshal.PtrToStructure(lParam, typeof(NativeMethods.KeyboardHookStruct));

                if (KeyDown != null && (wParam == NativeMethods.WM_KEYDOWN || wParam == NativeMethods.WM_SYSKEYDOWN))
                {
                    Keys keyData = (Keys)MyKeyboardHookStruct.vkCode;
                    KeyEventArgs e = new KeyEventArgs(keyData);
                    KeyDown(null, e);
                    handled = handled || e.Handled;
                }

                if (KeyPress != null && wParam == NativeMethods.WM_KEYDOWN)
                {
                    bool isDownShift = (NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) & 0x80) == 0;
                    bool isDownCapslock = NativeMethods.GetKeyState(NativeMethods.VK_CAPITAL) != 0;

                    byte[] keyState = new byte[256];
                    NativeMethods.GetKeyboardState(keyState);
                    byte[] inBuffer = new byte[2];
                    if (NativeMethods.ToAscii(MyKeyboardHookStruct.vkCode, MyKeyboardHookStruct.scanCode, keyState, inBuffer, MyKeyboardHookStruct.flags) == 1)
                    {
                        char key = (char)inBuffer[0];
                        if ((isDownCapslock ^ isDownShift) && Char.IsLetter(key))
                        {
                            key = Char.ToUpper(key);
                        }

                        KeyPressEventArgs e = new KeyPressEventArgs(key);
                        KeyPress(null, e);
                        handled = handled || e.Handled;
                    }
                }

                if (KeyUp != null && (wParam == NativeMethods.WM_KEYUP || wParam == NativeMethods.WM_SYSKEYUP))
                {
                    Keys keyData = (Keys)MyKeyboardHookStruct.vkCode;
                    KeyEventArgs e = new KeyEventArgs(keyData);
                    KeyUp(null, e);
                    handled = handled || e.Handled;
                }
            }

            return handled ? 1 : NativeMethods.CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
        }
    }
}
