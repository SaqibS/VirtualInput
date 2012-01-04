namespace VirtualInput
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    /// <summary>Enables applications to intercept keystrokes</summary>
    public static class VirtualKeyboard
    {
        private static int hHook = 0;
        private static NativeMethods.HookProc hookProc;

        /// <summary>Fires when a key is depressed on the keyboard</summary>
        public static event KeyEventHandler KeyDown;
        /// <summary>Fires when a key is pressed and released on the keyboard</summary>
        public static event KeyPressEventHandler KeyPress;
        /// <summary>Fires when a key is released on the keyboard</summary>
        public static event KeyEventHandler KeyUp;

        /// <summary>Starts intercepting keystrokes</summary>
        public static void StartInterceptor()
        {
            if (hHook == 0)
            {
                hookProc = new NativeMethods.HookProc(KeyboardHookProc);
                hHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, hookProc, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
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

        /// <summary>Called by the Win32 functions when a keyboard event occurs</summary>
        /// <param name="nCode">A code the hook procedure uses to determine how to process the message. If code is less than zero, the hook procedure must pass the message to the CallNextHookEx function without further processing and should return the value returned by CallNextHookEx.</param>
        /// <param name="wParam">The virtual-key code of the key that generated the keystroke message.</param>
        /// <param name="lParam">The repeat count, scan code, extended-key flag, context code, previous key-state flag, and transition-state flag.</param>
        /// <returns>1 if the event was handled, otherwise the return value of CallNextHookEx.</returns>
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

            return handled ? 1 : NativeMethods.CallNextHookEx(hHook, nCode, wParam, lParam);
        }
    }
}
