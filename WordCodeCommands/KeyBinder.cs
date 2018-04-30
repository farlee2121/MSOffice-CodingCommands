//using System;

//using System.Diagnostics;

//using System.Windows.Forms;

//using System.Runtime.InteropServices;

//using Microsoft.Office.Core;

//using Microsoft.Office.Interop.Word;

//using Word = Microsoft.Office.Interop.Word;



//namespace WordCodeCommands

//{

//    public class KeyBinder

//    {

//        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);



//        private static LowLevelKeyboardProc_proc = HookCallback;

 

//private static IntPtr _hookID = IntPtr.Zero;



//        privates tatic Microsoft.Office.Tools.CustomTaskPane ctpRef = null;



//        //Declare the mouse hook constant. //For other hook types, you can obtain these values from Winuser.h in the Microsoft SDK.

//        private const int WH_KEYBOARD = 2;

//        private const int HC_ACTION = 0;

//        public static void SetHook()

//        {

//            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());

//        }

//        public static void ReleaseHook()

//        {

//            UnhookWindowsHookEx(_hookID);

//        }

//        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)

//        {

//            int PreviousStateBit = 31;

//            bool KeyWasAlreadyPressed = false;

//            Int64 bitmask = (Int64)Math.Pow(2, (PreviousStateBit - 1));

//            if (nCode < 0)

//            {

//                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);

//            }

//            else

//            {

//                if (nCode == HC_ACTION)

//                {

//                    Keys keyData = (Keys)wParam;

//                    KeyWasAlreadyPressed = ((Int64)lParam & bitmask) > 0;

//                    if (Functions.IsKeyDown(Keys.ControlKey) && keyData == Keys.D1 && KeyWasAlreadyPressed == false)

//                    {

//                        object missing = System.Reflection.Missing.Value;

//                        Word.

//                        Workbook exBook = Excel_CustomTaskPane_ACC.Globals.ThisAddIn.Application.ActiveWorkbook;

//                        Word.

//                        Worksheet exSheet = (Word.Worksheet)exBook.ActiveSheet;

//                        exSheet.get_Range(

//                        "A4", missing).Value2 = "A";

//                    }

//                }

//                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);

//            }

//        }



//        [

//        DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]



//        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

//        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]

//        [return: MarshalAs(UnmanagedType.Bool)]

//        private static extern bool UnhookWindowsHookEx(IntPtrhhk);

//        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]

//        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtrlParam);

//    }

//    public classFunctions

//{

//public static bool IsKeyDown(Keyskeys)

//    {

//        return (GetKeyState((int)keys) & 0x8000) == 0x8000;

//    }

//    [DllImport("user32.dll")]

//    static extern short GetKeyState(intnVirtKey);

//}

//}
