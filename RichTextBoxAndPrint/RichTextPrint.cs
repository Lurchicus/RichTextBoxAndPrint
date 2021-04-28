using System;
using System.Runtime.InteropServices;
using System.Drawing.Printing;

namespace RichTextBoxAndPrint
{
    /// <summary>
    /// RichTextPrint - Class extension of the .NET RichTextBox that provides 
    /// printing
    /// </summary>
    public partial class RichTextPrint : System.Windows.Forms.RichTextBox
    {
        private const double anInch = 14.4;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public Int32 Left;
            public Int32 Top;
            public Int32 Right;
            public Int32 Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CHARRANGE
        {
            public Int32 cpMin;
            public Int32 cpMax;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FORMATRANGE
        {
            public IntPtr hdc;
            public IntPtr hdcTarget;
            public RECT rc;
            public RECT rcPage;
            public CHARRANGE chrg;
        }

        private const Int32 WM_USER = 0x400;
        private const Int32 EM_FORMATRANGE = WM_USER + 57;

        [DllImport("USER32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, Int32 msg, IntPtr wp, IntPtr lp);
        
        /// <summary>
        /// New Print method
        /// </summary>
        /// <param name="charFrom"></param>
        /// <param name="charTo"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public Int32 Print(Int32 charFrom, Int32 charTo, PrintPageEventArgs e)
        {
            //Mark starting and ending character
            CHARRANGE cRange;
            cRange.cpMin = charFrom;
            cRange.cpMax = charTo;

            //Calculate the area to render and print
            RECT rectToPrint;
            rectToPrint.Top = (Int32)(e.MarginBounds.Top * anInch);
            rectToPrint.Bottom = (Int32)(e.MarginBounds.Bottom * anInch);
            rectToPrint.Left = (Int32)(e.MarginBounds.Left * anInch);
            rectToPrint.Right = (Int32)(e.MarginBounds.Right * anInch);

            //Calculate the size of the page
            RECT rectPage;
            rectPage.Top = (Int32)(e.PageBounds.Top * anInch);
            rectPage.Bottom = (Int32)(e.PageBounds.Bottom * anInch);
            rectPage.Left = (Int32)(e.PageBounds.Left * anInch);
            rectPage.Right = (Int32)(e.PageBounds.Right * anInch);

            IntPtr hdc = e.Graphics.GetHdc();

            FORMATRANGE fmtRange;
            fmtRange.chrg = cRange;
            fmtRange.hdc = hdc;
            fmtRange.hdcTarget = hdc;
            fmtRange.rc = rectToPrint;
            fmtRange.rcPage = rectPage;
            IntPtr wparam = new IntPtr(1);
            IntPtr lparam = Marshal.AllocCoTaskMem(Marshal.SizeOf(fmtRange));
            Marshal.StructureToPtr(fmtRange, lparam, false);
            IntPtr res = SendMessage(Handle, EM_FORMATRANGE, wparam, lparam);

            Marshal.FreeCoTaskMem(lparam);

            e.Graphics.ReleaseHdc(hdc);

            return res.ToInt32();
        }
    }
}
