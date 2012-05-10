using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace SpamGrabberControl
{
    public static class UIControl
    {
        private const int WM_SETREDRAW = 0x000B;

        private const int WM_USER = 0x400;

        private const int EM_GETEVENTMASK = (WM_USER + 59);

        private const int EM_SETEVENTMASK = (WM_USER + 69);


        public static void StopWindowUpdating(IntPtr Handle)
        {

            IntPtr eventMask = IntPtr.Zero;

            try
            {

                // Stop redrawing:

                SendMessage(Handle, WM_SETREDRAW, 0, IntPtr.Zero);

                // Stop sending of events:

                eventMask = SendMessage(Handle, EM_GETEVENTMASK, 0, IntPtr.Zero);



                // change colors and stuff in the RichTextBox

            }

            finally
            {

                // turn on events

                SendMessage(Handle, EM_SETEVENTMASK, 0, eventMask);

                // turn on redrawing

                SendMessage(Handle, WM_SETREDRAW, 1, IntPtr.Zero);

            }



        }


        public static void StartWindowUpdating(IntPtr Handle)
        {

            IntPtr eventMask = IntPtr.Zero;

            try
            {

                // Start redrawing:

                SendMessage(Handle, WM_SETREDRAW, 1, IntPtr.Zero);

                // Start sending of events:

                eventMask = SendMessage(Handle, EM_GETEVENTMASK, 0, IntPtr.Zero);



                // change colors and stuff in the RichTextBox

            }

            finally
            {

                // turn on events

                SendMessage(Handle, EM_SETEVENTMASK, 0, eventMask);

                // turn on redrawing

                SendMessage(Handle, WM_SETREDRAW, 1, IntPtr.Zero);

            }



        }


        [DllImport("user32", CharSet = CharSet.Auto)]

        private extern static IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);


    }
}
