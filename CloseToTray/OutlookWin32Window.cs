using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CloseToTray
{
    public class OutlookWin32Window : NativeWindow, IDisposable
    {
        private static uint WM_CLOSE = 0x10;


        public OutlookWin32Window(IntPtr handle)
        {
            this.AssignHandle(handle);
        }

        public OutlookWin32Window(Microsoft.Office.Interop.Outlook.Explorer explorer)
        {
            IntPtr handle = Native.FindWindow("rctrl_renwnd32", explorer.Caption);
            this.AssignHandle(handle);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLOSE)
            {
                CancelEventArgs args = new CancelEventArgs();
                Closing?.Invoke(this, args);
                if (args.Cancel)
                    return;
            }

            base.WndProc(ref m);
        }

        public void Dispose()
        {
            this.ReleaseHandle();
        }

        public event EventHandler<CancelEventArgs> Closing;

        public class CancelEventArgs : EventArgs
        {
            public bool Cancel { get; set; }
        }
    }
}
