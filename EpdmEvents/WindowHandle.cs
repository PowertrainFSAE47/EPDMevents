using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace EpdmEvents
{
    class WindowHandle :IWin32Window
    {
        // Esto es un wrapper para manejar ventanas emergentes de mensajes, asi no se van al background.
        private IntPtr mHwnd;

        public WindowHandle(int hWnd)
        {
            mHwnd = new IntPtr(hWnd);
        }
        public IntPtr Handle
        {
            get { return mHwnd; }
        }
    }

}
