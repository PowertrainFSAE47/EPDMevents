using System;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;


namespace EpdmEvents
{
    public class errorHandler
    {
        IWin32Window pWdw;
        public errorHandler(IWin32Window ParentWnd)
        {
            this.pWdw = ParentWnd;
        }
        public enum ErrorMsgs
        {
            genericMsg
        }

        public void throwMessage(Enum caso, string msgTail = "")
        {
            switch (caso)
            {
                case ErrorMsgs.genericMsg:
                    MessageBox.Show(pWdw, string.Format("Mensaje genérico\n\n{0}", msgTail), "Solidworks PDM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                default:
                    break;
            }
        }
    }
}
