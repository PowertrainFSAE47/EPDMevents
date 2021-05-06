using EPDM.Interop.epdm;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EpdmEvents
{

    [Guid("F3FF19F2-2652-4241-8D20-3DFB395C45E1"), ComVisible(true)]
    public class HookSetup : IEdmAddIn5
    {
        private EdmVault5 thisVault;
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //MessageBox.Show(String.Format(" Attach debugger to process {0} (ID: {1})", System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().Id));
            try
            {
                //Información miscelanea del add-in
                poInfo.mbsAddInName = "EpdmEvents";
                poInfo.mbsCompany = "ADDR";
                poInfo.mbsDescription = string.Format("Testing de HOOKS");
                poInfo.mlAddInVersion = 1;
                poInfo.mlRequiredVersionMajor = 5;
                poInfo.mlRequiredVersionMinor = 2;

                // Instancia del almacen.
                thisVault = (EdmVault5)poVault;

                //Creación de hooks. Un hook es una suscripción a un evento de pdm.
                //Los eventos están definidos en el enum EdmCmdType

                // Hook para botones de tarjeta.
                poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton);

                // Hook para un comando, usado en testing
                poCmdMgr.AddCmd(1, "C# Add-in", (int)EdmMenuFlags.EdmMenu_Nothing);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("(GetAddInInfo) HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            {
                if (poCmd.mlCmdID == 1)
                {
                    System.Windows.Forms.MessageBox.Show("C# Add-in");
                }
            }
        }
    }
}
