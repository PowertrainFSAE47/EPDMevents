using EPDM.Interop.epdm;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.IO.Compression;

namespace EpdmEvents
{

    [Guid("F3FF19F2-2652-4241-8D20-3DFB395C45E1"), ComVisible(true)]
    public class HookSetup : IEdmAddIn5
    {
        private EdmVault5 thisVault;
        private IEdmVault8 thisVaultV8;
        private IWin32Window parentWnd;
        private int fileId;
        private int folderId;
        private int parentFolderId;
        private string fileName;
        private string filePath;
        private string fileDir;


        public void LoadProblematicAssemblies()
        {
            var thisAssembly = new FileInfo(this.GetType().Assembly.Location);
            var location = thisAssembly.Directory.FullName;

            // extract references
            ZipFile.ExtractToDirectory(System.IO.Path.Combine(location, $@"com.zip"), location);

            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.kernel.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\bouncycastle.crypto.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\common.logging.core.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\common.logging.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.barcodes.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.io.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.layout.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.pdfa.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.sign.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.styledxmlparser.dll"));
            System.Reflection.Assembly.LoadFrom(System.IO.Path.Combine(location, $@"com\itext.svg.dll"));

        }

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //MessageBox.Show(String.Format(" Attach debugger to process {0} (ID: {1})", System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().Id));
            try
            {
                //Información miscelanea del add-in
                poInfo.mbsAddInName = "EpdmEvents";
                poInfo.mbsCompany = "ADDR";
                poInfo.mbsDescription = string.Format("Manipulacion de PDF");
                poInfo.mlAddInVersion = 13;
                poInfo.mlRequiredVersionMajor = 5;
                poInfo.mlRequiredVersionMinor = 2;

                // Instancia del almacen.
                
                
                //Creación de hooks. Un hook es una suscripción a un evento de pdm.
                //Los eventos están definidos en el enum EdmCmdType

                // Hook para un comando, usado en testing
                poCmdMgr.AddCmd(1, "PDF MERGE", (int)EdmMenuFlags.EdmMenu_Nothing);


                LoadProblematicAssemblies();

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
            enableDebugger();

            //Holder de ppoData
            EdmCmdData[] fileData = (EdmCmdData[])ppoData;

            //Almacen
            thisVault = (EdmVault5)poCmd.mpoVault;
            thisVaultV8 = (IEdmVault8)thisVault;

            //parent windows
            parentWnd = (IWin32Window)thisVaultV8.GetWin32Window(poCmd.mlParentWnd);

            //Error handler
            errorHandler err = new errorHandler(parentWnd);

            //pdfHandler
            pdfHandler pdfDocHdlr = new pdfHandler(err);

            //This file
            IEdmFile5 thisFile;

            if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            {
                try
                {
                    if (poCmd.mlCmdID == 1)
                    {
                        foreach (EdmCmdData selectedItem in fileData)
                        {
                            //File name and id. Parent folder ID.
                            fileName = selectedItem.mbsStrData1;
                            fileId = selectedItem.mlObjectID1;
                            parentFolderId = selectedItem.mlObjectID3;

                            //Get the file.
                            thisFile =(IEdmFile5)thisVault.GetObject(EdmObjectType.EdmObject_File, fileId);

                            //Get local directory for file
                            fileDir = Path.GetDirectoryName((thisFile.GetLocalPath(parentFolderId)));

                            //Shave extension off fileName
                            fileName = Path.GetFileNameWithoutExtension(fileName);

                            //the PDF Path is the concatenation of the filedir, the file name and pdf extension.
                            string pdfPath = string.Format("{0}\\{1}.pdf", fileDir, fileName);

                            err.throwMessage(errorHandler.ErrorMsgs.genericMsg, pdfPath);
                            //Creating a testing PDF
                            pdfDocHdlr.crearPdf(pdfPath, "This is a PDF generation test!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    err.throwMessage(errorHandler.ErrorMsgs.genericMsg, ex.Message);
                    throw;
                }

            }

        }

        private void enableDebugger()
        {
            MessageBox.Show(String.Format(" Attach debugger to process {0} (ID: {1})", System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().Id));
        }
    }
}
