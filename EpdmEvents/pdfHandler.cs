using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using System;
using System.Windows.Forms;

namespace EpdmEvents
{
    public class pdfHandler
    {

        private errorHandler err;
        public pdfHandler(errorHandler err)
        {
            this.err = err;
        }

        Document documento;

        public void crearPdf(string pdfPath,string paragraph)
        {
            try
            {
                // Instancia de PDFWriter, escribe el archivo fisico en disco del pdf.
                PdfWriter pdfWriter = new PdfWriter(pdfPath);

                MessageBox.Show(pdfWriter.CanWrite.ToString());
                //Instancia de pdfDocument. Este es el administrador del documento, organiza lo escrito
                PdfDocument pdf = new PdfDocument(pdfWriter);

                //Document es la abstracción del documento. Sobre esto se trabaja
                documento = new Document(pdf, PageSize.LETTER);

                documento.Add(new Paragraph(paragraph));
                documento.Close();
            }
            catch (Exception ex)
            {
                err.throwMessage(errorHandler.ErrorMsgs.genericMsg, ex.Message);
                err.throwMessage(errorHandler.ErrorMsgs.genericMsg, ex.StackTrace);
                throw;
            }
            finally
            {
                documento.Close();
            }

        }
    }
}
