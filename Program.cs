using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace powerpoint_interop
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Current directory: " + Directory.GetCurrentDirectory());
            var pptx = "./canva poster result.pptx";
            var pdf = "result.pdf";

            if (args.Length < 2)
            {
                Console.WriteLine("Not enough arguments were given. Using default files");
            } else
            {
                pptx = args[0];
                pdf = args[1];
            }
            pptx = Path.GetFullPath(pptx);
            Console.WriteLine($"PPTX: {pptx}\nPDF: {pdf}");
            PPTXToPDF(pptx, pdf);
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
        public static void PPTXToPDF(string originalPptPath, string pdfPath)
        {
            // Create COM Objects
            Application pptApplication = null;
            Presentation pptPresentation = null;
            try
            {
                object unknownType = Type.Missing;

                //start power point
                pptApplication = new Application();

                //open powerpoint document
                pptPresentation = pptApplication.Presentations.Open(originalPptPath,
                    MsoTriState.msoTrue, MsoTriState.msoTrue,
                    MsoTriState.msoFalse);

                // save PowerPoint as PDF
                pptPresentation.ExportAsFixedFormat(pdfPath,
                    PpFixedFormatType.ppFixedFormatTypePDF,
                    PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null,
                    PpPrintRangeType.ppPrintAll, string.Empty, true, true, true,
                    true, false, unknownType);
            } catch (Exception e)
            {
                Console.WriteLine("While converting to pdf, encountered an error");
                Console.WriteLine(e.ToString());
            }
            finally
            {
                // Close and release the Document object.
                if (pptPresentation != null)
                {
                    pptPresentation.Close();
                    pptPresentation = null;
                }

                // Quit PowerPoint and release the ApplicationClass object.
                if (pptApplication != null)
                {
                    pptApplication.Quit();
                    pptApplication = null;
                }
            }
        }
    }
}
