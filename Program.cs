using System;
using Microsoft.Office.Interop.Visio;

namespace VisioSaveAs
{
  class Program
  {
    static void Main(string[] args)
    {
      if (args.Length != 1) Environment.Exit(0);

      string path = Environment.CurrentDirectory + "\\";
      
      InvisibleAppClass app = new InvisibleAppClass();
      Document doc = app.Documents.Open(path + args[0]);
      doc.Pages[1].Shapes[1].Export(path + "temp.svg");
      app.ActiveDocument.Close();
      doc = app.Documents.Open(path + "temp.svg");
      VisFixedFormatTypes pdfType = VisFixedFormatTypes.visFixedFormatPDF;
      doc.ExportAsFixedFormat(pdfType, path + args[0].Split('.')[0] + ".pdf",
        Microsoft.Office.Interop.Visio.VisDocExIntent.visDocExIntentScreen,
        Microsoft.Office.Interop.Visio.VisPrintOutRange.visPrintAll);
      app.ActiveDocument.Close();
      app.Quit();
    }
  }
}
