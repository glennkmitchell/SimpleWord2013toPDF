using System;
using Microsoft.Office.Interop.Word;

namespace WordToPdf
{
    internal class Converter
    {

        private readonly string _outputFile;
        private object _inputFile;


        public Converter(string inputFile, string outputFile)
        {
            _inputFile = inputFile;
            _outputFile = outputFile;
        }

        public bool Convert()
        {
            // create the document and the word application objects 
            var wordApp = new Application();

            try
            {
                // create missing type for unused values
                var missing = Type.Missing;

                // ensure no changes occur to source
                object readOnly = true;

                // do not hide the file
                object isVisible = true;

                // open the document
                var document = wordApp.Documents.Open(
                    ref _inputFile,
                    ref missing,
                    ref readOnly,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref isVisible,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing);

                // now export the document
                if (document != null)
                {
                    document.ExportAsFixedFormat(_outputFile, WdExportFormat.wdExportFormatPDF, false,
                                                 WdExportOptimizeFor.wdExportOptimizeForPrint,
                                                 WdExportRange.wdExportAllDocument, 0, 0,
                                                 WdExportItem.wdExportDocumentContent, true, true,
                                                 WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false,
                                                 ref missing);
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}