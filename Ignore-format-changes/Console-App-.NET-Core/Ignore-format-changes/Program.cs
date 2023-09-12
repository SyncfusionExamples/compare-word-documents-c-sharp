using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.IO;

namespace Ignore_format_changes
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream. 
            using (FileStream originalDocumentFileStream = new FileStream(Path.GetFullPath(@"../../../Data/OriginalDocument.docx"), FileMode.Open))
            {
                //Loads the original Word document.
                using (WordDocument originalDocument = new WordDocument(originalDocumentFileStream, FormatType.Docx))
                {
                    //Open the file as Stream. 
                    using (FileStream revisedDocumentFileStream = new FileStream(Path.GetFullPath(@"../../../Data/RevisedDocument.docx"), FileMode.Open))
                    {
                        //Loads the revised Word document.
                        using (WordDocument revisedDocument = new WordDocument(revisedDocumentFileStream, FormatType.Docx))
                        {
                            //Disable the flag to ignore the formatting changes while comparing the documents
                            ComparisonOptions comparisonOptions = new ComparisonOptions();
                            comparisonOptions.DetectFormatChanges = false;
                            originalDocument.Compare(revisedDocument, "Nancy Davolio", DateTime.Now.AddDays(-1), comparisonOptions);
                            //Create the output file stream.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Output.docx")))
                            {
                                //Save the document.
                                originalDocument.Save(fileStreamOutput, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
