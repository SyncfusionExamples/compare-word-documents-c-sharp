using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.IO;

namespace Compare_Word_documents
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream. 
            using (FileStream originalDocumentFileStream = new FileStream(Path.GetFullPath(@"../../../Data/OriginalDocument.docx"), FileMode.Open))
            {
                //Load the original Word document.
                using (WordDocument originalDocument = new WordDocument(originalDocumentFileStream, FormatType.Docx))
                {
                    //Open the file as Stream. 
                    using (FileStream revisedDocumentFileStream = new FileStream(Path.GetFullPath(@"../../../Data/RevisedDocument.docx"), FileMode.Open))
                    {
                        //Load the revised Word document.
                        using (WordDocument revisedDocument = new WordDocument(revisedDocumentFileStream, FormatType.Docx))
                        {
                            //Compares the original document with revised document by detecting the formatting changes
                            originalDocument.Compare(revisedDocument, "Nancy Davolio", DateTime.Now.AddDays(-1));
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
