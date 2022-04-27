using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using drawing = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;
using Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace CreateImageUsingOpenXml
{
    class FIleGeneration
    {
        public void GenerateFile()
        {
            // Create a document by path, and write some text in it.  
            string fileName = @"Test.docx";
            string txt = "Hello, world!";

            // Create the Word document.   
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Create(fileName,
                    WordprocessingDocumentType.Document);

            // Create a MainDocumentPart instance.  
            MainDocumentPart mainDocumentPart =
                wordprocessingDocument.AddMainDocumentPart();
            mainDocumentPart.Document = new Document();

            // Create a Document instance.  
            Document document = mainDocumentPart.Document;

            // Create a Body instance.  
            Body body = document.AppendChild(new Body());

            // Create a Paragraph instance.  
            Paragraph para = body.AppendChild(new Paragraph());

            // Create a Run instance.  
            Run run = para.AppendChild(new Run());

            // Add Text to the Run element.  
            run.AppendChild(new Text(txt));

            // Close the document handle  
            wordprocessingDocument.Close();

            Console.WriteLine("The document has been created. Press any key.");
            Console.ReadKey();
        }

        public void Generate()
        {
            string filepath = @"C:\Users\mukes\source\repos\CreateImageUsingOpenXml\CreateImageUsingOpenXml\bin\Debug\Test.docx";
            string ImgfileName = @"C:\Users\mukes\source\repos\CreateImageUsingOpenXml\CreateImageUsingOpenXml\bin\Debug\sign.png";
            InsertAPicture(filepath, ImgfileName);
            Console.WriteLine("Image added");
        }
        public static void InsertAPicture(string document, string fileName)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart),fileName);
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId,string imgFile)
        {
            // Define the reference of the image.
            Int64Value iWidth = 990000L;
            Int64Value iHeight = 792000L;
            DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            using (Bitmap bm = new Bitmap(imgFile))
            {
                iWidth = bm.Width;
                iHeight = bm.Height;

                extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
            }
            // iWidth = (Int64Value)Math.Round((decimal)iWidth * 5025);
            // iHeight = (Int64Value)Math.Round((decimal)iHeight * 3025);

            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = extents.Cx, Cy = extents.Cy },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                         new drawing.Graphic(
                             new drawing.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new drawing.Blip(
                                             new drawing.BlipExtensionList(
                                                 new drawing.BlipExtension()
                                                 {
                                                     Uri = new Guid().ToString()
                                                        
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             drawing.BlipCompressionValues.Print
                                         },
                                         new drawing.Stretch(
                                             new drawing.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new drawing.Transform2D(
                                             new drawing.Offset() { X = 0L, Y = 0L },
                                             new drawing.Extents() { Cx = extents.Cx, Cy = extents.Cy }),
                                         new drawing.PresetGeometry(
                                             new drawing.AdjustValueList()
                                         )
                                         { Preset = drawing.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 800U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                         //EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }



    }
}
