using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml.Linq;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using OpenXmlPowerTools;
using PdfSharp.Pdf;
using TheArtOfDev.HtmlRenderer.PdfSharp;
using ColorMode = DinkToPdf.ColorMode;
//using OpenXmlPowerTools;


namespace CreateImageUsingOpenXml
{
    class ConvertHtmlToPDF
    {
        string filePath;
        static string DocxConvertedToHtmlDirectory = "DocxConvertedToHtml/";

        public ConvertHtmlToPDF(string filePath)
        {
            this.filePath = filePath;
        }
        public void ModifyDocumentFIle()
        {
            byte[] byteArray = File.ReadAllBytes(filePath);
            if (byteArray != null)
            {
                try
                {

                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        memoryStream.Write(byteArray, 0, byteArray.Length);
                        using (WordprocessingDocument wDoc =
                            WordprocessingDocument.Open(memoryStream, true))
                        {
                            var body = wDoc.MainDocumentPart.Document.Body;
                            var lastPara = body.Elements<Paragraph>().LastOrDefault();
                            var newPara = new Paragraph(
                                new Run(
                                    new Text("Hello World!")));
                            lastPara.InsertAfterSelf(newPara);
                        }
                       // Session["ByteArray"] = memoryStream.ToArray();
                       // lblMessage.Text = "Paragraph added to DOCX";
                    }
                }
                catch (Exception ex)
                {
                   
                }
            }
            else
            {
                
            }
        }

        public void ConvertFIle()
        {
            byte[] byteArray = File.ReadAllBytes(filePath);

            if (byteArray != null)
            {
                try
                {
                    DirectoryInfo convertedDocsDirectory =
                        new DirectoryInfo(Path.Combine(DocxConvertedToHtmlDirectory));
                    if (!convertedDocsDirectory.Exists)
                        convertedDocsDirectory.Create();
                    Guid g = Guid.NewGuid();
                    var htmlFileName = g.ToString() + ".html";
                    ConvertToHtml(byteArray, convertedDocsDirectory, htmlFileName);
                    CreatePDFDink(Path.Combine(DocxConvertedToHtmlDirectory, htmlFileName),g.ToString());
                }
                catch (Exception ex)
                {
                    
                }
            }
            else
            {
             
            }

        }
        public static void CreatePDF(string path,string pdffile)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string htmlContent = File.ReadAllText(path);
            PdfDocument pdf = PdfGenerator.GeneratePdf(htmlContent, PdfSharp.PageSize.A4,2);
            pdf.Save(Path.Combine(DocxConvertedToHtmlDirectory, pdffile + ".pdf"));
            PdfDocument pdf2 = PdfGenerator.GeneratePdf("<p><h1>Hello World</h1>This is html rendered text</p>", PdfSharp.PageSize.A4);
            pdf2.Save("document.pdf");
        }

        public void CreatePDFDink(string path, string pdffile)
        {
            string htmlContent = File.ReadAllText(path);
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings { Top = 10 },
                DocumentTitle = "PDF Report",
                Out = Path.Combine(DocxConvertedToHtmlDirectory, pdffile + ".pdf")
            };
            var objectSettings = new ObjectSettings
            {
                PagesCount = true,
                HtmlContent = htmlContent,
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "assets", "styles.css") },
                HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Report Footer" }
            };
            var pdf = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
            };
            IConverter converter = new SynchronizedConverter(new PdfTools());
            converter.Convert(pdf);
        }
        public static void ConvertToHtml(byte[] byteArray, DirectoryInfo destDirectory, string htmlFileName)
        {
            FileInfo fiHtml = new FileInfo(Path.Combine(destDirectory.FullName, htmlFileName));
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var imageDirectoryFullName =
                        fiHtml.FullName.Substring(0, fiHtml.FullName.Length - fiHtml.Extension.Length) + "_files";
                    var imageDirectoryRelativeName =
                        fiHtml.Name.Substring(0, fiHtml.Name.Length - fiHtml.Extension.Length) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc
                        .CoreFilePropertiesPart
                        .GetXDocument()
                        .Descendants(DC.title)
                        .FirstOrDefault();

                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryFullName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to jpeg.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            FileInfo imageFileName = new FileInfo(imageDirectoryFullName + "/image" +
                                imageCounter.ToString() + "." + extension);
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName.FullName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageDirectoryRelativeName + "/" + imageFileName.Name),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = HtmlConverter.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    //var body = html.Descendants(Xhtml.body).First();
                    //body.AddFirst(
                    //    new XElement(Xhtml.p,
                    //        new XElement(Xhtml.a,
                    //            new XAttribute("href", "/WebForm1.aspx"), "Go back to Upload Page")));

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);

                    File.WriteAllText(fiHtml.FullName, htmlString, Encoding.UTF8);
                }
            }
        }
    }
}
