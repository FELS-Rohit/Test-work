using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace Read_Excel_File
{
    public class ConvertToHTMLWithImages
    {
        public static void ConVertDOcWithImages()
        {
            string sourceDocumentFileName = "Test1.docx";
            FileInfo fileInfo = new FileInfo(sourceDocumentFileName);
            string imageDirectoryName = fileInfo.Name.Substring(0,
                fileInfo.Name.Length - fileInfo.Extension.Length) + "_files";
            DirectoryInfo dirInfo = new DirectoryInfo(imageDirectoryName);
            if (dirInfo.Exists)
            {
                // Delete the directory and files.
                foreach (var f in dirInfo.GetFiles())
                    f.Delete();
                dirInfo.Delete();
            }
            int imageCounter = 0;
            byte[] byteArray = File.ReadAllBytes(sourceDocumentFileName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(memoryStream, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "Test Title",
                        ConvertFormatting = false,
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings,
                        imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert the .png file to a .jpeg file.
                                extension = "jpeg";
                                imageFormat = ImageFormat.Jpeg;
                            }
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                                imageFormat = ImageFormat.Tiff;

                            // If the image format is not one that you expect, ignore it,
                            // and do not return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        });

                    // Note: the XHTML returned by the ConvertToHtmlTransform method contains objects of type
                    // XEntity. PtOpenXmlUtil.cs define the XEntity class. For more information
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx.

                    //
                    // If you transform the XML tree returned by the ConvertToHtmlTransform method further, you
                    // must do it correctly, or entities do not serialize correctly.

                    File.WriteAllText(fileInfo.Directory.FullName + "/" + fileInfo.Name.Substring(0,
                        fileInfo.Name.Length - fileInfo.Extension.Length) + ".html",
                        html.ToStringNewLineOnAttributes());
                }
            }
        }
    }
}
