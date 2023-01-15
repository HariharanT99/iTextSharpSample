using DocumentFormat.OpenXml.Packaging;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using static iTextSharp.text.pdf.AcroFields;

namespace Final
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var filePath = "D:\\Learning\\ModifyWordDoc\\TestDoc\\ActiveTemplate1.pdf";
            //var filePath = "D:\\Learning\\ModifyWordDoc\\TestDoc\\Alignment\\AlliedHealthmemberCertificateTemplate.pdf";
            //var filePath = "D:\\Learning\\ModifyWordDoc\\TestDoc\\ActiveTemplateSample.pdf";
            var pathToSave = "D:\\Learning\\ModifyWordDoc\\TestDoc\\Modified\\Final\\";


            //Create File Save Path 
            if (!Directory.Exists(pathToSave))
                Directory.CreateDirectory(pathToSave);

            using (FileStream newFileStream = new FileStream(pathToSave + "test" + DateTime.UtcNow.Ticks, FileMode.Create))
            {
                PdfReader pdfReader = new PdfReader(filePath);
                PdfStamper pdfstamp = new PdfStamper(pdfReader, newFileStream);
                iTextSharp.text.pdf.PdfStream pdfStream = null;
                List<System.Drawing.Image> ImgList = new List<System.Drawing.Image>();

                
                PdfContentByte contentByte = null;
                PdfContentByte contentByteTemp = null;
                PdfContentByte contentByteTemp1 = null;
                BaseFont baseFont = null;

                TextWithFontExtractionStategy strategy = new TextWithFontExtractionStategy();

                string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, 1, strategy);


                var list = strategy.GetResultantTextInfo();
                var userId = "«User_id» || Calibri ";
                var test = userId.ToLower().Contains("<<user_id>>");
                var textInfoList = list.Where(w => !string.IsNullOrWhiteSpace(w.Text) && (w.Text.ToLower().Contains(MembershipCertificatePlaceholder.UserId) || w.Text.ToLower().Contains(MembershipCertificatePlaceholder.FirstName) || w.Text.ToLower().Contains(MembershipCertificatePlaceholder.LastName) || w.Text.ToLower().Contains(MembershipCertificatePlaceholder.ApprovedDate) || w.Text.ToLower().Contains(MembershipCertificatePlaceholder.Title)));

                foreach (var textInfo in textInfoList)
                {
                    String[] sperator = { "||" };

                    var textList = textInfo.Text.Split(sperator, StringSplitOptions.RemoveEmptyEntries);

                    var sourceText = textList[0];
                    var fontName = textList.Length > 1 ? textList[1] : string.Empty;

                    string fontPath = string.Empty;
                    if (!string.IsNullOrWhiteSpace(fontName))
                        fontPath = "D:\\Learning\\ConsoleApp1\\Font\\" + fontName.Trim().ToLower() + ".TTF";

                    var DefaultFont = "D:\\Learning\\ConsoleApp1\\Font\\" + "lucida calligraphy.TTF";

                    if (File.Exists(fontPath))
                    {
                        baseFont = BaseFont.CreateFont(fontPath, "", false);
                    }
                    else
                    {
                        baseFont = BaseFont.CreateFont(DefaultFont, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    }

                    contentByteTemp = pdfstamp.GetOverContent(1);
                    contentByteTemp.Rectangle(textInfo.TextStart, textInfo.TextEnd - 10, textInfo.Width, textInfo.Height + 13);
                    contentByteTemp.SetColorFill(BaseColor.WHITE);
                    contentByteTemp.SetFontAndSize(baseFont, textInfo.FontHeight);
                    contentByteTemp.Fill();
                    //contentByteTemp1 = pdfstamp.GetOverContent(1).ShowText("TEsdasdfsadsgfad");
                    //contentByteTemp1.Fill();

                    //var pdfObj = pdfReader.GetPdfObject(1);

                    //for (int i = 0; i <= pdfReader.XrefSize - 1; i++)
                    //{
                    //    pdfObj = pdfReader.GetPdfObject(i);
                    //    PdfContentByte contentByteImage = null;

                    //    if ((pdfObj != null) && pdfObj.IsStream())
                    //    {
                    //        pdfStream = (iTextSharp.text.pdf.PdfStream)pdfObj;
                    //        iTextSharp.text.pdf.PdfObject subtype = pdfStream.Get(iTextSharp.text.pdf.PdfName.SUBTYPE);
                    //        if ((subtype != null) && subtype.ToString() == iTextSharp.text.pdf.PdfName.IMAGE.ToString())
                    //        {
                    //        }
                    //        if ((subtype != null) && subtype.ToString() == iTextSharp.text.pdf.PdfName.IMAGE.ToString())
                    //        {
                    //            try
                    //            {
                    //                iTextSharp.text.pdf.parser.PdfImageObject pdfImageObject =
                    //         new iTextSharp.text.pdf.parser.PdfImageObject((iTextSharp.text.pdf.PRStream)pdfStream);

                    //                System.Drawing.Image ImgPDF = pdfImageObject.GetDrawingImage();
                    //                ImgList.Add(ImgPDF);

                    //                //var imagePath = System.IO.Path.Combine(pathToSave, i.ToString());
                    //                //ImgPDF.Save(imagePath, ImageFormat.Jpeg);

                    //                var img = Image.GetInstance(pdfImageObject.GetImageAsBytes());
                    //                img.SetAbsolutePosition(0, 0);
                    //                img.ScaleAbsolute(1500f, 0f);
                    //                img.ScalePercent(0.3f * 100);
                    //                contentByteImage.AddImage(img);

                    //            }
                    //            catch (Exception)
                    //            {

                    //            }
                    //        }
                    //    }

                    //}

                    var replacerText = sourceText.ToLower().Replace(MembershipCertificatePlaceholder.UserId, "123456");
                    replacerText = replacerText.Replace(MembershipCertificatePlaceholder.FirstName, "Hariharan");
                    replacerText = replacerText.Replace(MembershipCertificatePlaceholder.LastName, "Tamilarasan Test");
                    replacerText = replacerText.Replace(MembershipCertificatePlaceholder.Title, "Mr");
                    replacerText = replacerText.Replace(MembershipCertificatePlaceholder.ApprovedDate, "12-Aug-2023").Trim();

                    var replacerTextWidth = baseFont.GetWidthPoint(replacerText, textInfo.FontHeight);
                    var sourceTextWidth = baseFont.GetWidthPoint(sourceText.Trim(), textInfo.FontHeight);

                    var textStartPoint = textInfo.TextStart;

                    if (replacerTextWidth <= sourceTextWidth)
                        textStartPoint += (sourceTextWidth / 2 - replacerTextWidth / 2);
                    else
                        textStartPoint -= (replacerTextWidth / 2 - sourceTextWidth / 2);



                    contentByte = pdfstamp.GetOverContent(1);
                    contentByte.SetColorFill(BaseColor.BLACK);
                    contentByte.SetFontAndSize(baseFont, textInfo.FontHeight);
                    contentByte.BeginText();
                    contentByte.ShowTextAligned(Element.ALIGN_LEFT, replacerText, textStartPoint, textInfo.TextEnd, 0);
                    contentByte.EndText();
                    contentByte.Fill();
                }

                pdfstamp.Close();
                pdfReader.Close();
            }
            
            Console.WriteLine("Found table in the document");

            Console.ReadLine();
        }

        public class TextWithFontExtractionStategy : ITextExtractionStrategy
        {
            private StringBuilder result = new StringBuilder();
            private float llx;
            private float lly;
            private float urx;
            private float ury;
            private StringBuilder text = new StringBuilder();
            private List<TextInformation> textInformationList = new List<TextInformation>();
            TextInformation textInformation = null;

            private Vector lastBaseLine;

            public void RenderText(iTextSharp.text.pdf.parser.TextRenderInfo renderInfo)
            {
                string curFont = renderInfo.GetFont().PostscriptFontName;
                this.text.Append(renderInfo.GetText());

                //This code assumes that if the baseline changes then we're on a newline
                Vector curBaseline = renderInfo.GetBaseline().GetStartPoint();
                Vector topRight = renderInfo.GetAscentLine().GetEndPoint();
                iTextSharp.text.Rectangle rect = new iTextSharp.text.Rectangle(curBaseline[Vector.I1], curBaseline[Vector.I2], topRight[Vector.I1], topRight[Vector.I2]);
                Single curFontSize = rect.Height + 3;

                if (this.lastBaseLine == null || (curBaseline[Vector.I2] != lastBaseLine[Vector.I2]))
                {
                    textInformation = new TextInformation();
                    llx = curBaseline[Vector.I1];
                    lly = curBaseline[Vector.I2];
                    urx = topRight[Vector.I1];
                    textInformation.TextStart = llx;
                    textInformation.TextEnd = lly;
                    textInformation.Height = topRight[Vector.I2] - curBaseline[Vector.I2];
                    textInformation.FontHeight = curFontSize;
                    textInformation.FontFamily = curFont;
                    textInformation.Text = this.text.ToString();
                    textInformation.Width = urx - textInformation.TextStart;
                    this.text = new StringBuilder();
                    this.textInformationList.Add(textInformation);
                }
                else if ((curBaseline[Vector.I2] == lastBaseLine[Vector.I2]))
                {

                    //ury = curBaseline[Vector.I2];
                    urx = topRight[Vector.I1];
                    this.textInformationList.Last().Text = this.textInformationList.Last().Text + this.text.ToString();
                    this.textInformationList.Last().Width = urx - this.textInformationList.Last().TextStart;
                    this.text = new StringBuilder();
                }


                //Set currently used properties
                this.lastBaseLine = curBaseline;
            }

            public List<TextInformation> GetResultantTextInfo()
            {
                return textInformationList;
            }

            //public Font GetFont(string fontName, string filename)
            //{
            //    if (!FontFactory.IsRegistered(fontName))
            //    {
            //        var fontPath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\" + filename;
            //        FontFactory.Register(fontPath);
            //    }
            //    return FontFactory.GetFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            //}
            public void BeginTextBlock()
            {
            }

            public void EndTextBlock()
            {
            }

            public string GetResultantText()
            {
                return result.ToString();
            }

            public void RenderImage(ImageRenderInfo renderInfo)
            {
            }
        }

        public static class MembershipCertificatePlaceholder
        {
            public const string FirstName = "<<first_name>>";
            public const string LastName = "<<last_name>>";
            public const string UserId = "<<user_id>>";
            public const string ApprovedDate = "<<approved_date>>";
            public const string Title = "<<title>>";
        }
        public class TextInformation
        {
            public string Text { get; set; }
            public float FontHeight { get; set; }
            public float TextStart { get; set; }
            public float TextEnd { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public string FontFamily { get; set; }
        }
    }
}
