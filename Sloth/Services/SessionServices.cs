using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Novacode;
using System.Drawing.Imaging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using Xbim.BCF;
using Xbim.BCF.XMLNodes;
using Sloth.Models;

namespace Sloth.Services
{


    public static class SessionExtensionMethods
    {
        public static void SetObject(this ISession session, string key, object value)
        {
            string stringValue = JsonConvert.SerializeObject(value);
            session.SetString(key, stringValue);
        }

        public static T GetObject<T>(this ISession session, string key)
        {
            string stringValue = session.GetString(key);
            T value = JsonConvert.DeserializeObject<T>(stringValue);
            return value;
        }
    }


    public static class BCFServices
    {
        public static string GetImageFromBytes(byte[] bytes, bool compress)
        {
            System.IO.MemoryStream myMemStream = new System.IO.MemoryStream(bytes);
            System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

            if (compress)
            {
                System.Drawing.Image newImage = fullsizeImage.GetThumbnailImage(512, 512, null, IntPtr.Zero);
                myMemStream = new System.IO.MemoryStream();
                newImage.Save(myMemStream, System.Drawing.Imaging.ImageFormat.Jpeg);  //Or whatever format you want.
            }


            string base64 = Convert.ToBase64String(myMemStream.ToArray());
            return String.Format("data:image/gif;base64,{0}", base64);
        }

        public static MemoryStream GetImageStreamFromBytes(byte[] bytes, bool compress)
        {
            System.IO.MemoryStream myMemStream = new System.IO.MemoryStream(bytes);

            if (compress)
            {
                System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);
                myMemStream.Close();
                System.Drawing.Image compressedImage = fullsizeImage.GetThumbnailImage(512, 512, null, IntPtr.Zero);
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                compressedImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);  //Or whatever format you want.
                return ms;
            }
            else
            {
                return myMemStream;
            }


        }

    }

    public static class ExportServices
    {
        public static string _wordTemplatePath;

        public static MemoryStream ExportAsWord(List<FilePath> BCFFiles, string wordTemplatePath)
        {
            _wordTemplatePath = wordTemplatePath;

            List<Topic> topics = new List<Topic>();

            foreach (FilePath BCFFile in BCFFiles)
            {
                topics.AddRange(LoadTopics(BCFFile.FullFilePath));
            }

            return CreateWordFile(topics);
        }

        public static MemoryStream ExportAsExcel(List<FilePath> BCFFiles)
        {
            List<Topic> topics = new List<Topic>();

            foreach (FilePath BCFFile in BCFFiles)
            {
                topics.AddRange(LoadTopics(BCFFile.FullFilePath));
            }

            return CreateExcelFile(topics);
        }

        private static List<Topic> LoadTopics(string bcfFilePath)
        {
            using (FileStream stream = System.IO.File.Open(bcfFilePath, FileMode.Open))
            {
                BCF bcf = BCF.Deserialize(stream);

                return bcf.Topics;
            }
        }

        private static MemoryStream CreateWordFile(List<Topic> topics)
        {
            MemoryStream ms = new MemoryStream();

            // Create a document in memory:
            using (DocX doc = DocX.Create("wordFileName.docx"))
            {

                if (!string.IsNullOrEmpty(_wordTemplatePath))
                {
                    doc.ApplyTemplate(_wordTemplatePath);
                }

                int i = 1;

                //Sort topics by date
                List<Topic> sortedTopics = topics.OrderBy(x => x.Markup.Topic.Description).ToList();

                foreach (Topic topic in sortedTopics)
                {
                    //Add note to the report
                    AddTopicToWord(topic, doc);
                    //(sender as BackgroundWorker).ReportProgress(i);
                    i++;
                }

                doc.SaveAs(ms);
            }

            return ms;
        }

        private static MemoryStream CreateExcelFile(List<Topic> topics)
        {
            MemoryStream ms = new MemoryStream();

            using (ExcelPackage package = new ExcelPackage(ms))
            {
                int i = 2;

                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Report");
                //Add the headers
                worksheet.Cells[1, 1].Value = "Title";
                worksheet.Cells[1, 2].Value = "Author";
                worksheet.Cells[1, 3].Value = "Date";
                worksheet.Cells[1, 4].Value = "Time";
                worksheet.Cells[1, 5].Value = "Status";
                worksheet.Cells[1, 6].Value = "VerbalStatus";
                worksheet.Cells[1, 7].Value = "Comment";
                worksheet.Cells[1, 8].Value = "Picture";

                //Sort topics by date
                List<Topic> sortedTopics = topics.OrderBy(x => x.Markup.Topic.Description).ToList();

                foreach (Topic topic in sortedTopics)
                {
                    //Add note to the report
                    AddTopicToExcel(topic, worksheet, i);
                    //(sender as BackgroundWorker).ReportProgress(i);
                    i++;
                }


                // set some document properties
                package.Workbook.Properties.Title = "BCF Report";
                package.Workbook.Properties.Author = "Simon Moreau";
                package.Workbook.Properties.Comments = "This is an Excel report of your BCF file";

                // set some extended property values
                package.Workbook.Properties.Company = "BIM 42";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Simon Moreau");

                package.Save();
                //package.SaveAs(ms);
            }

            return ms;
        }

        private static void AddTopicToWord(Topic topic, DocX doc)
        {
            Novacode.Paragraph p;

            // Insert a paragrpah:
            p = doc.InsertParagraph(topic.Markup.Topic.Title);
            //p.StyleName = _styles.TitleStyle;
            p.StyleName = "Heading1";


            //Add elements of the topics
            p = AddLine("Note created ", "on ", topic.Markup.Topic.CreationDate.ToString(), " by ", topic.Markup.Topic.CreationAuthor,p,doc);
            if (!string.IsNullOrEmpty(topic.Markup.Topic.CreationDate.ToString()) || !string.IsNullOrEmpty(topic.Markup.Topic.CreationAuthor))  p.StyleName = "Heading2";
            p = AddLine("Note assigned to ", "", topic.Markup.Topic.AssignedTo, "", "", p, doc);
            p = AddLine("", "", topic.Markup.Topic.Description, "", "", p, doc);
            p = AddLine("Note modified ", "on ", topic.Markup.Topic.ModifiedDate.ToString(), " by ", topic.Markup.Topic.ModifiedAuthor, p, doc);

            //Check if any comment contain a viewpoint
            if (topic.Markup.Comments.Where(x=>x.Viewpoint != null).Count() == 0)
            {
                p = AddPicture(topic, "snapshot", doc, p);
            }

            p = AddComments(topic, doc,p);

            p.InsertPageBreakAfterSelf();
        }

        private static void AddTopicToExcel(Topic topic, ExcelWorksheet worksheet, int i)
        {
            //Set the row heigth
            worksheet.Row(i).Height = 150;
            // Insert a paragrpah:
            worksheet.Cells[i, 1].Value = topic.Markup.Topic.Title;
            //p.StyleName = _styles.TitleStyle;

            //Insert the date of the note
            if (topic.Markup.Comments.FirstOrDefault() != null)
            {
                //OfficeOpenXml.Style.ExcelNumberFormat timeFormat = new OfficeOpenXml.Style.ExcelNumberFormat();

                worksheet.Cells[i, 2].Value = topic.Markup.Comments[0].Author;
                worksheet.Cells[i, 3].Value = topic.Markup.Comments[0].Date.ToString("yyyy/mm/dd");
                //worksheet.Cells[i, 3].Style.Numberformat.Format = "yyyy/mm/dd";
                worksheet.Cells[i, 4].Value = topic.Markup.Comments[0].Date.ToString("HH:mm:ss");
                //worksheet.Cells[i, 4].Style.Numberformat.Format = "hh:mm:ss";
                worksheet.Cells[i, 5].Value = topic.Markup.Comments[0].Status;
                worksheet.Cells[i, 6].Value = topic.Markup.Comments[0].VerbalStatus;
                worksheet.Cells[i, 7].Value = topic.Markup.Comments[0].Comment;
                worksheet.Cells[i, 7].Style.WrapText = true;
                //worksheet.Cells
            }


            if (topic.Snapshots != null)
            {
                if (topic.Snapshots.Count != 0)
                {
                    System.IO.MemoryStream myMemStream = Services.BCFServices.GetImageStreamFromBytes(topic.Snapshots.FirstOrDefault().Value, true);
                    System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

                    // Add an Image to the xlsx file
                    ExcelPicture shape = worksheet.Drawings.AddPicture(topic.Markup.Topic.Guid.ToString(), fullsizeImage);

                    double sizeRatio = fullsizeImage.Width / fullsizeImage.Height;

                    shape.SetPosition(i - 1, 0, 7, 0);
                    int height = 200;
                    int width = Convert.ToInt32(Math.Round(height * sizeRatio));

                    shape.SetSize(width, height);
                }
            }


            //if (topic.Markup.Comments != null)
            //{
            //    int commentCount = topic.Markup.Comments.Count();
            //    if (commentCount > 1)
            //    {
            //        for (int j = 1; j < commentCount; j++)
            //        {
            //            p = doc.InsertParagraph("Note created on " + topic.Markup.Comments[j].Date.ToString() + " by " + topic.Markup.Comments[0].Author);
            //            //p.StyleName = _styles.DateStyle;

            //            p = doc.InsertParagraph(topic.Markup.Comments[j].Comment);
            //            //p.StyleName = _styles.ContentStyle;
            //        }
            //    }
            //}
        }

        private static Paragraph AddPicture(Topic topic, string pictureName, DocX doc, Paragraph p)
        {
            if (topic.Snapshots == null) return p;

            List<KeyValuePair<string, byte[]>> snapshots = topic.Snapshots.Where(item => item.Key.Contains(pictureName)).ToList();

            if (snapshots.Count == 0) return p;

            System.IO.MemoryStream myMemStream = Services.BCFServices.GetImageStreamFromBytes(snapshots.FirstOrDefault().Value, false);
            //System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

            // Add an Image to the docx file
            Novacode.Image img = doc.AddImage(myMemStream);
            //myMemStream.Close();
            Novacode.Picture pic = img.CreatePicture();
            //Create a image with more or less the width of the page
            int newWidth = 568;
            double ratio = Convert.ToDouble(pic.Height) / Convert.ToDouble(pic.Width);
            int newHeight = Convert.ToInt16(Math.Round(ratio * newWidth));
            pic.Height = newHeight;
            pic.Width = newWidth;

            p = doc.InsertParagraph("");
            p = doc.InsertParagraph("");

            p.InsertPicture(pic);

            return p;
        }

        private static Paragraph AddComments(Topic topic, DocX doc, Paragraph p)
        {
            if (topic.Markup.Comments == null) return p;

            List<BCFComment> sortedComment = topic.Markup.Comments.OrderBy(x => x.Date).ToList();

            foreach (BCFComment comment in sortedComment)
            {
                p = AddLine("Comment created ", "on ", comment.Date.ToString(), " by ", comment.Author, p, doc);
                p.StyleName = "Heading2";
                p = AddLine("Status: ", "", comment.Status, " - ", comment.VerbalStatus, p, doc);
                p = AddLine("", "", comment.Comment, "", "", p, doc);
                p = AddLine("Comment modified ", "on ", comment.ModifiedDate.ToString(), " by ", comment.ModifiedAuthor, p, doc);

                if (comment.Viewpoint != null)
                {
                    List<BCFViewpoint> viewpoints = topic.Markup.Viewpoints.Where(viewpoint => viewpoint.ID == comment.Viewpoint.ID).ToList();

                    if (viewpoints.Count != 0)
                    {
                        BCFViewpoint viewpoint = viewpoints.FirstOrDefault();

                        p = AddPicture(topic, viewpoint.Snapshot, doc, p);
                    }
                }

                p = doc.InsertParagraph("");
            }

            return p;
        }

        private static Paragraph AddLine(string prefix, string text1, string value1, string text2, string value2, Paragraph p, DocX doc)
        {
            if (!string.IsNullOrEmpty(value1) && !string.IsNullOrEmpty(value2))
            {
                p = doc.InsertParagraph(prefix + text1 + value1 + text2 + value2);
            }
            else if (!string.IsNullOrEmpty(value1))
            {
                p = doc.InsertParagraph(prefix + text1 + value1);
            }
            else if (!string.IsNullOrEmpty(value2))
            {
                p = doc.InsertParagraph(prefix + text2 + value2);
            }

            return p;
        }
    }

    
}
