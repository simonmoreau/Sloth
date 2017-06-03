using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading.Tasks;
using Xbim.BCF;
using Xbim.BCF.XMLNodes;
using System.IO;
using Novacode;
using System.Drawing.Imaging;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;


namespace Sloth.Models
{
    public class DisplayBCF
    {
        public DisplayBCF(BCF bcf, string filename)
        {
            BCF = bcf;
            FileName = filename;
        }

        ///// <summary>
        ///// A list of topics to be displayed
        ///// </summary>
        //public List<DisplayTopic> DisplayTopics { get; }

        /// <summary>
        /// The BCF filename
        /// </summary>
        /// 
        public string FileName { get; }

        /// <summary>
        /// The original BCF file
        /// </summary>
        /// 
        public BCF BCF { get; }

        public MemoryStream ExportAsWord()
        {
            MemoryStream ms = new MemoryStream();
            string wordFileName = Path.GetFileNameWithoutExtension(this.FileName) + ".docx";
            // Create a document in memory:
            using (DocX doc = DocX.Create(wordFileName))
            {
                //if (_templatePath != null)
                //{
                //    doc.ApplyTemplate(_templatePath);
                //}

                int i = 1;

                foreach (Topic topic in BCF.Topics)
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

        public MemoryStream ExportAsExcel()
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

                foreach (Topic topic in BCF.Topics)
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

        private void AddTopicToWord(Topic topic, DocX doc)
        {
            Novacode.Paragraph p;

            // Insert a paragrpah:
            p = doc.InsertParagraph(topic.Markup.Topic.Title);
            //p.StyleName = _styles.TitleStyle;

            doc.InsertParagraph("");

            //Insert the date of the note
            if (topic.Markup.Comments.FirstOrDefault() != null)
            {
                p = doc.InsertParagraph("Note created on " + topic.Markup.Comments[0].Date.ToString() + " by " + topic.Markup.Comments[0].Author);
                //p.StyleName = _styles.DateStyle;

                p = doc.InsertParagraph("Status : " + topic.Markup.Comments[0].Status + " - " + topic.Markup.Comments[0].VerbalStatus);
                //p.StyleName = _styles.DateStyle;

                p = doc.InsertParagraph(topic.Markup.Comments[0].Comment);
                //p.StyleName = _styles.ContentStyle;
            }


            if (topic.Snapshots != null)
            {
                if (topic.Snapshots.Count != 0)
                {
                    System.IO.MemoryStream myMemStream = Services.BCFServices.GetImageStreamFromBytes(topic.Snapshots.FirstOrDefault().Value, true);
                    //System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

                    // Add an Image to the docx file
                    Novacode.Image img = doc.AddImage(myMemStream);
                    //myMemStream.Close();
                    Novacode.Picture pic = img.CreatePicture(); // img.CreatePicture(450, 600);

                    p = doc.InsertParagraph("", false);
                    p.InsertPicture(pic);
                }
            }


            if (topic.Markup.Comments != null)
            {
                int commentCount = topic.Markup.Comments.Count();
                if (commentCount > 1)
                {
                    for (int j = 1; j < commentCount; j++)
                    {
                        p = doc.InsertParagraph("Note created on " + topic.Markup.Comments[j].Date.ToString() + " by " + topic.Markup.Comments[0].Author);
                        //p.StyleName = _styles.DateStyle;

                        p = doc.InsertParagraph(topic.Markup.Comments[j].Comment);
                        //p.StyleName = _styles.ContentStyle;
                    }
                }
            }

            p.InsertPageBreakAfterSelf();
        }

        private void AddTopicToExcel(Topic topic, ExcelWorksheet worksheet, int i)
        {
            //Set the row heigth
            worksheet.Row(i).Height = 150;
            // Insert a paragrpah:
            worksheet.Cells[i,1].Value = topic.Markup.Topic.Title;
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

                    shape.SetPosition(i-1, 0, 7, 0);
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
    }

    public class DisplayTopic
    {
        public DisplayTopic(Topic topic)
        {
            BCFComment firstComment = topic.Markup.Comments.FirstOrDefault();

            Guid = topic.Markup.Topic.Guid;
            Title = topic.Markup.Topic.Title ;
            TopicType = topic.Markup.Topic.TopicType;
            Description = !String.IsNullOrEmpty(topic.Markup.Topic.Description) ? topic.Markup.Topic.Description : firstComment.Comment;
            Index = topic.Markup.Topic.Index;
            CreationDate = topic.Markup.Topic.CreationDate != null ? topic.Markup.Topic.CreationDate : firstComment.Date;
            CreationAuthor = !String.IsNullOrEmpty(topic.Markup.Topic.CreationAuthor) ? topic.Markup.Topic.CreationAuthor : firstComment.Author;
            ModifiedDate =  topic.Markup.Topic.ModifiedDate != null ? topic.Markup.Topic.ModifiedDate: firstComment.ModifiedDate;
            ModifiedAuthor = !String.IsNullOrEmpty(topic.Markup.Topic.ModifiedAuthor) ? topic.Markup.Topic.ModifiedAuthor : firstComment.ModifiedAuthor;
            AssignedTo = topic.Markup.Topic.AssignedTo;
            TopicStatus = !String.IsNullOrEmpty(topic.Markup.Topic.TopicStatus) ? topic.Markup.Topic.TopicStatus: firstComment.Status;

            if (topic.Snapshots.FirstOrDefault().Value!=null)
            {
                ImageSource = Services.BCFServices.GetImageFromBytes(topic.Snapshots.FirstOrDefault().Value, true);
            }
            else
            {
                ImageSource = "";
            }

        }

        public int DisplayTopicId { get; set; }

        /// <summary>
        /// The path to the image in the first comments
        /// </summary>
        public string ImageSource { get; }


        /// <summary>
        /// The topic identifier
        /// </summary>
        public Guid Guid { get; }

        /// <summary>
        /// Title of the topic
        /// </summary>
        public String Title {get;}

        /// <summary>
        /// The type of the topic (the options can be specified in the extension schema)
        /// </summary>
        public String TopicType { get;  }

        /// <summary>
        /// Description of the topic
        /// </summary>
        public String Description { get;  }

        /// <summary>
        /// Number to maintain the order of the topics
        /// </summary>
        public int? Index { get; }

        /// <summary>
        /// Date when the topic was created
        /// </summary>
        public DateTime? CreationDate { get; }

        /// <summary>
        /// User who created the topic
        /// </summary>
        public String CreationAuthor { get;  }

        /// <summary>
        /// Date when the topic was last modified
        /// </summary>
        public DateTime? ModifiedDate { get; }

        /// <summary>
        /// User who modified the topic
        /// </summary>

        public String ModifiedAuthor { get;  }

        /// <summary>
        /// The user to whom this topic is assigned to
        /// </summary>

        public String AssignedTo { get;  }

        /// <summary>
        /// The status of the topic (the options can be specified in the extension schema)
        /// </summary>

        public String TopicStatus { get; }


    }
}
