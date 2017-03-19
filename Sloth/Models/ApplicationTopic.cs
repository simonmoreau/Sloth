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

        public Stream ExportAsWord(string filename)
        {
            // Create a document in memory:
            using (DocX doc = DocX.Create(filename))
            {
                //if (_templatePath != null)
                //{
                //    doc.ApplyTemplate(_templatePath);
                //}

                int i = 1;

                foreach (Topic topic in BCF.Topics)
                {
                    //Add note to the report
                    AddNoteToReport(topic, doc);
                    //(sender as BackgroundWorker).ReportProgress(i);
                    i++;
                }

                doc.Save();
            }

            System.IO.FileStream stream = new System.IO.FileStream(filename, System.IO.FileMode.Open);


            return stream;
        }

        private void AddNoteToReport(Topic topic, DocX doc)
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
                System.IO.MemoryStream myMemStream = new System.IO.MemoryStream(topic.Snapshots.FirstOrDefault().Value);
                //System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

                // Add an Image to the docx file
                Novacode.Image img = doc.AddImage(myMemStream);
                Novacode.Picture pic = img.CreatePicture(); // img.CreatePicture(450, 600);

                p = doc.InsertParagraph("", false);
                p.InsertPicture(pic);

                //using (Stream stream = Services.BCFServices.GetImageStreamFromBytes(,false))
                //{

                //}
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
