using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading.Tasks;
using Xbim.BCF;
using Xbim.BCF.XMLNodes;


namespace Sloth.Models
{


    public class DisplayTopic
    {
        public DisplayTopic(Topic topic)
        {
            BCFComment firstComment = topic.Markup.Comments.FirstOrDefault();

            if (firstComment == null)
            {
                firstComment = new BCFComment(System.Guid.NewGuid(), topic.Markup.Topic.Guid, " ", DateTime.Now, " ", " ");
            }

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

            ImageSource = "";
            if (topic.Snapshots.Count != 0)
            {
                byte[] snapshot = topic.Snapshots.First(kvp => kvp.Key.Contains("snapshot")).Value;
                if (snapshot != null)
                {
                    ImageSource = Services.BCFServices.GetImageFromBytes(snapshot, true);
                }
                else
                {
                    ImageSource = Services.BCFServices.GetImageFromBytes(topic.Snapshots.FirstOrDefault().Value, true);
                }
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
