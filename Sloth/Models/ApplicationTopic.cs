using System;
using System.Collections.Generic;
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
            Guid = topic.Markup.Topic.Guid;
            Title = topic.Markup.Topic.Title;
            TopicType = topic.Markup.Topic.TopicType;
            Description = topic.Markup.Topic.Description;
            Index = topic.Markup.Topic.Index;
            CreationDate = topic.Markup.Topic.CreationDate;
            CreationAuthor = topic.Markup.Topic.CreationAuthor;
            ModifiedDate = topic.Markup.Topic.ModifiedDate;
            ModifiedAuthor = topic.Markup.Topic.ModifiedAuthor;
            AssignedTo = topic.Markup.Topic.AssignedTo;
            TopicStatus = topic.Markup.Topic.TopicStatus;

            if (topic.Snapshots.FirstOrDefault().Value!=null)
            {
                string base64 = Convert.ToBase64String(topic.Snapshots.FirstOrDefault().Value);
                ImageSource = String.Format("data:image/gif;base64,{0}", base64);
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
