using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Giga.User
{
    /// <summary>
    /// User account
    /// </summary>
    [DataContract]
    public class Account
    {
        /// <summary>
        /// Unique identity of account
        /// </summary>
        [DataMember]
        public String ID { get; set; }

        /// <summary>
        /// User name
        /// </summary>
        [DataMember]
        public String Name { get; set; }

        /// <summary>
        /// Time when this account was created
        /// </summary>
        [DataMember]
        public DateTime CreatedTime { get; set;  }

        /// <summary>
        /// People who created this account
        /// </summary>
        [DataMember]
        public String CreatedBy { get; set; }

        /// <summary>
        /// Last time when this account been modified
        /// </summary>
        [DataMember]
        public DateTime ModifiedTime { get; set; }

        /// <summary>
        /// People who modified this account last time
        /// </summary>
        [DataMember]
        public String ModifiedBy { get; set; }
    }
}
