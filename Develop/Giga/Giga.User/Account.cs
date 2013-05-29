using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Security.Cryptography;
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
        [Key]
        [MaxLength(256)]
        [Required(AllowEmptyStrings=false)]
        public String ID { get; set; }

        /// <summary>
        /// User name
        /// </summary>
        [DataMember]
        [MaxLength(256)]
        [Required(AllowEmptyStrings=false)]
        public String Name { get; set; }

        /// <summary>
        /// Password
        /// </summary>
        [DataMember]
        [MaxLength(256)]
        internal String Password { get; set; }

        /// <summary>
        /// Time when this account was created
        /// </summary>
        [DataMember]
        public DateTime CreatedTime { get; set;  }

        /// <summary>
        /// People who created this account
        /// </summary>
        [DataMember]
        [MaxLength(256)]
        public String CreatedBy { get; set; }

        /// <summary>
        /// Last time when this account been modified
        /// </summary>
        [DataMember]
        public DateTime? ModifiedTime { get; set; }

        /// <summary>
        /// People who modified this account last time
        /// </summary>
        [DataMember]
        [MaxLength(256)]
        public String ModifiedBy { get; set; }

        /// <summary>
        /// Tool to encrypt password string
        /// </summary>
        /// <param name="pwd">Original password string</param>
        /// <returns>Encrypted password</returns>
        internal static String EncryptPassword(String pwd)
        {
            byte[] eBytes = MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(pwd));
            StringBuilder b = new StringBuilder();
            foreach (byte e in eBytes)
            {
                b.AppendFormat("{0:X2}", e);
            }
            return b.ToString();
        }
    }
}
