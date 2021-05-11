
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace cip_api.models
{
    public class NotificationSchema
    {
        [Key]
        public int id { get; set; }
        [ForeignKey("userSchemaempNo"), Column(TypeName = "nvarchar"), StringLength(10)]
        public string userSchemaempNo { get; set; }

        [Column(TypeName="nvarchar"), StringLength(8), Required]
        public string status { get; set; } // viewed, created
        [Column(TypeName="nvarchar"), StringLength(30)]
        public string title { get; set; }
        [Column(TypeName="nvarchar"), StringLength(80)]
        public string message { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string createDate { get; set; }  // yyyy/MM/dd
    }
}