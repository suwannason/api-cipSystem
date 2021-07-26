
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace cip_api.models
{

    public class ApprovalSchema
    {

        [Key]
        public int id { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(15), Required]
        public string onApproveStep { get; set; } // cc-prepared, cc-checked, cc-approved, --> cost-checked, cost-approveed
        
        [Column(TypeName="nvarchar"), StringLength(15)]
        public string empNo { get; set; }
        
        [ForeignKey("cipSchemaid"), Column(TypeName = "int"), StringLength(6), Required]
        public int cipSchemaid { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(20)]
        public string date { get; set; }
    }
}