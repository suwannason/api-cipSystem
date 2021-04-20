
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace cip_api.models
{

    public class ApprovalSchema
    {

        [Key]
        public int id { get; set; }
        public string onApproveStep { get; set; } // input, check approve
        public string empNo { get; set; }

        public int cipUpdateid { get; set; }
    }
}