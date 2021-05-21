using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace cip_api.models
{

    public class PermissionSchema
    {

        [Key, Column(TypeName = "int"), StringLength(6)]
        public int id { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(20)]
        public string empNo { get; set; }
        [Column(TypeName = "nvarchar"), StringLength(20)]
        public string action { get; set; }
        [Column(TypeName = "nvarchar"), StringLength(20)]
        public string deptCode { get; set; }
        [Column(TypeName = "nvarchar"), StringLength(20)]
        public string deptShortName { get; set; }
        [Column(TypeName = "nvarchar"), StringLength(30)]
        public string email { get; set; }
    }
}