
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace cip_api.models
{

    public class users
    {
        public string empNo { get; set; }
        public string deptCode { get; set; }
        public string dept { get; set; }
        public string div { get; set; }
        public string band { get; set; }
        public string name { get; set; }
    }

    public class userSchema {

        [Key, Column(TypeName="nvarchar"), StringLength(6)]
        public string empNo { get; set; }
        [Column(TypeName="nvarchar"), StringLength(20)]
        public string action { get; set; }
    }
}