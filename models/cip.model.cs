
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace cip_api.models {

    public class cipSchema {
        // [Key]
        [Key, Column(TypeName="int"), StringLength(6)]
        public int id { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string cipNo { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string subCipNo { get; set; }

        [Column(TypeName="nvarchar"), StringLength(30)]
        public string poNo { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string vendorCode { get; set; }

        [Column(TypeName="nvarchar"), StringLength(20)]
        public string vendor { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string acqDate { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string invDate { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string receivedDate { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string invNo { get; set; }

        [Column(TypeName="nvarchar"), StringLength(80)]
        public string name { get; set; }

        [Column(TypeName="nvarchar"), StringLength(5)]
        public string qty { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string exRate { get; set; }

        [Column(TypeName="nvarchar"), StringLength(4)]
        public string cur { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string perUnit { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string totalJpy { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string totalThb { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string averageFreight { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10)]
        public string averageInsurance { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15)]
        public string totalJpy_1 { get; set; }

        [Column(TypeName="nvarchar"), StringLength(30)]
        public string totalThb_1 { get; set; }

        [Column(TypeName="nvarchar"), StringLength(30)]
        public string perUnitThb { get; set; }

        [Column(TypeName="nvarchar"), StringLength(5)]
        public string cc { get; set; }

        [Column(TypeName="nvarchar"), StringLength(30), DefaultValue("-")]
        public string totalOfCip { get; set; }

        [Column(TypeName="nvarchar"), StringLength(20), DefaultValue("-")]
        public string budgetCode { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10), DefaultValue("-")]
        public string prDieJig { get; set; }

        [Column(TypeName="nvarchar"), StringLength(10), DefaultValue("-") ]
        public string model { get; set; }

        [Column(TypeName="nvarchar"), StringLength(15), DefaultValue("-") ]

        public ICollection<cipUpdateSchema> cip { get; set; }
        public string status { get; set; }

    }
}