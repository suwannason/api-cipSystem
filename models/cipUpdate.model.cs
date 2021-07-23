
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace cip_api.models
{
    public class cipUpdateSchema
    {

        [Key]
        public int id { get; set; }

        [ForeignKey("cipSchema"), Column(TypeName = "int"), StringLength(6), Required]
        public int cipSchemaid { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string planDate { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string actDate { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(50)]
        public string result { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(100)]
        public string reasonDiff { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string fixedAssetCode { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string classFixedAsset { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(120)]
        public string fixAssetName { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(40)]
        public string partNumberDieNo { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(40)]
        public string serialNo { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string processDie { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string model { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string costCenterOfUser { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string tranferToSupplier { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string upFixAsset { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string newBFMorAddBFM { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(100)]
        public string reasonForDelay { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(60)]
        public string addCipBfmNo { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(100)]
        public string remark { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(10)]
        public string boiType { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(10)]
        public string createDate { get; set; }

        [Column(TypeName = "nvarchar"), StringLength(10)]
        public string status { get; set; }
        // public virtual cipSchema cip { get; set; }

    }
}