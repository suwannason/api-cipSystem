
using System.ComponentModel.DataAnnotations;

namespace cip_api.request
{

    public class cipUpdate
    {
        [Required]
        public int[] cipSchemaid { get; set; }
        public string planDate { get; set; }
        public string actDate { get; set; }
        public string result { get; set; }
        public string reasonDiff { get; set; }
        public string fixedAssetCode { get; set; }
        public string classFixedAsset { get; set; }
        public string fixAssetName { get; set; }
        public string partNumberDieNo { get; set; }
        public string serialNo { get; set; }
        public string processDie { get; set; }
        public string model { get; set; }
        public string costCenterOfUser { get; set; }
        public string tranferToSupplier { get; set; }
        public string upFixAsset { get; set; }
        public string newBFMorAddBFM { get; set; }
        public string reasonForDelay { get; set; }
        public string boiType { get; set; }
        public string remark { get; set; }
    }


    public class cipUpdateEdit
    {
        public string id { get; set; }
        [Required]
        public string planDate { get; set; }
        [Required]
        public string actDate { get; set; }
        [Required]
        public string result { get; set; }
        [Required]
        public string reasonDiff { get; set; }
        [Required]
        public string fixedAssetCode { get; set; }
        [Required]
        public string classFixedAsset { get; set; }
        [Required]
        public string fixAssetName { get; set; }
        [Required]
        public string partNumberDieNo { get; set; }
        [Required]
        public string serialNo { get; set; }
        [Required]
        public string processDie { get; set; }
        [Required]
        public string model { get; set; }
        [Required]
        public string costCenterOfUser { get; set; }
        [Required]
        public string tranferToSupplier { get; set; }
        [Required]
        public string upFixAsset { get; set; }
        [Required]
        public string newBFMorAddBFM { get; set; }
        [Required]
        public string reasonForDelay { get; set; }
        [Required]
        public string addCipBfmNo { get; set; }
        [Required]
        public string remark { get; set; }
        public string boiType { get; set; }
    }
}