
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
}