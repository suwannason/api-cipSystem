
namespace cip_api.request {
    public class CIPupload {
        public Microsoft.AspNetCore.Http.IFormFile file { get; set; }
        public string type { get; set; }
    }

    public class cipDownlode {

        public int[] id { get; set; }
    }
    public class DateRange {
        public string startDate { get; set; }
        public string endDate { get; set; }
    }

    public class rejectCip {
        public System.Int32[] id { get; set; }
        public string commend { get; set; }
    }

    public class rejectCostCenter {
        public System.Int32[] id { get; set; }
        public string commend { get; set; }
    }

    public class editCip {
        public System.Int32 id { get; set; }
        public string workType { get; set; }
        public string projectNo { get; set; }
        public string cipNo { get; set; }
        public string subCipNo { get; set; }
        public string poNo { get; set; }
        public string vendorCode { get; set; }
        public string vendor { get; set; }
        public string acqDate { get; set; }
        public string invDate { get; set; }
        public string receivedDate { get; set; }
        public string invNo { get; set; }
        public string name { get; set; }
        public string qty { get; set; }
        public string exRate { get; set; }
        public string cur { get; set; }
        public string perUnit { get; set; }
        public string totalJpy { get; set; }
        public string totalThb { get; set; }
        public string averageFreight { get; set; }
        public string averageInsurance { get; set; }
        public string totalJpy_1 { get; set; }
        public string totalThb_1 { get; set; }
        public string perUnitThb { get; set; }
        public string cc { get; set; }
        public string totalOfCip { get; set; }
        public string budgetCode { get; set; }
        public string prDieJig { get; set; }
        public string model { get; set; }
        public string partNoDieNo { get; set; }
    }
}