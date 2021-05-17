
namespace cip_api.request {
    public class CIPupload {
        public Microsoft.AspNetCore.Http.IFormFile file { get; set; }
    }

    public class cipDownlode {

        public int[] id { get; set; }
    }
    public class DateRange {
        public string startDate { get; set; }
        public string endDate { get; set; }
    }
}