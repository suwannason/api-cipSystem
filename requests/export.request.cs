
namespace cip_api.request {

    public class ExportACCrequest {
        public string workType { get; set; }
    }

    public class exportToForm {
        public string workType { get; set; }
        public string[] id { get; set; }
    }

    public class getHistory {
        public string workType { get; set; }
        public System.Int32 page { get; set; }
        public System.Int32 perPage { get; set; } 
    }
}