
namespace cip_api.request
{

    public class ACCsendBack
    {
        public string[] id { get; set; }
        public string commend { get; set; }
        public string toStep { get; set; } // requester-prepare, user-prepare
    }
}