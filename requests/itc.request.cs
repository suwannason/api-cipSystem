
namespace cip_api.request
{

    public class ITCConfirm
    {
        public confirmBox[] confirm { get; set; }
    }

    public class confirmBox
    {
        public System.Int32 id { get; set; }
        public string boiType { get; set; }
    }
}