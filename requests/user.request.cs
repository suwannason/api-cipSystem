
using System.ComponentModel.DataAnnotations;

namespace cip_api.request.user
{

    public class Login
    {

        [Required]
        public string username { get; set; }

        [Required]
        public string password { get; set; }
    }

    public class DeptResponse
    {
        public bool success { get; set; }
        public deptItem[] data { get; set; }

    }
    public class deptItem
    {
        public string DEPT_CODE { get; set; }
        public string DEPT_NAME { get; set; }
        public string DEPT_ABB_NAME { get; set; }
        public string DIV_NAME_WC { get; set; }
    }
    public class Upload
    {

        public Microsoft.AspNetCore.Http.IFormFile file { get; set; }
    }

    public class ProfileUser
    {
        public string username { get; set; }
        public string deptCode { get; set; }
    }

    public class createUser
    {
        [Required]
        public string empNo { get; set; }
        public string permission { get; set; }
        public string action { get; set; }
    }

    public class updateUser
    {
        public string oldEmpNo { get; set; }
        public string newEmpNo { get; set; }
        public string oldPremission { get; set; }
        public string newPremission { get; set; }
        public string oldAction { get; set; }
        public string newAction { get; set; }
    }
}