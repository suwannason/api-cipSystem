

public class LdapResponse {

    public bool success { get; set; }
    public string message { get; set; }
    public Profile data { get; set; }
}

public class ADprofileResponse {
    public bool success { get; set; }
    public string message { get; set; }
    public ADprofile data { get; set; }
}
public class ADprofile {
    public string email { get; set; }
    public string name { get; set; }
    public string department { get; set; }
}

public class Profile {
    public string empNo { get; set; }
    public string fnameEn { get; set; }
    public string lnameEn { get; set; }
    public string deptCode { get; set; }
    public string deptShortName { get; set; }
    public string deptFullName { get; set; }
    public string divShortName { get; set; }
    public string band { get; set; }
}

public class HRMSResponse {
    public bool success { get; set; }
    public HRMSProfile[] data { get; set; }
}

public class HRMSProfile {
    public string EMP_NO { get; set; }
    public string SNAME_ENG { get; set; }
    public string GNAME_ENG { get; set; }
    public string FNAME_ENG { get; set; }
    public string DEPT_ABB_NAME { get; set; }
    public string DEPT_CODE { get; set; }
}