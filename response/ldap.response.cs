
public class LdapResponse {

    public bool success { get; set; }
    public string message { get; set; }
    public Profile data { get; set; }
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