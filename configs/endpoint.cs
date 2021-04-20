
public class Endpoint: IEndpoint {
    public string node_api { get; set; }
    public string ldap_auth { get; set; }
    
}

public interface IEndpoint {
    string node_api { get; set; }
    string ldap_auth { get; set; }
}