
public class Endpoint: IEndpoint {
    public string node_api { get; set; }
    public string csharp_api { get; set; }
}

public interface IEndpoint {
    string node_api { get; set; }
    string csharp_api { get; set; }
}