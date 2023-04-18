namespace EmailConceptG;

public class OAuth2DTO
{
    public string ClientId { get; set; }
    public string TenantId { get; set; }
    public string UserName { get; set; }
    public string Password { get; set; }
}

public class CachedTokenDTO
{
    public string UserName { get; set; }
    public string TokenString { get; set; }
    public DateTime ExpirationDate { get; set; }
}