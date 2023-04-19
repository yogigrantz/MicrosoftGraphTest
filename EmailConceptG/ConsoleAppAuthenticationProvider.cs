using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;

namespace EmailConceptG;

public class ConsoleAppAuthenticationProvider : IAuthenticationProvider
{
    private string _cachefileName;
    private string[] _scopes;

    private string _accessToken = null;
    private NetworkCredential _nc = null;
    private IPublicClientApplication _app = null;

    public ConsoleAppAuthenticationProvider(OAuth2DTO oauth2, string cacheFileName, string[] scopes)
    {
        _scopes = scopes;
        _cachefileName = cacheFileName;
        _app = PublicClientApplicationBuilder
                                         .Create(oauth2.ClientId)
                                         .WithTenantId(oauth2.TenantId)
                                         .Build();
        _nc = new NetworkCredential(oauth2.UserName, oauth2.Password);
    }

    public  Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
    {
        if (this._accessToken == null)
            this._accessToken = GetCachedToken(_nc.UserName);

        if (this._accessToken == null || this._accessToken.StartsWith("AuthError"))
        {
            try
            {
                AuthenticationResult? tokenResult = _app.AcquireTokenByUsernamePassword(_scopes, _nc.UserName, _nc.SecurePassword).ExecuteAsync().Result;
                this._accessToken = tokenResult.AccessToken;
                SetCachedToken(_nc.UserName, this._accessToken);
            }
            catch (AggregateException ex)
            {
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        try
        {
            if (!request.Headers.TryGetValue("Authorization", out var auth))
            {
                if (this._accessToken != null && !this._accessToken.StartsWith("AuthError"))
                    request.Headers.Add("Authorization", "Bearer " + this._accessToken);
            }
        }
        catch (Exception)
        {
            if (this._accessToken != null && !this._accessToken.StartsWith("AuthError"))
                request.Headers.Add("Authorization", "Bearer " + this._accessToken);
        }

        return Task.CompletedTask;

    }

    public string? GetCachedToken(string username)
    {
        if (!System.IO.File.Exists(_cachefileName))
            return null;

        try
        {
            string tokensArr = System.IO.File.ReadAllText(_cachefileName);
            JArray ja = JArray.Parse(tokensArr);
            foreach (JObject jo in ja)
            {
                if (jo["UserName"]?.ToString().ToLower() == username.ToLower() && (DateTime.Parse(jo["ExpirationDate"].ToString()) - DateTime.Now).TotalMinutes > 0)
                {
                    return jo["TokenString"]?.ToString();
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    private void SetCachedToken(string username, string token)
    {
        CachedTokenDTO ct = new CachedTokenDTO()
        {
            UserName = username,
            TokenString = token,
            ExpirationDate = DateTime.Now.AddMinutes(20)
        };

        JObject jo = JObject.Parse(JsonConvert.SerializeObject(ct));
        if (!System.IO.File.Exists(_cachefileName))
        {
            List<JObject> jos = new List<JObject>();
            jos.Add(jo);
            System.IO.File.WriteAllText(_cachefileName, JsonConvert.SerializeObject(jos.ToArray()));
        }
        else
        {
            try
            {
                string tokensArr = System.IO.File.ReadAllText(_cachefileName);
                JArray ja = JArray.Parse(tokensArr);
                bool joExist = false;
                foreach (JObject jo1 in ja)
                {
                    if (jo1["UserName"]?.ToString().ToLower() == username.ToLower())
                    {
                        jo1["ExpirationDate"] = ct.ExpirationDate;
                        jo1["TokenString"] = ct.TokenString;
                        joExist = true;
                        break;
                    }
                }

                List<JObject> jos = new List<JObject>();
                foreach (JObject jo1 in ja)
                    jos.Add(jo1);

                if (!joExist)
                    jos.Add(jo);

                System.IO.File.WriteAllText(_cachefileName, JsonConvert.SerializeObject(jos.ToArray()));
            }
            catch (Exception ex)
            {

            }
        }
    }

}
