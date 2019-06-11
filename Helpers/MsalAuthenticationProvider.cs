using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System.Linq;

namespace MSGraphTest
{
// This class encapsulates the details of getting a token from MSAL and exposes it via the
// IAuthenticationProvider interface so that GraphServiceClient or AuthHandler can use it.
// A significantly enhanced version of this class will in the future be available from
// the GraphSDK team.  It will supports all the types of Client Application as defined by MSAL.
public class MsalAuthenticationProvider : IAuthenticationProvider
{
private ConfidentialClientApplication _clientApplication;
private string[] _scopes;

public MsalAuthenticationProvider(ConfidentialClientApplication clientApplication, string[] scopes) {
_clientApplication = clientApplication;
_scopes = scopes;
}

/// &lt;summary&gt;
/// Update HttpRequestMessage with credentials
/// &lt;/summary&gt;
public async Task AuthenticateRequestAsync(HttpRequestMessage request)
{
var token = await GetTokenAsync();
request.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
}

/// &lt;summary&gt;
/// Acquire Token
/// &lt;/summary&gt;
public async Task<string> GetTokenAsync()
{
AuthenticationResult authResult = null;
authResult = await _clientApplication.AcquireTokenForClientAsync(_scopes);
return authResult.AccessToken;
}
}
}