using System.IO;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
namespace Valo.Teams.Authentication.Zendesk
{
    public static class TeamsZendeskAuthentication
    {

        [FunctionName("TeamsZendesk")]
        public static async Task<HttpResponseMessage> TeamsZendeskRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/zendesk/{zendeskInstance}/{zendeskClientId}")] HttpRequest req,
            string zendeskInstance,
            string zendeskClientId,
            ILogger log)
        {

            /*
             *
             * 1. Teams app opens iFrame to teams/zendesk/{zendeskInstance}/{zendeskClientId} route
             * 2. teams/zendesk/{zendeskInstance}/{zendeskClientId} redirects to {subdomain}.zendesk.com oauth endpoint
             * 3. zendesk oauth endpoint is instructed to redirect back to teams/response/zendesk/{zendeskInstance}/{zendeskClientId} route
             * 4. teams/response/zendesk/{zendeskInstance}/{zendeskClientId} route provides response back to Teams via the SDK
             *
             */

            string[] requestedScopes = { "read" };

            string html = $@"<html>
                <head>
                    <script type=""text/javascript"">

                        //////////////////////////////////////////////////////////////////////
                        // PKCE HELPER FUNCTIONS

                        // Generate a secure random string using the browser crypto functions
                        function generateRandomString() {{
                            var array = new Uint32Array(28);
                            window.crypto.getRandomValues(array);
                            return Array.from(array, dec => ('0' + dec.toString(16)).substr(-2)).join('');
                        }}

                        window.addEventListener('load', (event) => {{

                            window.log = function(message) {{ 
                                var logContainer = document.getElementById(""log"");
                                if (logContainer) logContainer.innerHTML += `${{message}}<br/>`; 
                                console.log(message);  
                            }};

                            redirectToAuthentication();
                        
                        }});

                        async function redirectToAuthentication() {{

                            var splitPath = document.location.pathname.split('/');
                            var zendeskInstance = splitPath[splitPath.length - 2];
                            var zendeskClientId = splitPath[splitPath.length - 1];

                            // Create and store a new PKCE code_verifier (the plaintext random secret)
                            var state = generateRandomString();

                            var redirectUri = `${{document.location.origin}}/teams/response/zendesk/${{zendeskInstance}}/${{zendeskClientId}}`;
 
                            window.localStorage.setItem('ZENDESK:' + zendeskInstance + ':' + zendeskClientId + ':State', state);
                            
                            var loginUrl = `https://${{zendeskInstance}}.zendesk.com/oauth/authorizations/new?response_type=code`
                                + '&client_id=' + encodeURIComponent(zendeskClientId) 
                                + '&redirect_uri=' + encodeURIComponent(redirectUri)
                                + '&scope=' + encodeURIComponent('{string.Join(" ", requestedScopes)}')
                                + '&state=' + encodeURIComponent(state);

                            window.location.assign(loginUrl);
                            log('redirecting...');
                        }}

                    </script>
                </head>
                <body>
                    <div id=""iframeContainer""></div>
                    <button onclick=""javascript:redirectToAuthentication();"" value=""Login"" />
                    <div id=""log""></div>
                </body>
            </html>";


            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StringContent(html);
            response.Content.Headers.Add("Content-Security-Policy", "default-src 'self' 'unsafe-inline' https://static.zdassets.com https://ekr.zdassets.com https://{{zendeskInstance}}.zendesk.com https://*.zopim.com wss://{{zendeskInstance}}.zendesk.com wss://*.zopim.com; connect-src 'self' 'unsafe-inline' https://static.zdassets.com https://ekr.zdassets.com https://{{zendeskInstance}}.zendesk.com https://*.zopim.com wss://{{zendeskInstance}}.zendesk.com wss://*.zopim.com;");
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            return response;
        }

        [FunctionName("TeamsZendeskResponse")]
        public static async Task<HttpResponseMessage> TeamsZendeskResponse(
           [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/response/zendesk/{zendeskInstance}/{zendeskClientId}")] HttpRequest req,
           ExecutionContext context,
           string zendeskInstance,
           string zendeskClientId,
           ILogger log)
        {

            string error = string.Empty;

            if (!req.Query.ContainsKey("code"))
            {
                error = "Missing response value: code";
            }

            var config = new ConfigurationBuilder()
                .SetBasePath(context.FunctionAppDirectory)
                // This gives you access to your application settings 
                // in your local development environment
                .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
                // This is what actually gets you the 
                // application settings in Azure
                .AddEnvironmentVariables()
                .Build();

            string zendeskClientSecret = config.GetValue<string>($"ZENDESK:{zendeskInstance}:{zendeskClientId}:ClientSecret");

            if (string.IsNullOrEmpty(zendeskClientSecret))
            {
                error = $"Check client secret for instance / client id: {zendeskInstance} / {zendeskClientId}";
            }

            if (string.IsNullOrEmpty(error))
            {

                string code = req.Query["code"];
                string redirectUrl = $"{req.Scheme}://{req.Host.Value}/teams/response/zendesk/{zendeskInstance}/{zendeskClientId}";

                using (HttpClient zendeskTokenRequest = new HttpClient())
                {

                    string body = "grant_type=authorization_code"
                                + $"&code={WebUtility.UrlEncode(code)}"
                                + $"&redirect_uri={WebUtility.UrlEncode(redirectUrl)}"
                                + $"&client_id={zendeskClientId}"
                                + $"&client_secret={WebUtility.UrlEncode(zendeskClientSecret)}";

                    StringContent requestContent = new StringContent(body);
                    requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                    HttpResponseMessage zendeskResponse = await zendeskTokenRequest.PostAsync($"https://{zendeskInstance}.zendesk.com/oauth/tokens", requestContent);

                    string responseContent = $"var response = {await zendeskResponse.Content.ReadAsStringAsync()};";

                    string html = $@"<html>
                        <head>
                            <script src=""https://statics.teams.cdn.office.net/sdk/v1.10.0/js/MicrosoftTeams.min.js"" integrity=""sha384-6oUzHUqESdbT3hNPDDZUa/OunUj5SoxuMXNek1Dwe6AmChzqc6EJhjVrJ93DY/Bv"" crossorigin=""anonymous""></script>
                            <script type=""text/javascript"">
                            
                                function getQueryParameter(name) {{
                                    name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
                                    var regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
                                    var results = regex.exec(window.location.search);
                                    return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
                                }};

                                window.addEventListener('load', (event) => {{

                                    {responseContent}
                                    var splitPath = document.location.pathname.split('/');
                                    var zendeskInstance = splitPath[splitPath.length - 2];
                                    var zendeskClientId = splitPath[splitPath.length - 1];

                                    var returnedState = getQueryParameter('state');
                                    var expectedState = window.localStorage.getItem('ZENDESK:' + zendeskInstance + ':' + zendeskClientId + ':State');

                                    var error;

                                    if (expectedState !== returnedState) {{

                                        error = `Unexpected state received from Zendesk instance ${{zendeskInstance}}`;

                                    }}

                                    if (!error && '{zendeskResponse.IsSuccessStatusCode}'.toLowerCase() === 'true') {{

                                        var success = {{
                                            client_id: zendeskClientId,
                                            ...response
                                        }}
                                        
                                        microsoftTeams.initialize();
                                        window.setTimeout(function() {{ 
                                            microsoftTeams.authentication.notifySuccess(JSON.stringify(success));
                                        }}, 500);

                                    }} 

                                    else {{

                                        var error = {{
                                            client_id: zendeskClientId,
                                            ...response
                                        }}

                                        microsoftTeams.initialize();
                                        window.setTimeout(function() {{ 
                                            microsoftTeams.authentication.notifyFailure(JSON.stringify(errors));
                                        }}, 500);

                                    }}
                                    
                                }});

                            </script>
                        </head>
                        <body>
                        </body>
                    </html>";


                    HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                    response.Content = new StringContent(html);
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                    return response;

                }
            }

            else
            {
                string html = $@"<html>
                    <head>
                        <script src=""https://statics.teams.cdn.office.net/sdk/v1.10.0/js/MicrosoftTeams.min.js"" integrity=""sha384-6oUzHUqESdbT3hNPDDZUa/OunUj5SoxuMXNek1Dwe6AmChzqc6EJhjVrJ93DY/Bv"" crossorigin=""anonymous""></script>
                        <script type=""text/javascript"">

                            window.addEventListener('load', (event) => {{
                                microsoftTeams.initialize();
                                window.setTimeout(function() {{ 
                                    microsoftTeams.authentication.notifyFailure(JSON.stringify({{ error: 'auth_error', error_description: '{error}' }}));
                                }}, 500);
                                
                            }});

                        </script>
                    </head>
                    <body>
                    </body>
                </html>";

                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                response.Content = new StringContent(html);
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                return response;

            }

        }
    }
}
