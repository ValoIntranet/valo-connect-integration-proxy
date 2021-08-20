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

namespace Valo.Teams.Authentication.ServiceNow
{
    public static class TeamsAuth
    {

        [FunctionName("TeamsServiceNow")]
        public static async Task<HttpResponseMessage> TeamsServiceNowRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/snow/{snowInstance}/{snowClientId}")] HttpRequest req,
            string snowInstance,
            string snowClientId,
            ILogger log)
        {

            /*
             *
             * 1. Teams app opens iFrame to /teams/snow/{snowClientId} route
             * 2. /teams/snow/{snowClientId} redirects to login.xero.com auth endpoint
             * 3. login.xero.com auth endpoint is instructed to redirect back to /teams/snow/response route
             * 4. /teams/snow/response 
             * 4. /teams/snow/response route provides response back to Teams via the SDK
             *
             */

            string[] requestedScopes = {"useraccount"};

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
                            var snowInstance = splitPath[splitPath.length - 2];
                            var snowClientId = splitPath[splitPath.length - 1];

                            // Create and store a new PKCE code_verifier (the plaintext random secret)
                            var state = generateRandomString();

                            var redirectUri = `${{document.location.origin}}/teams/response/snow/${{snowInstance}}/${{snowClientId}}`;

                            window.localStorage.setItem('SNOW:' + snowInstance + ':' + snowClientId + ':State', state);
                            window.localStorage.setItem('SNOW:' + snowInstance + ':' + snowClientId + ':RedirectUri', redirectUri);
                            
                            var loginUrl = `https://${{snowInstance}}.service-now.com/oauth_auth.do?response_type=code`
                                + '&client_id=' + encodeURIComponent(snowClientId) 
                                + '&redirect_uri=' + encodeURIComponent(redirectUri)
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
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            response.Content.Headers.Add("Content-Security-Policy", "default-src 'self' 'unsafe-inline' authorize.xero.com login.xero.com sorry.xero.com; connect-src 'self' 'unsafe-inline' authorize.xero.com login.xero.com;");
            return response;
        }


        [FunctionName("TeamsServiceNowResponse")]
        public static async Task<HttpResponseMessage> TeamsServiceNowResponse(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/response/snow/{snowInstance}/{snowClientId}")] HttpRequest req,
            ExecutionContext context,
            string snowInstance,
            string snowClientId,
            ILogger log)
        {

            string error = string.Empty;

            if (!req.Query.ContainsKey("code")) {
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

            string snowClientSecret = config.GetValue<string>($"SNOW:{snowInstance}:{snowClientId}:ClientSecret");

            if (string.IsNullOrEmpty(snowClientSecret)) {
                error = $"Check client secret for instance / client id: {snowInstance} / {snowClientId}";
            }

            if (string.IsNullOrEmpty(error)) {
                
                string code = req.Query["code"];
                string redirectUrl = $"{req.Scheme}://{req.Host.Value}/teams/response/snow/{snowInstance}/{snowClientId}";

                using (HttpClient snowTokenRequest = new HttpClient()) {

                    string body = "grant_type=authorization_code"
                                + $"&code={WebUtility.UrlEncode(code)}"
                                + $"&redirect_uri={WebUtility.UrlEncode(redirectUrl)}"
                                + $"&client_id={snowClientId}"
                                + $"&client_secret={WebUtility.UrlEncode(snowClientSecret)}";

                    StringContent requestContent = new StringContent(body);
                    requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                    HttpResponseMessage snowResponse = await snowTokenRequest.PostAsync($"https://{snowInstance}.service-now.com/oauth_token.do", requestContent);

                    string responseContent = snowResponse.IsSuccessStatusCode ? 
                                                $"var response = {await snowResponse.Content.ReadAsStringAsync()};"
                                                : $"var response = {await snowResponse.Content.ReadAsStringAsync()};";

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
                                    var snowInstance = splitPath[splitPath.length - 2];
                                    var snowClientId = splitPath[splitPath.length - 1];

                                    var returnedState = getQueryParameter('state');
                                    var expectedState = window.localStorage.getItem('SNOW:' + snowInstance + ':' + snowClientId + ':State');

                                    var error;

                                    if (expectedState !== returnedState) {{

                                        error = `Unexpected state received from SNOW instance ${{snowInstance}}`;

                                    }}

                                    if (!error && '{snowResponse.IsSuccessStatusCode}'.toLowerCase() === 'true') {{

                                        var success = {{
                                            client_id: snowClientId,
                                            ...response
                                        }}
                                        
                                        microsoftTeams.initialize();
                                        window.setTimeout(function() {{ 
                                            microsoftTeams.authentication.notifySuccess(JSON.stringify(success));
                                        }}, 500);

                                    }} 

                                    else {{

                                        var error = {{
                                            client_id: snowClientId,
                                            ...response
                                        }}

                                        microsoftTeams.initialize();
                                        window.setTimeout(function() {{ 
                                            microsoftTeams.authentication.notifyFailure(error);
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
            else {

                string html = $@"<html>
                    <head>
                        <script src=""https://statics.teams.cdn.office.net/sdk/v1.10.0/js/MicrosoftTeams.min.js"" integrity=""sha384-6oUzHUqESdbT3hNPDDZUa/OunUj5SoxuMXNek1Dwe6AmChzqc6EJhjVrJ93DY/Bv"" crossorigin=""anonymous""></script>
                        <script type=""text/javascript"">

                            window.addEventListener('load', (event) => {{

                                var splitPath = document.location.pathname.split('/');
                                var snowInstance = splitPath[splitPath.length - 2];
                                var snowClientId = splitPath[splitPath.length - 1];

                                microsoftTeams.initialize();
                                window.setTimeout(function() {{ 
                                    microsoftTeams.authentication.notifyFailure({{ error: 'auth_error', error_description: '{error}' }});
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

        [FunctionName("TeamsServiceNowRefresh")]
        public static async Task<HttpResponseMessage> TeamsServiceNowRefresh(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "teams/refresh/snow/{snowInstance}/{snowClientId}")] HttpRequest req,
            ExecutionContext context,
            string snowInstance,
            string snowClientId,
            ILogger log)
        {

            string[] requestBody = (await new StreamReader(req.Body).ReadToEndAsync()).Split("=");
            string key = requestBody[0];
            string value = requestBody[1];

            if (key.ToLower() == "refresh_token") {
                
                var config = new ConfigurationBuilder()
                    .SetBasePath(context.FunctionAppDirectory)
                        // This gives you access to your application settings 
                        // in your local development environment
                    .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true) 
                        // This is what actually gets you the 
                        // application settings in Azure
                    .AddEnvironmentVariables() 
                    .Build();

                string snowClientSecret = config.GetValue<string>($"SNOW:{snowInstance}:{snowClientId}:ClientSecret");

                using (HttpClient snowTokenRequest = new HttpClient()) {

                    string body = "grant_type=refresh_token"
                                + $"&refresh_token={WebUtility.UrlEncode(value)}"
                                + $"&client_id={snowClientId}"
                                + $"&client_secret={WebUtility.UrlEncode(snowClientSecret)}";

                    StringContent requestContent = new StringContent(body);
                    requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                    HttpResponseMessage snowResponse = await snowTokenRequest.PostAsync($"https://{snowInstance}.service-now.com/oauth_token.do", requestContent);
                    string responseContent = await snowResponse.Content.ReadAsStringAsync();
                     

                    if (responseContent.IndexOf("Instance Hibernating page") > -1) {

                        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.ServiceUnavailable);
                        response.Content = new StringContent($"{{ \"error\": \"Hibernating\", \"error_description\": \"The instance of ServiceNow {snowInstance} is currently in a hibernation state\" }}");
                        response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                        return response;

                    }
                    else {

                        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                        response.Content = new StringContent(responseContent);
                        response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                        return response;

                    }
                }

            }
            else {

                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.BadRequest);
                response.Content = new StringContent("{ error: \"\" }");
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                return response;

            }



        }

    }

}
