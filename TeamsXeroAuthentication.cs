using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;

namespace Valo.Teams.Authentication.Xero
{
    public static class TeamsXeroAuthentication
    {

        [FunctionName("TeamsXero")]
        public static async Task<HttpResponseMessage> TeamsXeroRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/xero/{xeroClientId}")] HttpRequest req,
            string xeroClientId,
            ILogger log)
        {

            /*
             *
             * 1. Teams app opens iFrame to /teams/xero/{xeroClientId} route
             * 2. /teams/xero/{xeroClientId} redirects to login.xero.com auth endpoint
             * 3. login.xero.com auth endpoint is instructed to redirect back to /teams/xero/response route
             * 4. /teams/xero/response 
             * 4. /teams/xero/response route provides response back to Teams via the SDK
             *
             */

            string[] requestedScopes = {"openid", "profile", "email", "payroll.payruns.read", "payroll.payslip.read"};

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

                        // Calculate the SHA256 hash of the input text. 
                        // Returns a promise that resolves to an ArrayBuffer
                        function sha256(plain) {{
                            const encoder = new TextEncoder();
                            const data = encoder.encode(plain);
                            return window.crypto.subtle.digest('SHA-256', data);
                        }}

                        // Base64-urlencodes the input string
                        function base64urlencode(str) {{
                            // Convert the ArrayBuffer to string using Uint8 array to conver to what btoa accepts.
                            // btoa accepts chars only within ascii 0-255 and base64 encodes them.
                            // Then convert the base64 encoded to base64url encoded
                            //   (replace + with -, replace / with _, trim trailing =)
                            return btoa(String.fromCharCode.apply(null, new Uint8Array(str)))
                                .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
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

                            var xeroClientId = document.location.pathname.substr(document.location.pathname.lastIndexOf('/') + 1);

                            // Create and store a new PKCE code_verifier (the plaintext random secret)
                            var codeVerifier = generateRandomString();

                            // Hash and base64-urlencode the secret to use as the challenge
                            var codeChallenge = await pkceChallengeFromVerifier(codeVerifier);

                            var redirectUri = document.location.origin + '/teams/xero/response/' + xeroClientId

                            window.localStorage.setItem('Xero:' + xeroClientId + ':CodeVerifier', codeVerifier);
                            window.localStorage.setItem('Xero:' + xeroClientId + ':CodeChallenge', codeChallenge);
                            window.localStorage.setItem('Xero:' + xeroClientId + ':RedirectUri', redirectUri);

                            
                            var loginUrl = 'https://login.xero.com/identity/connect/authorize?response_type=code'
                                + '&client_id=' + encodeURIComponent(xeroClientId) 
                                + '&redirect_uri=' + encodeURIComponent(redirectUri)
                                + '&scope={string.Join(" ", requestedScopes)}'
                                + '&state=123&code_challenge=' + encodeURIComponent(codeChallenge)
                                + '&code_challenge_method=S256'

                            window.location.assign(loginUrl);
                            log('redirecting...');
                        }}

                        // Return the base64-urlencoded sha256 hash for the PKCE challenge
                        async function pkceChallengeFromVerifier(v) {{
                            hashed = await sha256(v);
                            return base64urlencode(hashed);
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
            return response;
        }


        [FunctionName("TeamsXeroResponse")]
        public static async Task<HttpResponseMessage> TeamsXeroResponse(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/xero/response/{xeroClientId}")] HttpRequest req,
            string xeroClientId,
            ILogger log)
        {

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

                            var code = getQueryParameter('code');
                            var scopes = getQueryParameter('scopes');
                            var error = getQueryParameter('error');
                            var xeroClientId = document.location.pathname.substr(document.location.pathname.lastIndexOf('/') + 1);

                            if (code != '') {{

                                var success = {{
                                    grantType: 'authorization_code',
                                    clientId: xeroClientId,
                                    code: code, 
                                    scopes: scopes.split(' '),
                                    redirectUri: window.location.origin + window.location.pathname, 
                                    codeVerifier: window.localStorage.getItem('Xero:' + xeroClientId + ':CodeVerifier')
                                }}
                                
                                microsoftTeams.initialize();
                                window.setTimeout(function() {{ 
                                    microsoftTeams.authentication.notifySuccess(JSON.stringify(success));
                                }}, 500);

                                
                            }} 

                            if (error != '') {{

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

}
