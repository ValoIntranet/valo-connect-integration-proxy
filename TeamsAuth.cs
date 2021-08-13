using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.Text;  
using System.Security.Cryptography;

namespace Valo.Integration.Xero
{
    public static class TeamsAuth
    {

        [FunctionName("Teams")]
        public static async Task<HttpResponseMessage> TeamsRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "teams/{flowType}/{xeroClientId}")] HttpRequest req,
            string flowType,
            string xeroClientId,
            ILogger log)
        {

            /*
             *
             * 1. Teams app opens iFrame to /teams-open-window route
             * 2. /teams-open-window opens a JS window to login.xero.com auth endpoint
             * 3. login.xero.com auth endpoint is instructed to redirect back to /xero-response route
             * 4. /xero-response route provides response back to the opening iFrame's context, via a localStorage property
             * 5. Opening iFrame /teams-open-window waits for localStorage property with success or failure
             * 6. When localStorage property is received either success or failure, message is reported to the iFrame's parent window via window.parent.postMessage(message, '*');
             *
             */

            // string xeroClientId = $"3C053FCDE1F542AA910C17421E036CA9";
            // string localStorageKey = $"xero_{xeroClientId}";
            string redirectUrl = $"{req.Scheme}://{req.Host.Value}/xero-code-response/{xeroClientId}";

            // string xeroCodeVerifier = CreateString(128);
            // string hashedVerifier = ComputeSha256Hash(xeroCodeVerifier);
            // string xeroCodeChallenge = EncodeBase64(hashedVerifier);

            // log.LogInformation($"xeroCodeVerifier: {xeroCodeVerifier}");
            // log.LogInformation($"hashedVerifier: {hashedVerifier}");
            // log.LogInformation($"xeroCodeChallenge: {xeroCodeChallenge}");
            

            // string loginUrl = $"https://login.xero.com/identity/connect/authorize?response_type=code"
            //     + $"&client_id={WebUtility.UrlEncode(xeroClientId)}&redirect_uri={WebUtility.UrlEncode(redirectUrl)}&scope=openid profile email payroll.payslip.read"
            //     + $"&state=123&code_challenge={{codeChallenge}}&code_challenge_method=S256";


            // string injectFlow = $"window.location.assign('{loginUrl}');";
            // switch (flowType) {
            //     case "open-window": injectFlow = $"window.open('{loginUrl}', 'xero_{xeroClientId}', 'width=400,height=600,location,scrollbars');"; break;
            //     case "iframe": injectFlow = $@"var iframeContainer = document.getElementById('iframeContainer');
            //                 iframe.src = '{loginUrl}';
            //                 iframe.style = 'width: 100%; height: 100%;';
            //                 iframe.sandbox = 'allow-forms allow-modals allow-popups allow-popups-to-escape-sandbox allow-pointer-lock allow-scripts allow-same-origin allow-downloads allow-top-navigation'
            //                 iframeContainer.appendChild(iframe);"; break;
            // }

            string[] requestedScopes = {"openid", "profile", "email", "payroll.payruns.read", "payroll.payslip.read"};

            string html = $@"<html>
                <head>
                    <script type=""text/javascript"">

                        // window.addEventListener('message', function(e) {{
                        //     // Get the sent data
                        //     const data = e.data;
                            
                        //     // If you encode the message in JSON before sending them, 
                        //     // then decode here
                        //     // const decoded = JSON.parse(data);

                        //     log(data);

                        // }});

                        // function checkLocalStorage() {{
                        //     log(`checkLocalStorage() {xeroClientId}`);
                        //     var value = window.localStorage.getItem('Xero:{xeroClientId}');
                        //     if (value) {{
                        //         const message = JSON.stringify({{
                        //             data: value,
                        //             date: Date.now(),
                        //         }});
                        //         log(`posting message: ${{message}}`);
                        //         window.parent.postMessage(message, '*');
                        //         localStorage.removeItem('Xero:{xeroClientId}');
                        //     }}
                        //     else {{
                        //         window.setTimeout(checkLocalStorage, 500);
                        //     }}
                        // }}

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

                            openXeroAuthenticationWindow();
                        
                        }});

                        async function openXeroAuthenticationWindow() {{

                            var xeroClientId = document.location.pathname.substr(document.location.pathname.lastIndexOf('/') + 1);

                            // Create and store a new PKCE code_verifier (the plaintext random secret)
                            var codeVerifier = generateRandomString();

                            // Hash and base64-urlencode the secret to use as the challenge
                            var codeChallenge = await pkceChallengeFromVerifier(codeVerifier);

                            var redirectUri = document.location.origin + '/xero-code-response/' + xeroClientId

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
                    <button onclick=""javascript:openXeroAuthenticationWindow();"" value=""Login"" />
                    <div id=""log""></div>
                </body>
            </html>";


            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StringContent(html);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            response.Content.Headers.Add("Content-Security-Policy", "default-src 'self' 'unsafe-inline' authorize.xero.com login.xero.com sorry.xero.com; connect-src 'self' 'unsafe-inline' authorize.xero.com login.xero.com;");
            return response;
        }


        [FunctionName("XeroCodeResponse")]
        public static async Task<HttpResponseMessage> XeroCodeResponseRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "xero-code-response/{xeroClientId}")] HttpRequest req,
            string xeroClientId,
            ILogger log)
        {

            string redirectUrl = $"{req.Scheme}://{req.Host.Value}/xero-code-response/{xeroClientId}";
            
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

                                
                                // window.localStorage.setItem('Xero:{xeroClientId}:Code:QueryString', window.location.search);
                                // window.localStorage.setItem('Xero:{xeroClientId}:Code', code);

                                // var form = document.forms[0];
                                // form.code.value = code;
                                // form.code_verifier.value = window.localStorage.getItem('Xero:{xeroClientId}:CodeVerifier');
                                // //form.submit();

                            }} 

                            if (error != '') {{

                                microsoftTeams.initialize();
                                window.setTimeout(function() {{ 
                                    microsoftTeams.authentication.notifyFailure(error);
                                }}, 500);

                            }}
                            // else {{

                            //     var token = {{
                            //         accessToken: getQueryParameter('access_token'),
                            //         idToken: getQueryParameter('id_token'),
                            //         expiresIn: getQueryParameter('expires_in'),
                            //         tokenType: getQueryParameter('token_type'),
                            //         refreshToken: getQueryParameter('refresh_token')
                            //     }};
                            //     window.localStorage.setItem('Xero:{xeroClientId}:Token:accessToken', token.acessToken);
                            //     window.localStorage.setItem('Xero:{xeroClientId}:Token:idToken', token.idToken);
                            //     window.localStorage.setItem('Xero:{xeroClientId}:Token:expiresIn', token.expireIn);
                            //     window.localStorage.setItem('Xero:{xeroClientId}:Token:tokenType', token.tokenType);
                            //     window.localStorage.setItem('Xero:{xeroClientId}:Token:refreshToken', token.refreshToken);
                                
                            //     microsoftTeams.initialize();
                            //     window.setTimeout(function() {{ 
                            //         microsoftTeams.authentication.notifySuccess(JSON.stringify(token));
                            //     }}, 500);

                            // }}
                            
                        }});

                    </script>
                </head>
                <body>
                    <form method=""post"" action=""https://identity.xero.com/connect/token"">
                        <input type=""hidden"" name=""grant_type"" value=""authorization_code"" />
                        <input type=""hidden"" name=""client_id"" value=""{xeroClientId}"" />
                        <input type=""hidden"" name=""code"" value="""" />
                        <input type=""hidden"" name=""redirect_uri"" value=""{redirectUrl}"" />
                        <input type=""hidden"" name=""code_verifier"" value="""" />
                        <input type=""submit"" value=""Continue"" />
                </body>
            </html>";


            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StringContent(html);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            return response;
        }

        [FunctionName("XeroAccessTokenResponse")]
        public static async Task<HttpResponseMessage> XeroAccessTokenResponseRun(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "xero-identity/{localStorageKey}")] HttpRequest req,
            string localStorageKey,
            ILogger log)
        {

            string html = $@"<html>
                <head>
                    <script src=""https://statics.teams.cdn.office.net/sdk/v1.10.0/js/MicrosoftTeams.min.js"" integrity=""sha384-6oUzHUqESdbT3hNPDDZUa/OunUj5SoxuMXNek1Dwe6AmChzqc6EJhjVrJ93DY/Bv"" crossorigin=""anonymous""></script>
                    <script type=""text/javascript"">
                    
                        window.addEventListener('load', (event) => {{

                            const message = '{req.QueryString}';
                            window.localStorage.setItem('{localStorageKey}', message);
                            
                            microsoftTeams.initialize();
                            window.setTimeout(function() {{ 
                                microsoftTeams.authentication.notifySuccess(message);
                            }}, 500);

                            // microsoftTeams.authentication.notifySuccess({{idToken: message}});

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
        static Random rd = new Random();
        internal static string CreateString(int stringLength)
        {
            const string allowedChars = "ABCDEFGHJKLMNOPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz0123456789";
            char[] chars = new char[stringLength];

            for (int i = 0; i < stringLength; i++)
            {
                chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
            }

            return new string(chars);
        }

        internal static string ComputeSha256Hash(string rawData)  
        {  
            // Create a SHA256   
            using (SHA256 sha256Hash = SHA256.Create())  
            {  
                // ComputeHash - returns byte array  
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(rawData));  
  
                // Convert byte array to a string   
                StringBuilder builder = new StringBuilder();  
                for (int i = 0; i < bytes.Length; i++)  
                {  
                    builder.Append(bytes[i].ToString("x2"));  
                }  
                return builder.ToString();  
            }  
        }

        internal static string EncodeBase64(this string value)
        {
            var valueBytes = Encoding.UTF8.GetBytes(value);
            return Convert.ToBase64String(valueBytes);
        }
    }

}
