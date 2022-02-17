using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Web;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Sushi.Mediakiwi.Data;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    internal class GoogleAuthListener
    {
        private readonly RequestDelegate _next;
        private const string UserID = nameof(GoogleSheetLogic);

        public GoogleAuthListener(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext httpContext)
        {
            var config = httpContext.RequestServices.GetService<IConfiguration>();
            var uri = httpContext.Request.GetDisplayUrl();

            var visMan = new VisitorManager(httpContext);
            var visitor = await visMan.SelectAsync();

            if (visitor?.ApplicationUserID.GetValueOrDefault(0) > 0)
            {
                IApplicationUser user = await ApplicationUser.SelectOneAsync(visitor.ApplicationUserID.Value);

                IAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow(
                          new GoogleAuthorizationCodeFlow.Initializer
                          {
                              ClientSecrets = new ClientSecrets
                              {
                                  ClientId = config["GoogleSheetsCredentials:client-id"],
                                  ClientSecret = config["GoogleSheetsCredentials:client-secret"]
                              },
                              DataStore = new GoogleTokenStore(user, UserID),
                              Scopes = new string[] { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile }
                          }
                     );

                var code = httpContext.Request.Query["code"].FirstOrDefault();
                if (code != null)
                {
                    var token = await flow.ExchangeCodeForTokenAsync(UserID, code, uri.Substring(0, uri.IndexOf("?")), CancellationToken.None);

                    // Extract the right state.
                    var oauthState = await AuthWebUtility.ExtracRedirectFromState(flow.DataStore, UserID, httpContext.Request.Query["state"]);
                    //await httpContext.Response.WriteAsync($"<script>window.open('{oauthState}','_self');</script>");
                    httpContext.Response.Redirect(oauthState);
                }
            }
        }
    }
}
