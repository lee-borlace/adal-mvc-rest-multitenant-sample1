using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using WebApplication1.Models;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Text;
using Microsoft.SharePoint.Client;

namespace WebApplication1.Controllers
{
    [Authorize]
    public class UserProfileController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();
        private string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        private string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private string graphResourceID = "https://graph.windows.net";

        private string SHAREPOINT_BASE_URL = "https://lee79.sharepoint.com";
        private string TASKS_SITE_URL1 = "https://lee79.sharepoint.com/sites/dev/TestTeamSiteWithTasks";

        // GET: UserProfile
        public async Task<ActionResult> Index()
        {
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            try
            {
                Uri servicePointUri = new Uri(graphResourceID);
                Uri serviceRoot = new Uri(servicePointUri, tenantID);
                ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                      async () => await GetTokenForApplication());

                // use the token for querying the graph to get the user details

                var result = await activeDirectoryClient.Users
                    .Where(u => u.ObjectId.Equals(userObjectID))
                    .ExecuteAsync();
                IUser user = result.CurrentPage.ToList().First();


                var sb = new StringBuilder();


                var token = await GetTokenForGraph();
                var plannerTasks = await GetPlannerTasks(token);

                foreach (var task in plannerTasks)
                {
                    sb.Append(task.Title);
                    sb.Append(",");
                }

                ViewBag.PlannerTasks = sb.ToString();

                token = await GetTokenForOutlook();
                var outlookTasks = await GetOutlookTasks(token);

                sb.Clear();

                foreach (var task in outlookTasks)
                {
                    sb.Append(task.Title);
                    sb.Append(",");
                }

                ViewBag.OutlookTasks = sb.ToString();

                var sharePointTasks = await GetSharePointTasks();

                sb.Clear();

                foreach (var task in sharePointTasks)
                {
                    sb.Append(task.Title);
                    sb.Append(",");
                }

                ViewBag.SharePointTasks = sb.ToString();


                return View(user);
            }
            catch (AdalException)
            {
                // Return to error page.
                return View("Error");
            }
            // if the above failed, the user needs to explicitly re-authenticate for the app to obtain the required token
            catch (Exception)
            {
                return View("Relogin");
            }
        }

        public void RefreshSession()
        {
            HttpContext.GetOwinContext().Authentication.Challenge(
                new AuthenticationProperties { RedirectUri = "/UserProfile" },
                OpenIdConnectAuthenticationDefaults.AuthenticationType);
        }

        public async Task<string> GetTokenForApplication()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
            // initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            AuthenticationContext authenticationContext = new AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signedInUserID));
            AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenSilentAsync(graphResourceID, clientcred, new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
            return authenticationResult.AccessToken;
        }


        public async Task<string> GetTokenForGraph()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
            // initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            AuthenticationContext authenticationContext = new AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signedInUserID));
            AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenSilentAsync("https://graph.microsoft.com", clientcred, new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
            return authenticationResult.AccessToken;
        }



        public async Task<string> GetTokenForOutlook()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
            // initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            AuthenticationContext authenticationContext = new AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signedInUserID));
            AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenSilentAsync("https://outlook.office.com", clientcred, new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
            return authenticationResult.AccessToken;
        }


        public string GetTokenForSharePoint()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            // get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
            // initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            AuthenticationContext authenticationContext = new AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signedInUserID));

            // TODO : Make the resource here dynamic as it will need to match the specifics of the tenant.
            AuthenticationResult authenticationResult = authenticationContext.AcquireTokenSilent(SHAREPOINT_BASE_URL, clientcred, new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
            return authenticationResult.AccessToken;
        }


        public async Task<List<TaskBase>> GetPlannerTasks(string accessToken)
        {
            string endpoint = "https://graph.microsoft.com/beta/me/tasks";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        var retVal = new List<TaskBase>();

                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            var rawResponse = await response.Content.ReadAsStringAsync();

                            dynamic results = JsonConvert.DeserializeObject<dynamic>(rawResponse);

                            foreach (var task in results.value)
                            {
                                retVal.Add(new TaskBase() { Title = task.title });
                            }
                        }

                        return retVal;
                    }
                }
            }
        }



        public async Task<List<TaskBase>> GetOutlookTasks(string accessToken)
        {
            string endpoint = "https://outlook.office.com/api/v2.0/me/tasks";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        var retVal = new List<TaskBase>();

                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            var rawResponse = await response.Content.ReadAsStringAsync();

                            dynamic results = JsonConvert.DeserializeObject<dynamic>(rawResponse);

                            foreach (var task in results.value)
                            {
                                retVal.Add(new TaskBase() { Title = task.Subject });
                            }
                        }

                        return retVal;
                    }
                }
            }
        }


        public async Task<List<TaskBase>> GetSharePointTasks()
        {
            var retVal = new List<TaskBase>();

            using (var context = new ClientContext(TASKS_SITE_URL1))
            {
                context.ExecutingWebRequest += ctx_ExecutingWebRequest;

                List tasksList = context.Web.Lists.GetByTitle("Tasks");
                var listItems = tasksList.GetItems(CamlQuery.CreateAllItemsQuery());

                context.Load(listItems);
                context.ExecuteQuery();

                foreach (ListItem item in listItems)
                {
                    retVal.Add(new TaskBase() { Title = item["Title"].ToString() });
                }
            }



            return retVal;
        }


        async void ctx_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            var token = GetTokenForSharePoint();

            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + token;
        }

    }
}
