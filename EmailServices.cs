using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.Net;

namespace OpsAccountingWF
{
    public class EmailServices
    {
        public bool UpdateEmailCategory(string status, ExchangeService services,string item)
        {
            EmailMessage message = EmailMessage.Bind(services, item, new PropertySet(EmailMessageSchema.Categories));

            if(message.Categories == null)
            {
                message.Categories.Add(status);
                message.Update(ConflictResolutionMode.AutoResolve);
                return true;
            }
            else if (!message.Categories.Contains(status))
            {
                if (message.Categories.Contains("Pending") || message.Categories.Contains("Already Processed") || message.Categories.Contains("Completed") || message.Categories.Contains("Query"))
                {
                    string cat = message.Categories.ToString();
                    message.Categories.Remove(cat);
                    message.Categories.Add(status);
                }
                else
                {
                    message.Categories.Add(status);
                }
                message.Update(ConflictResolutionMode.AutoResolve);
                return true;
            }
            return false;
        }
        public  ExchangeService ConnectService(string emailaddress, string emailpassword, string url)
        {
            ExchangeService service = null;

            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            try
            {
                service = new ExchangeService(ExchangeVersion.Exchange2013)
                {
                    //Credentials = new WebCredentials(emailaddress, emailpassword,"sharedservices"),
                    //Credentials = new OAuthCredentials()
                    Credentials = new WebCredentials(emailaddress, emailpassword),
                    UseDefaultCredentials = false,

                    Url = new Uri(url)
                };
                

                // service.WebProxy = new WebProxy(new Uri("http://10.38.193.48:8080"));
                //service.Credentials = new WebCredentials("sharedservices/u310475", "Vnl@123@");
                return service;
            }
            catch (Exception ex) { Console.Write(ex.ToString()); return null; }

        }

        public string ConnectService(string MailBox)
        {
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            string cliendId = "f64cb9ac-f813-4488-9f57-34590d700eb8";
            string tenantId = "19470c39-fb57-4cea-8f8d-7c581679a164";
            string redirectUri = "http://localhost";

            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = cliendId,
                TenantId = tenantId,
                RedirectUri = redirectUri,
            };

            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

            // Make the interactive token request
            var authResult = pca.AcquireTokenInteractive(ewsScopes)
                .WithUseEmbeddedWebView(false).ExecuteAsync().Result;

            var TokenResult = authResult.AccessToken;

            //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            //service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");
            //service.Credentials = new OAuthCredentials(TokenResult);
            //service.HttpHeaders.Add("X-AnchorMailbox", MailBox);


            return TokenResult;

        }
    }
}
