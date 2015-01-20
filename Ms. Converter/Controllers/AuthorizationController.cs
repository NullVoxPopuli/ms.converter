using Ms.Converter.Models;
using OAuth2;
using OAuth2.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Mvc;

namespace Ms.Converter.Controllers
{
    public class AuthorizationController : Controller
    {
        private readonly AuthorizationRoot authorizationRoot;

        private const string ProviderNameKey = "GoogleAuth";

        private string GoogleAuth
        {
            get { return (string)Session[ProviderNameKey]; }
            set { Session[ProviderNameKey] = value; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AuthorizationController"/> class.
        /// </summary>
        /// <param name="authorizationRoot">The authorization manager.</param>
        public AuthorizationController(AuthorizationRoot authorizationRoot)
        {
            this.authorizationRoot = authorizationRoot;
        }

        public AuthorizationController() : this(new AuthorizationRoot()){

        }

        /// <summary>
        /// Renders home page with login link.
        /// </summary>
        public ActionResult Index()
        {
            var model = authorizationRoot.Clients.Select(client => new LoginInfo
            {
                ProviderName = client.Name
            });
            return View(model);
        }


        /// <summary>
        /// Redirect to login url of selected provider.
        /// </summary>        
        /// <summary>
        /// Renders information received from authentication service.
        /// </summary>
        public ActionResult Auth()
        {
            return View(GetClient().GetUserInfo(Request.QueryString));
        }

        private IClient GetClient()
        {
            return authorizationRoot.Clients.First(c => c.Name == GoogleAuth);
        }
    }
}
