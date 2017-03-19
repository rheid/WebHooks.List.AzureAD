using Microsoft.SharePoint.Client;
using SharePoint.WebHooks.Common;
using SharePoint.WebHooks.Common.Models;
using SharePoint.WebHooks.MVC.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;



namespace SharePoint.WebHooks.MVC.Controllers
{
    [Authorize]
    public class TestSetupController : Controller
    {
        private string accessToken = null;
        public async Task<ActionResult> Index()
        {
            using (var cc = ContextProvider.GetWebApplicationClientContext(Settings.SiteCollection))
            {
                if (cc != null)
                {
                    // Usage tracking
                    SampleUsageTracking(cc);

                    // Hookup event to capture access token
                    cc.ExecutingWebRequest += Cc_ExecutingWebRequest;

                    var lists = cc.Web.Lists;
                    cc.Load(cc.Web, w => w.Url);
                    cc.Load(lists, l => l.Include(p => p.Title, p => p.Id, p => p.Hidden));
                    cc.ExecuteQueryRetry();

                    WebHookManager webHookManager = new WebHookManager();

                    // Grab the current lists
                    List<SharePointList> modelLists = new List<SharePointList>();
                    List<SubscriptionModel> webHooks = new List<SubscriptionModel>();

                    foreach (var list in lists)
                    {
                        // Let's only take the hidden lists
                        if (!list.Hidden)
                        {
                            modelLists.Add(new SharePointList() { Title = list.Title, Id = list.Id });

                            // Grab the currently applied web hooks
                            var existingWebHooks = await webHookManager.GetListWebHooksAsync(cc.Web.Url, list.Id.ToString(), this.accessToken);

                            if (existingWebHooks.Value.Count > 0)
                            {
                                foreach (var existingWebHook in existingWebHooks.Value)
                                {
                                    webHooks.Add(existingWebHook);
                                }
                            }
                        }
                    }

                    // Prepare the data model
                    SharePointSiteModel sharePointSiteModel = new SharePointSiteModel();
                    sharePointSiteModel.SharePointSite = Settings.SiteCollection;
                    sharePointSiteModel.Lists = modelLists;
                    sharePointSiteModel.WebHooks = webHooks;
                    sharePointSiteModel.SelectedSharePointList = modelLists[0].Id;

                    return View();
                }
                else
                {
                    throw new Exception("Issue with obtaining a valid client context object, should not happen");
                }
            }
        }


        private void Cc_ExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            // grab the OAuth access token as we need this token in our REST calls
            this.accessToken = e.WebRequestExecutor.RequestHeaders.Get("Authorization").Replace("Bearer ", "");
        }

        /// <summary>
        /// We would love to understand which samples are populair to prioritize work
        /// </summary>
        /// <param name="cc">ClientContext object</param>
        private void SampleUsageTracking(ClientContext cc)
        {
            cc.ClientTag = "SPDev:WebHooks";
            cc.Load(cc.Web, p => p.Description);
            cc.ExecuteQuery();
        }



        // GET: TestSetup/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: TestSetup/Create
        public ActionResult Create()
        {
            return View();
        }

    }
}
