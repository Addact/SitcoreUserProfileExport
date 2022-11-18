//using Sitecore.Pipelines;
using Sitecore.Pipelines;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Addact.Export.Infrastructure.Pipelines
{
    public class RegisterWebApiRoutes
    {
        public void Process(PipelineArgs args)
        {
            RouteTable.Routes.MapRoute("Addact.Api", "api/bghExportData/{action}", new
            {
                controller = "AddactExportData"
            });
        }


    }
}