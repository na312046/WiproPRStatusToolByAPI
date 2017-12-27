using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;


namespace PRStatusAPITool
{
    public class PrList
    {
        public string BugId { get; set; }
        public string PRLink { get; set; }
        public string PRNum { get; set; }
        public string PRStatus { get; set; }

    }
}
