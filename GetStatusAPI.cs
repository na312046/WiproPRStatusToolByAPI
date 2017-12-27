using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;

namespace PRStatusAPITool
{
    public class GetStatusAPI
    {
        private VssConnection connection;
        private readonly GitHttpClient gitClient;
        private readonly WorkItemTrackingHttpClient witClient;

        public GetStatusAPI(VssConnection connection)
        {
            this.connection = connection;
            this.gitClient = connection.GetClient<GitHttpClient>();
            this.witClient = connection.GetClient<WorkItemTrackingHttpClient>();
        }

        public async Task<List<String>> GetPrStatus(int prID)
        {
            var results = new List<String>();
            try
            {
                var targetPr = await gitClient.GetPullRequestByIdAsync(prID).ConfigureAwait(false);

                var Stats = targetPr.Status.ToString();
                results.Add(Stats);
            }
            catch (Exception ex)
            {
                string stat = ex.Message;
                results.Add(stat);
            }
            return results;

        }
    }
}