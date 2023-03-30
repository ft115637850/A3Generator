using A3Generator.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace A3Generator
{
    public class AdoService
    {
        
        private const string DEFAULT_EXPEND = "Children($expand=AssignedTo($select=UserName);$select=WorkItemId, Title, WorkItemType, State, CompletedWork,RemainingWork,OriginalEstimate,AssignedTo),AssignedTo($select=UserName)";
        private const string DEFAULT_SELECT = "WorkItemId, Title, WorkItemType, AssignedTo, StoryPoints, State";
        private readonly string _PAT;
        private readonly string _baseAddress;
        private readonly string _analyticsBaseAddress;
        private HttpClient _client;
        private HttpClient _analyticsClient;

        public AdoService(string pat, string orgnization)
        {
            _PAT = pat;
            _baseAddress = $"https://dev.azure.com/{orgnization}";
            _client = new HttpClient() { 
                BaseAddress = new Uri(_baseAddress)
            };
            _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", _PAT))));

            _analyticsBaseAddress = $"https://analytics.dev.azure.com/{orgnization}";
            _analyticsClient = new HttpClient()
            {
                BaseAddress = new Uri(_analyticsBaseAddress)
            };
            _analyticsClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", _PAT))));
        }

        public async Task<Teams> GetProjectTeamsAsync(string projectId)
        {
            var requestUri = $"{_baseAddress}/_apis/projects/{projectId}/teams?api-version=7.0";
            var request = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUri));
            var response = await _client.SendAsync(request).ConfigureAwait(false);
            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            var result = JsonConvert.DeserializeObject<Teams>(responseContent);
            return result;
        }

        public async Task<Projects> GetAllProjectsAsync()
        {
            var requestUri = $"{_baseAddress}/_apis/projects?api-version=7.0";
            var request = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUri));
            var response = await _client.SendAsync(request).ConfigureAwait(false);
            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            var result = JsonConvert.DeserializeObject<Projects>(responseContent);
            return result;
        }

        public async Task<WorkItems> GetBoardListAsync(string projectId, string filter, string expand = DEFAULT_EXPEND, string select = DEFAULT_SELECT)
        {
            try
            {
                var requestUri = $"{_analyticsBaseAddress}/{projectId}/_odata/v3.0-preview/WorkItems?$filter={filter}&$expand={expand}&$select={select}";
                var request = new HttpRequestMessage(HttpMethod.Get, new Uri(requestUri));
                var response = await _analyticsClient.SendAsync(request).ConfigureAwait(false);
                var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                var result = JsonConvert.DeserializeObject<WorkItems>(responseContent);
                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            
        }
    }
}
