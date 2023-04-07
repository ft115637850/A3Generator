using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3Generator.Models
{
    public class InputCache
    {
        [JsonProperty("PAT")]
        public string PAT { get; set; }

        [JsonProperty("Projects")]
        public List<Project> Projects { get; set; }

        [JsonProperty("SelectedProject")]
        public Project SelectedProject { get; set; }

        [JsonProperty("Members")]
        public string Members { get; set; }

        [JsonProperty("Interation")]
        public string Interation { get; set; }

        [JsonProperty("Orgnization")]
        public string Orgnization { get; set; }
    }
}
