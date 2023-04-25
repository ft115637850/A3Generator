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
        [JsonProperty("Profiles")]
        public List<Profile> Profiles { get; set; }
    }

    public class Profile
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

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

        [JsonProperty("Query")]
        public string Query { get; set; }

        [JsonProperty("ExportFilePrefix")]
        public string ExportFilePrefix { get; set; }

        [JsonProperty("Orgnization")]
        public string Orgnization { get; set; }
    }
}
