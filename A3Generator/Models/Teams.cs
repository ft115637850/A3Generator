using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3Generator.Models
{
    public class Teams
    {
        [JsonProperty("value")]
        public List<Team> Value { get; set; }
    }

    public class Team : BaseModel
    {
        [JsonProperty("identityUrl")]
        public string IdentityUrl { get; set; }

        [JsonProperty("projectName")]
        public string ProjectName { get; set; }

        [JsonProperty("projectId")]
        public string ProjectId { get; set; }
    }
}
