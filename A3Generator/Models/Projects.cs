using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3Generator.Models
{
    public class Projects
    {
        [JsonProperty("value")]
        public List<Project> Value { get; set; }
    }

    public class Project : BaseModel
    {
        [JsonProperty("state")]
        public string State { get; set; }

        [JsonProperty("revision")]
        public long Revision { get; set; }

        [JsonProperty("visibility")]
        public string Visibility { get; set; }

        [JsonProperty("lastUpdateTime")]
        public DateTime LastUpdateTime { get; set; }
    }
}
