using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A3Generator.Models
{
    public class WorkItems
    {
        [JsonProperty("value")]
        public List<UserStory> Value { get; set; }
    }

    public class WorkItem
    {
        [JsonProperty("WorkItemId")]
        public string WorkItemId { get; set; }

        [JsonProperty("Title")]
        public string Title { get; set; }

        [JsonProperty("WorkItemType")]
        public string WorkItemType { get; set; }

        [JsonProperty("State")]
        public string State { get; set; }

        [JsonProperty("AssignedTo")]
        public User AssignedTo { get; set; }
    }

    public class UserStory : WorkItem
    {
        [JsonProperty("StoryPoints")]
        public decimal StoryPoints { get; set; }

        [JsonProperty("Children")]
        public List<WorkItemTask> Children { get; set; }
    }

    public class WorkItemTask : WorkItem
    {
        [JsonProperty("CompletedWork")]
        public decimal? CompletedWork { get; set; }

        [JsonProperty("RemainingWork")]
        public decimal? RemainingWork { get; set; }

        [JsonProperty("OriginalEstimate")]
        public decimal? OriginalEstimate { get; set; }
    }

    public class User
    {
        [JsonProperty("UserEmail", Required = Required.Default)]
        public string UserEmail { get; set; }

        [JsonProperty("UserName")]
        public string UserName { get; set; }
    }
}
