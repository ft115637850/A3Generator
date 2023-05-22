using A3Generator.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Text;

namespace A3Generator
{
    public partial class A3Generator : Form
    {
        private AdoService service;
        private List<Project> projects;
        private InputCache inputCache;
        private readonly List<UserStory> currentWorkItems;
        private string pat;
        private bool loadComplete = false;
        private bool savingProfile = false;
        public A3Generator()
        {
            InitializeComponent();
            currentWorkItems = new List<UserStory>();
            inputCache = new InputCache
            {
                Profiles = new List<Profile>
                {
                    new Profile
                    {
                        Name = "Default"
                    }
                }
            };
            this.comboBox2.ValueMember = "Name";
            this.comboBox2.DisplayMember = "Name";
            this.comboBox2.DataSource = this.inputCache.Profiles;
        }

        private void A3Generator_Load(object sender, EventArgs e)
        {
            LoadFromCache();
            loadComplete = true;
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.SelectedValue == null)
            {
                MessageBox.Show("Project Team is not selected");
                return;
            }

            if (string.IsNullOrEmpty(this.textBox1.Text))
            {
                MessageBox.Show("Filter is empty");
                return;
            }

            if (string.IsNullOrEmpty(this.textBox4.Text))
            {
                MessageBox.Show("Iteration is empty");
                return;
            }

            var projectId = this.comboBox1.SelectedValue.ToString();
            var filter = $"Iteration/IterationPath eq '{this.textBox4.Text.Trim()}' and {this.textBox1.Text.Trim()}";

            TaskScheduler syncSch = TaskScheduler.FromCurrentSynchronizationContext();
            Task.Run(async Task<WorkItems>? () => await service.GetBoardListAsync(projectId, filter).ConfigureAwait(false))
                .ContinueWith(
                task =>
                {
                    if (task.Result?.Value == null)
                        return;

                    var members = this.membersTextBox.Text.Trim().ToLowerInvariant();

                    this.listView1.BeginUpdate();
                    this.listView1.Items.Clear();
                    this.listView2.Items.Clear();
                    this.currentWorkItems.Clear();
                    foreach (var userStory in task.Result.Value)
                    {
                        var user = userStory.AssignedTo?.UserName.ToLowerInvariant();
                        if (string.IsNullOrEmpty(user) || !members.Contains(user))
                        {
                            var isMyTeamWorkItem = userStory.Children.Any(t => !string.IsNullOrEmpty(t.AssignedTo?.UserName)
                                && members.Contains(t.AssignedTo.UserName.ToLowerInvariant()));
                            if (!isMyTeamWorkItem) continue;
                        }

                        this.currentWorkItems.Add(userStory);
                        ListViewItem item = new ListViewItem();
                        item.Text = userStory.WorkItemId;
                        item.SubItems.Add(userStory.Title);
                        item.SubItems.Add(userStory.WorkItemType);
                        item.SubItems.Add(userStory.State);
                        item.SubItems.Add(userStory.StoryPoints.ToString());
                        item.SubItems.Add(user == null ? String.Empty : user);
                        item.Tag = userStory.Children;
                        this.listView1.Items.Add(item);
                    }
                    this.listView1.EndUpdate();
                    this.button4.Enabled = true;
                    CalculateProductivity();
                }, syncSch);
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (this.listView1.SelectedItems.Count == 0) return;
            this.button3.Enabled = true;
            var selectedItem = this.listView1.SelectedItems[0] as ListViewItem;
            if (selectedItem == null || selectedItem.Tag == null) return;

            this.listView2.BeginUpdate();
            this.listView2.Items.Clear();
            var workItemTasks = selectedItem.Tag as List<WorkItemTask>;
            foreach (var workItemTask in workItemTasks)
            {
                ListViewItem item = new ListViewItem();
                item.Text = workItemTask.WorkItemId;
                item.SubItems.Add(workItemTask.Title);
                item.SubItems.Add(workItemTask.WorkItemType);
                item.SubItems.Add(workItemTask.State);
                item.SubItems.Add(workItemTask.OriginalEstimate == null ? String.Empty: workItemTask.OriginalEstimate.ToString());
                item.SubItems.Add(workItemTask.RemainingWork == null ? String.Empty : workItemTask.RemainingWork.ToString());
                item.SubItems.Add(workItemTask.CompletedWork == null ? String.Empty : workItemTask.CompletedWork.ToString());
                item.SubItems.Add(workItemTask.AssignedTo?.UserName == null? String.Empty : workItemTask.AssignedTo.UserName);
                this.listView2.Items.Add(item);
            }
            this.listView2.EndUpdate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.listView1.SelectedItems.Count == 0) return;
            var selectedItem = this.listView1.SelectedItems[0] as ListViewItem;
            this.listView1.Items.Remove(selectedItem);
            this.listView2.Items.Clear();
            this.button3.Enabled = false;
            var deletedItem = this.currentWorkItems.Find(x => x.WorkItemId == selectedItem.Text);
            this.currentWorkItems.Remove(deletedItem);
            CalculateProductivity();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var excel = textBox2.Text.Trim();
            var iteration = textBox4.Text.Trim().Split('\\');
            var sprint = iteration[2];
            var projectId = this.comboBox1.SelectedValue.ToString();
            if (string.IsNullOrEmpty(excel) || string.IsNullOrEmpty(sprint)) return;

            var columns = new string[] { "ID", "Title", "Work Item Type", "Assigned To", "Story Points",
                "State", "Original Estimate", "Completed Work", "Remaining Work", "Iteration Path"};
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var fileStream = new FileStream($"{iteration[1]}{excel} {sprint}.xlsx", FileMode.Create))
            using (var package = new ExcelPackage(fileStream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sprint);
                worksheet.Cells.Style.Font.Name = "Segoe UI";
                worksheet.Cells.Style.Font.Size = 9;
                int rowIndex = 1;
                WriteHeaders(rowIndex, worksheet, columns);

                rowIndex++;


                for (int i = 0; i < this.currentWorkItems.Count; i++)
                {
                    var userStory = currentWorkItems[i];
                    WriteUserStories(rowIndex, i, projectId, worksheet, userStory);
                    if (userStory.Children == null) continue;
                    foreach(var task in userStory.Children)
                    {
                        rowIndex++;
                        WriteTasks(rowIndex, i, projectId, worksheet, task);
                    }
                }

                package.Save();
            }

            MessageBox.Show("Completed");
            
        }

        private void WriteHeaders(int rowIndex, ExcelWorksheet worksheet, string[] columns)
        {
            int colIndex = 1;

            for (int i = 0; i < columns.Length; i++)
            {
                worksheet.Cells[rowIndex, colIndex + i].Value = columns[i];
                worksheet.Cells[rowIndex, colIndex + i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[rowIndex, colIndex + i].Style.Fill.BackgroundColor.SetColor(25, 16, 110, 190);
                worksheet.Cells[rowIndex, colIndex + i].Style.Font.Color.SetColor(Color.White);
                worksheet.Cells[rowIndex, colIndex + i].Style.Font.Bold = true;
                worksheet.Column(colIndex + i).Width = 26;
            }
        }

        private void WriteUserStories(int rowIndex, int i, string projectId, ExcelWorksheet worksheet, UserStory userStory)
        {
            worksheet.Cells[rowIndex + i, 1].Value = Convert.ToInt32(userStory.WorkItemId);
            worksheet.Cells[rowIndex + i, 1].Style.Font.UnderLine = true;
            worksheet.Cells[rowIndex + i, 1].Style.Font.Name = "Calibri";
            worksheet.Cells[rowIndex + i, 1].Style.Font.Size = 11;
            worksheet.Cells[rowIndex + i, 1].Style.Font.Color.SetColor(25, 5, 99, 193);
            worksheet.Cells[rowIndex + i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 1].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 1].Hyperlink = new Uri($"https://dev.azure.com/AVEVA-VSTS/{projectId}/_workitems/edit/{userStory.WorkItemId}");

            worksheet.Cells[rowIndex + i, 2].Value = userStory.Title;
            worksheet.Cells[rowIndex + i, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 2].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 3].Value = userStory.WorkItemType;
            worksheet.Cells[rowIndex + i, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 3].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 4].Value = userStory.AssignedTo?.UserName;
            worksheet.Cells[rowIndex + i, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 4].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 5].Value = userStory.StoryPoints;
            worksheet.Cells[rowIndex + i, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 5].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 6].Value = "Resolved";
            worksheet.Cells[rowIndex + i, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 6].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);

            worksheet.Cells[rowIndex + i, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 7].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 8].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Cells[rowIndex + i, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 9].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);

            worksheet.Cells[rowIndex + i, 10].Value = textBox4.Text.Trim();
            worksheet.Cells[rowIndex + i, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rowIndex + i, 10].Style.Fill.BackgroundColor.SetColor(80, 221, 235, 247);
            worksheet.Row(rowIndex + i).CustomHeight = true;
        }

        private void WriteTasks(int rowIndex, int i, string projectId, ExcelWorksheet worksheet, WorkItemTask task)
        {
            worksheet.Cells[rowIndex + i, 1].Value = Convert.ToInt32(task.WorkItemId);
            worksheet.Cells[rowIndex + i, 1].Style.Font.UnderLine = true;
            worksheet.Cells[rowIndex + i, 1].Style.Font.Name = "Calibri";
            worksheet.Cells[rowIndex + i, 1].Style.Font.Size = 11;
            worksheet.Cells[rowIndex + i, 1].Style.Font.Color.SetColor(25, 5, 99, 193);
            worksheet.Cells[rowIndex + i, 1].Hyperlink = new Uri($"https://dev.azure.com/AVEVA-VSTS/{projectId}/_workitems/edit/{task.WorkItemId}");

            worksheet.Cells[rowIndex + i, 2].Value = task.Title;
            worksheet.Cells[rowIndex + i, 3].Value = task.WorkItemType;
            worksheet.Cells[rowIndex + i, 4].Value = task.AssignedTo?.UserName;

            worksheet.Cells[rowIndex + i, 6].Value = "Closed";

            worksheet.Cells[rowIndex + i, 7].Value = task.OriginalEstimate;
            worksheet.Cells[rowIndex + i, 8].Value = task.CompletedWork;
            worksheet.Cells[rowIndex + i, 9].Value = task.RemainingWork;

            worksheet.Cells[rowIndex + i, 10].Value = textBox4.Text.Trim();
            worksheet.Row(rowIndex + i).CustomHeight = true;
        }

        private void CalculateProductivity()
        {
            var membersCalc = new Dictionary<string, decimal>();
            var members = this.membersTextBox.Text.Trim().ToLowerInvariant().Split(',').Distinct();
            foreach (var member in members)
            {
                membersCalc[member] = 0;
            }

            foreach(var userStory in this.currentWorkItems)
            {
                var calcHours = new Dictionary<string, decimal>();
                foreach (var member in members)
                {
                    calcHours[member] = 0;
                }

                decimal lastStoryTotalHours = 0;
                foreach (var task in userStory.Children)
                {
                    var user = task.AssignedTo?.UserName?.ToLowerInvariant();
                    lastStoryTotalHours += task.CompletedWork ?? 0;
                    if (!string.IsNullOrEmpty(user) &&
                        calcHours.ContainsKey(user))
                    {
                        calcHours[user] += task.CompletedWork ?? 0;
                    }
                }

                if (lastStoryTotalHours > 0)
                {
                    foreach (var member in members)
                    {
                        var points = userStory.StoryPoints * calcHours[member] / lastStoryTotalHours;
                        membersCalc[member] += points;
                    }
                }
                else
                {
                    MessageBox.Show($"There is no completed hour in {userStory.WorkItemId} {userStory.Title}");
                }
            }

            var productiviy = new StringBuilder();
            foreach (var member in members)
            {
                var result = Math.Round(membersCalc[member], 2, MidpointRounding.AwayFromZero);
                productiviy.Append($"{member}: {result}   ");
            }
            this.label3.Text = productiviy.ToString();
        }

        private void GetProjects(string pat, string orgnization)
        {
            service = new AdoService(pat, orgnization);
            TaskScheduler syncSch = TaskScheduler.FromCurrentSynchronizationContext();
            Task.Run(async Task<Projects>? () => await service.GetAllProjectsAsync().ConfigureAwait(false))
                .ContinueWith(
                task =>
                {
                    this.comboBox1.ValueMember = "Id";
                    this.comboBox1.DisplayMember = "Name";
                    this.comboBox1.DataSource = task.Result.Value;
                    projects = task.Result.Value;
                    button2.Enabled = true;
                }, syncSch);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("Orgnization is required");
                return;
            }

            var logonForm = new LogonForm();
            if (logonForm.ShowDialog(this) == DialogResult.OK)
            {
                pat = logonForm.PAT;
                GetProjects(pat, textBox3.Text.Trim());
            }
        }

        private void A3Generator_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!string.IsNullOrEmpty(pat)) SaveProfile();
        }

        private void SaveProfile()
        {
            var projectId = this.comboBox1.SelectedValue.ToString();
            var profileName = this.comboBox2.Text.ToString();
            if (string.IsNullOrEmpty(profileName) ||string.IsNullOrEmpty(projectId))
            {
                MessageBox.Show("Profile Name and projectId can't be empty");
                return;
            }

            var profile = new Profile
            {
                Name = profileName,
                Projects = projects,
                SelectedProject = projects.Find(x => x.Id == projectId),
                Interation = textBox4.Text.Trim(),
                Members = this.membersTextBox.Text.Trim().ToLowerInvariant(),
                PAT = pat,
                Query = textBox1.Text.Trim(),
                ExportFilePrefix = textBox2.Text.Trim(),
                Orgnization = textBox3.Text.Trim()
            };

            var existing = inputCache.Profiles.Find(p => p.Name == profileName);
            if (existing != null)
            {
                inputCache.Profiles.Remove(existing);
            }

            inputCache.Profiles.Add(profile);
            var inputs = JsonConvert.SerializeObject(inputCache);
            File.WriteAllText("InputCache.json", inputs);
        }

        private void LoadFromCache()
        {
            if (!File.Exists("InputCache.json")) return;
            
                
            var inputs = File.ReadAllText("InputCache.json");
            inputCache = JsonConvert.DeserializeObject<InputCache>(inputs);
            if (inputCache?.Profiles?.Any() != true) return;

            this.comboBox2.DataSource = this.inputCache.Profiles;
            var profile = inputCache.Profiles.First();
            LoadProfile(profile);
        }

        private void LoadProfile(Profile profile)
        {
            this.comboBox1.ValueMember = "Id";
            this.comboBox1.DisplayMember = "Name";
            this.comboBox1.DataSource = profile.Projects;
            projects = profile.Projects;
            this.comboBox1.SelectedValue = profile.SelectedProject.Id;
            textBox3.Text = profile.Orgnization;
            textBox4.Text = profile.Interation;
            this.membersTextBox.Text = profile.Members;
            pat = profile.PAT;
            textBox1.Text = profile.Query;
            textBox2.Text = profile.ExportFilePrefix;
            service = new AdoService(pat, profile.Orgnization);
            button2.Enabled = true;
        }

        private void WorkItem_DoubleClick(object sender, EventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null || listView.SelectedItems.Count == 0) return;
            var projectId = this.comboBox1.SelectedValue.ToString();
            var selectedItem = listView.SelectedItems[0] as ListViewItem;
            var target = $"https://dev.azure.com/AVEVA-VSTS/{projectId}/_workitems/edit/{selectedItem.Text}";
            var chrome = @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
            Process.Start(chrome, target);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            savingProfile = true;
            var currentProfile = this.comboBox2.Text;
            SaveProfile();
            this.comboBox2.DataSource = null;
            this.comboBox2.ValueMember = "Name";
            this.comboBox2.DisplayMember = "Name";
            this.comboBox2.DataSource = this.inputCache.Profiles;
            this.comboBox2.SelectedValue = currentProfile;
            savingProfile = false;
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (savingProfile == false && loadComplete && !string.IsNullOrEmpty(this.comboBox2.SelectedValue?.ToString()))
            {
                var profileName = this.comboBox2.SelectedValue.ToString().Trim();
                var existing = inputCache.Profiles.Find(p => p.Name == profileName);
                if (existing == null) return;
                LoadProfile(existing);
                this.listView1.Items.Clear();
                this.listView2.Items.Clear();
            }
        }
    }
}