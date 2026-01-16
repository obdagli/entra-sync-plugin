using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace EntraSyncPlugin
{
    public partial class EntraSyncControl : PluginControlBase
    {
        private TextBox txtEmail;
        private Button btnSync;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Label lblSummary;
        private DataGridView dgvRoles;
        private DataGridView dgvTeams;

        public EntraSyncControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.btnSync = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblWarning = new System.Windows.Forms.Label();
            this.lblSummary = new System.Windows.Forms.Label();
            this.dgvRoles = new System.Windows.Forms.DataGridView();
            this.dgvTeams = new System.Windows.Forms.DataGridView();

            var cardPanel = new System.Windows.Forms.Panel();
            var lblTitle = new System.Windows.Forms.Label();
            var lblSub = new System.Windows.Forms.Label();
            var lblHowTitle = new System.Windows.Forms.Label();
            var lblHowBody = new System.Windows.Forms.Label();
            var footerPanel = new System.Windows.Forms.Panel();
            var lblEmail = new System.Windows.Forms.Label();
            var rootSplit = new System.Windows.Forms.SplitContainer();
            var rightSplit = new System.Windows.Forms.SplitContainer();
            var gridSplit = new System.Windows.Forms.SplitContainer();
            var summaryPanel = new System.Windows.Forms.Panel();
            var rolesPanel = new System.Windows.Forms.Panel();
            var teamsPanel = new System.Windows.Forms.Panel();
            var lblSummaryTitle = new System.Windows.Forms.Label();
            var lblRolesTitle = new System.Windows.Forms.Label();
            var lblTeamsTitle = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(rootSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(rightSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(gridSplit)).BeginInit();
            rootSplit.Panel1.SuspendLayout();
            rootSplit.Panel2.SuspendLayout();
            rightSplit.Panel1.SuspendLayout();
            rightSplit.Panel2.SuspendLayout();
            gridSplit.Panel1.SuspendLayout();
            gridSplit.Panel2.SuspendLayout();
            cardPanel.SuspendLayout();
            this.SuspendLayout();

            // 
            // rootSplit
            // 
            rootSplit.Dock = DockStyle.Fill;
            rootSplit.Orientation = Orientation.Vertical;
            rootSplit.SplitterWidth = 6;
            rootSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);

            // 
            // cardPanel
            // 
            cardPanel.BackColor = System.Drawing.Color.White;
            cardPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            cardPanel.Controls.Add(lblTitle);
            cardPanel.Controls.Add(lblSub);
            cardPanel.Controls.Add(lblEmail);
            cardPanel.Controls.Add(lblHowTitle);
            cardPanel.Controls.Add(lblHowBody);
            cardPanel.Controls.Add(footerPanel);
            cardPanel.Controls.Add(this.lblStatus);
            cardPanel.Controls.Add(this.btnSync);
            cardPanel.Controls.Add(this.txtEmail);
            cardPanel.Dock = DockStyle.Fill;
            cardPanel.Name = "cardPanel";
            cardPanel.TabIndex = 0;

            // 
            // lblTitle
            // 
            lblTitle.AutoSize = true;
            lblTitle.Font = new System.Drawing.Font("Trebuchet MS", 14F, System.Drawing.FontStyle.Bold);
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(31, 41, 55);
            lblTitle.Location = new System.Drawing.Point(20, 16);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new System.Drawing.Size(214, 29);
            lblTitle.Text = "Entra User Sync";

            // 
            // lblSub
            // 
            lblSub.AutoSize = true;
            lblSub.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            lblSub.ForeColor = System.Drawing.Color.FromArgb(107, 114, 128);
            lblSub.Location = new System.Drawing.Point(22, 48);
            lblSub.Name = "lblSub";
            lblSub.Size = new System.Drawing.Size(309, 18);
            lblSub.Text = "Impersonate a user and validate via WhoAmI.";

            // 
            // lblHowTitle
            // 
            lblHowTitle.AutoSize = true;
            lblHowTitle.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Bold);
            lblHowTitle.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblHowTitle.Location = new System.Drawing.Point(22, 150);
            lblHowTitle.Name = "lblHowTitle";
            lblHowTitle.Size = new System.Drawing.Size(125, 19);
            lblHowTitle.Text = "How it works";

            // 
            // lblHowBody
            // 
            lblHowBody.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            lblHowBody.ForeColor = System.Drawing.Color.FromArgb(75, 85, 99);
            lblHowBody.Location = new System.Drawing.Point(22, 174);
            lblHowBody.Name = "lblHowBody";
            lblHowBody.Size = new System.Drawing.Size(500, 72);
            lblHowBody.Text = "Flow: find user by email/UPN, impersonate, then run WhoAmI." + Environment.NewLine +
                              "Prereqs: delegate privilege; user must exist in Dataverse." + Environment.NewLine +
                              "Expected: 1-5s per sync; longer on large orgs." + Environment.NewLine +
                              "Output: user summary, roles, and teams shown on the right.";

            // 
            // lblEmail
            // 
            lblEmail.AutoSize = true;
            lblEmail.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Bold);
            lblEmail.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblEmail.Location = new System.Drawing.Point(22, 92);
            lblEmail.Name = "lblEmail";
            lblEmail.Size = new System.Drawing.Size(155, 19);
            lblEmail.Text = "User email or UPN";

            // 
            // lblWarning
            // 
            this.lblWarning.AutoSize = true;
            this.lblWarning.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            this.lblWarning.ForeColor = System.Drawing.Color.FromArgb(185, 28, 28);
            this.lblWarning.Location = new System.Drawing.Point(10, 8);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(351, 18);
            this.lblWarning.TabIndex = 3;
            this.lblWarning.Text = "Admin warning: requires Delegate privileges.";

            // 
            // txtEmail
            // 
            this.txtEmail.Font = new System.Drawing.Font("Trebuchet MS", 10F, System.Drawing.FontStyle.Regular);
            this.txtEmail.Location = new System.Drawing.Point(22, 116);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(360, 27);
            this.txtEmail.TabIndex = 0;
            this.txtEmail.Text = "user@example.com";
            this.txtEmail.ForeColor = Color.Gray;
            this.txtEmail.Enter += (s, e) => { if (txtEmail.Text == "user@example.com") { txtEmail.Text = ""; txtEmail.ForeColor = Color.Black; } };
            this.txtEmail.Leave += (s, e) => { if (string.IsNullOrWhiteSpace(txtEmail.Text)) { txtEmail.Text = "user@example.com"; txtEmail.ForeColor = Color.Gray; } };

            // 
            // btnSync
            // 
            this.btnSync.BackColor = System.Drawing.Color.FromArgb(15, 118, 110);
            this.btnSync.ForeColor = System.Drawing.Color.White;
            this.btnSync.FlatStyle = FlatStyle.Flat;
            this.btnSync.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Bold);
            this.btnSync.Location = new System.Drawing.Point(400, 114);
            this.btnSync.Name = "btnSync";
            this.btnSync.Size = new System.Drawing.Size(118, 32);
            this.btnSync.TabIndex = 1;
            this.btnSync.Text = "Force Sync";
            this.btnSync.UseVisualStyleBackColor = false;
            this.btnSync.FlatAppearance.BorderSize = 0;
            this.btnSync.Click += new System.EventHandler(this.BtnSync_Click);

            // 
            // footerPanel
            // 
            footerPanel.BackColor = System.Drawing.Color.FromArgb(249, 250, 251);
            footerPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            footerPanel.Location = new System.Drawing.Point(22, 248);
            footerPanel.Name = "footerPanel";
            footerPanel.Size = new System.Drawing.Size(496, 36);
            footerPanel.TabIndex = 10;
            footerPanel.Controls.Add(this.lblWarning);

            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Italic);
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(75, 85, 99);
            this.lblStatus.Location = new System.Drawing.Point(22, 132);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 20);
            this.lblStatus.TabIndex = 2;

            // 
            // rightSplit
            // 
            rightSplit.Dock = DockStyle.Fill;
            rightSplit.Orientation = Orientation.Horizontal;
            rightSplit.SplitterWidth = 6;
            rightSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);

            // 
            // summaryPanel
            // 
            summaryPanel.BackColor = System.Drawing.Color.White;
            summaryPanel.BorderStyle = BorderStyle.FixedSingle;
            summaryPanel.Padding = new Padding(10);
            summaryPanel.Dock = DockStyle.Fill;

            // 
            // lblSummaryTitle
            // 
            lblSummaryTitle.Text = "User Summary";
            lblSummaryTitle.Font = new Font("Trebuchet MS", 10.5F, FontStyle.Bold);
            lblSummaryTitle.ForeColor = Color.FromArgb(31, 41, 55);
            lblSummaryTitle.Dock = DockStyle.Top;
            lblSummaryTitle.Height = 22;

            // 
            // lblSummary
            // 
            this.lblSummary.Font = new Font("Trebuchet MS", 9.5F, FontStyle.Regular);
            this.lblSummary.ForeColor = Color.FromArgb(55, 65, 81);
            this.lblSummary.Dock = DockStyle.Fill;
            this.lblSummary.Text = "No sync data yet.";

            // 
            // gridSplit
            // 
            gridSplit.Dock = DockStyle.Fill;
            gridSplit.Orientation = Orientation.Vertical;
            gridSplit.SplitterWidth = 6;
            gridSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);

            // 
            // rolesPanel
            // 
            rolesPanel.BackColor = Color.White;
            rolesPanel.BorderStyle = BorderStyle.FixedSingle;
            rolesPanel.Padding = new Padding(10);
            rolesPanel.Dock = DockStyle.Fill;

            // 
            // lblRolesTitle
            // 
            lblRolesTitle.Text = "Personal Security Roles";
            lblRolesTitle.Font = new Font("Trebuchet MS", 10.5F, FontStyle.Bold);
            lblRolesTitle.ForeColor = Color.FromArgb(31, 41, 55);
            lblRolesTitle.Dock = DockStyle.Top;
            lblRolesTitle.Height = 22;

            // 
            // dgvRoles
            // 
            ConfigureGrid(this.dgvRoles, new[] { "Role", "Business Unit" });
            this.dgvRoles.Dock = DockStyle.Fill;

            // 
            // teamsPanel
            // 
            teamsPanel.BackColor = Color.White;
            teamsPanel.BorderStyle = BorderStyle.FixedSingle;
            teamsPanel.Padding = new Padding(10);
            teamsPanel.Dock = DockStyle.Fill;

            // 
            // lblTeamsTitle
            // 
            lblTeamsTitle.Text = "Assigned Teams";
            lblTeamsTitle.Font = new Font("Trebuchet MS", 10.5F, FontStyle.Bold);
            lblTeamsTitle.ForeColor = Color.FromArgb(31, 41, 55);
            lblTeamsTitle.Dock = DockStyle.Top;
            lblTeamsTitle.Height = 22;

            // 
            // dgvTeams
            // 
            ConfigureGrid(this.dgvTeams, new[] { "Team", "Type", "Business Unit" });
            this.dgvTeams.Dock = DockStyle.Fill;

            summaryPanel.Controls.Add(this.lblSummary);
            summaryPanel.Controls.Add(lblSummaryTitle);

            rolesPanel.Controls.Add(this.dgvRoles);
            rolesPanel.Controls.Add(lblRolesTitle);

            teamsPanel.Controls.Add(this.dgvTeams);
            teamsPanel.Controls.Add(lblTeamsTitle);

            gridSplit.Panel1.Controls.Add(rolesPanel);
            gridSplit.Panel2.Controls.Add(teamsPanel);

            rightSplit.Panel1.Controls.Add(summaryPanel);
            rightSplit.Panel2.Controls.Add(gridSplit);

            rootSplit.Panel1.Controls.Add(cardPanel);
            rootSplit.Panel2.Controls.Add(rightSplit);

            // 
            // EntraSyncControl
            // 
            this.BackColor = System.Drawing.Color.FromArgb(248, 245, 240);
            this.Controls.Add(rootSplit);
            this.Name = "EntraSyncControl";
            this.Size = new System.Drawing.Size(980, 520);

            cardPanel.ResumeLayout(false);
            cardPanel.PerformLayout();
            rootSplit.Panel1.ResumeLayout(false);
            rootSplit.Panel2.ResumeLayout(false);
            rightSplit.Panel1.ResumeLayout(false);
            rightSplit.Panel2.ResumeLayout(false);
            gridSplit.Panel1.ResumeLayout(false);
            gridSplit.Panel2.ResumeLayout(false);
            rootSplit.EndInit();
            rightSplit.EndInit();
            gridSplit.EndInit();
            rootSplit.Panel1.ResumeLayout(false);
            rootSplit.Panel2.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        private void BtnSync_Click(object sender, EventArgs e)
        {
            string email = txtEmail.Text.Trim();
            if (string.IsNullOrEmpty(email) || email == "user@example.com")
            {
                MessageBox.Show("Please enter a valid email address.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Attempting to sync user: {email}...",
                ProgressChanged = (progress) =>
                {
                    if (progress?.UserState is string status)
                    {
                        lblStatus.Text = status;
                        lblStatus.ForeColor = Color.DarkBlue;
                    }
                },
                Work = (worker, args) =>
                {
                    // 1. Find User by Email (Internalemailaddress or WindowsLiveID)
                    var query = new QueryExpression("systemuser");
                    query.ColumnSet = new ColumnSet("systemuserid", "fullname");
                    query.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, email);

                    // Fallback to domainname if email empty (often UPN)
                    var query2 = new QueryExpression("systemuser");
                    query2.ColumnSet = new ColumnSet("systemuserid", "fullname");
                    query2.Criteria.AddCondition("domainname", ConditionOperator.Equal, email);

                    var result = Service.RetrieveMultiple(query);
                    if (result.Entities.Count == 0)
                    {
                        result = Service.RetrieveMultiple(query2);
                    }

                    if (result.Entities.Count == 0)
                    {
                        throw new Exception($"User with email/UPN '{email}' not found in Dataverse.\nCannot impersonate a non-existent user record.");
                    }

                    var targetUser = result.Entities[0];
                    Guid targetUserId = targetUser.Id;
                    string fullName = targetUser.GetAttributeValue<string>("fullname");

                    worker.ReportProgress(20, $"User found: {fullName} ({targetUserId})");
                    worker.ReportProgress(40, "Impersonating...");

                    var details = new SyncResultDetails
                    {
                        FullName = fullName,
                        Email = email,
                        UserId = targetUserId,
                        Roles = LoadUserRoles(Service, targetUserId),
                        Teams = LoadUserTeams(Service, targetUserId)
                    };

                    try
                    {
                        var svc = this.ConnectionDetail.GetCrmServiceClient();
                        if (svc != null)
                        {
                            svc.CallerId = targetUserId;
                            var response = svc.Execute(new WhoAmIRequest());
                            details.WhoAmIUserId = ((WhoAmIResponse)response).UserId;
                            args.Result = details;
                        }
                        else
                        {
                            throw new Exception("Could not create an impersonated service client.");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Impersonation failed: {ex.Message}");
                    }
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        lblStatus.Text = "Error";
                        lblStatus.ForeColor = Color.Red;
                        MessageBox.Show(args.Error.Message, "Sync Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        var details = args.Result as SyncResultDetails;
                        lblStatus.Text = "Sync Complete";
                        lblStatus.ForeColor = Color.Green;
                        if (details != null)
                        {
                            UpdateDetails(details);
                        }
                    }
                }
            });
        }

        private void UpdateDetails(SyncResultDetails details)
        {
            lblSummary.Text =
                "Name: " + details.FullName + Environment.NewLine +
                "Email: " + details.Email + Environment.NewLine +
                "UserId: " + details.UserId + Environment.NewLine +
                "WhoAmI: " + (details.WhoAmIUserId.HasValue ? details.WhoAmIUserId.ToString() : "n/a");

            dgvRoles.Rows.Clear();
            if (details.Roles.Count == 0)
            {
                dgvRoles.Rows.Add("(none)", string.Empty);
            }
            else
            {
                foreach (var role in details.Roles)
                {
                    dgvRoles.Rows.Add(role.Name, role.BusinessUnit);
                }
            }

            dgvTeams.Rows.Clear();
            if (details.Teams.Count == 0)
            {
                dgvTeams.Rows.Add("(none)", string.Empty, string.Empty);
            }
            else
            {
                foreach (var team in details.Teams)
                {
                    dgvTeams.Rows.Add(team.Name, team.TeamType, team.BusinessUnit);
                }
            }
        }

        private static void ConfigureGrid(DataGridView grid, string[] headers)
        {
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.ReadOnly = true;
            grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.Font = new Font("Trebuchet MS", 9.5F, FontStyle.Regular);

            foreach (var header in headers)
            {
                grid.Columns.Add(header, header);
            }
        }

        private static List<RoleInfo> LoadUserRoles(IOrganizationService service, Guid userId)
        {
            var roles = new List<RoleInfo>();
            var query = new QueryExpression("systemuserroles")
            {
                ColumnSet = new ColumnSet("systemuserid", "roleid")
            };
            query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
            var link = query.AddLink("role", "roleid", "roleid");
            link.Columns = new ColumnSet("name", "businessunitid");
            link.EntityAlias = "r";

            var result = service.RetrieveMultiple(query);
            foreach (var entity in result.Entities)
            {
                var name = (entity.GetAttributeValue<AliasedValue>("r.name")?.Value as string) ?? "(unknown)";
                var buRef = entity.GetAttributeValue<AliasedValue>("r.businessunitid")?.Value as EntityReference;
                roles.Add(new RoleInfo
                {
                    Name = name,
                    BusinessUnit = buRef?.Name ?? buRef?.Id.ToString() ?? string.Empty
                });
            }

            return roles;
        }

        private static List<TeamInfo> LoadUserTeams(IOrganizationService service, Guid userId)
        {
            var teams = new List<TeamInfo>();
            var query = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("systemuserid", "teamid")
            };
            query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);
            var link = query.AddLink("team", "teamid", "teamid");
            link.Columns = new ColumnSet("name", "teamtype", "businessunitid");
            link.EntityAlias = "t";

            var result = service.RetrieveMultiple(query);
            foreach (var entity in result.Entities)
            {
                var name = (entity.GetAttributeValue<AliasedValue>("t.name")?.Value as string) ?? "(unknown)";
                var teamTypeValue = (entity.GetAttributeValue<AliasedValue>("t.teamtype")?.Value as OptionSetValue)?.Value;
                var buRef = entity.GetAttributeValue<AliasedValue>("t.businessunitid")?.Value as EntityReference;
                teams.Add(new TeamInfo
                {
                    Name = name,
                    TeamType = FormatTeamType(teamTypeValue),
                    BusinessUnit = buRef?.Name ?? buRef?.Id.ToString() ?? string.Empty
                });
            }

            return teams;
        }

        private static string FormatTeamType(int? teamType)
        {
            if (teamType == null)
            {
                return "Unknown";
            }

            switch (teamType.Value)
            {
                case 0:
                    return "Owner";
                case 1:
                    return "Access";
                case 2:
                    return "Security Group";
                case 3:
                    return "Office Group";
                default:
                    return teamType.Value.ToString();
            }
        }
    }
}
