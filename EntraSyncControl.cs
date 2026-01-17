using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
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
        private System.Windows.Forms.Label lblUserStatus;
        private System.Windows.Forms.Label lblGroupStatus;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Label lblSummary;
        private DataGridView dgvRoles;
        private DataGridView dgvTeams;
        private DataGridView dgvGroupMembers;
        private System.Windows.Forms.Label lblGroupResultsHint;

        private TextBox txtTenantId;
        private TextBox txtClientId;
        private TextBox txtClientSecret;
        private TextBox txtGroupId;
        private Button btnFetchGroup;
        private Button btnBeginGroupSync;
        private RadioButton rbClientSecret;
        private RadioButton rbDeviceCode;
        private CheckBox chkShowAppFields;
        private System.Windows.Forms.Label lblTenantId;
        private System.Windows.Forms.Label lblClientId;
        private System.Windows.Forms.Label lblClientSecret;
        private System.Windows.Forms.Label lblGroupId;

        private List<GraphGroupMember> groupMembers = new List<GraphGroupMember>();

        public EntraSyncControl()
        {
            InitializeComponent();
            AuthModeChanged(this, EventArgs.Empty);
        }

        private void InitializeComponent()
        {
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.btnSync = new System.Windows.Forms.Button();
            this.lblUserStatus = new System.Windows.Forms.Label();
            this.lblGroupStatus = new System.Windows.Forms.Label();
            this.lblWarning = new System.Windows.Forms.Label();
            this.lblSummary = new System.Windows.Forms.Label();
            this.dgvRoles = new System.Windows.Forms.DataGridView();
            this.dgvTeams = new System.Windows.Forms.DataGridView();
            this.dgvGroupMembers = new System.Windows.Forms.DataGridView();
            this.lblGroupResultsHint = new System.Windows.Forms.Label();

            this.txtTenantId = new System.Windows.Forms.TextBox();
            this.txtClientId = new System.Windows.Forms.TextBox();
            this.txtClientSecret = new System.Windows.Forms.TextBox();
            this.txtGroupId = new System.Windows.Forms.TextBox();
            this.btnFetchGroup = new System.Windows.Forms.Button();
            this.btnBeginGroupSync = new System.Windows.Forms.Button();
            this.rbClientSecret = new System.Windows.Forms.RadioButton();
            this.rbDeviceCode = new System.Windows.Forms.RadioButton();
            this.chkShowAppFields = new System.Windows.Forms.CheckBox();

            var leftSplit = new System.Windows.Forms.SplitContainer();
            var userCard = new System.Windows.Forms.Panel();
            var groupCard = new System.Windows.Forms.Panel();
            var lblTitle = new System.Windows.Forms.Label();
            var lblSub = new System.Windows.Forms.Label();
            var lblHowTitle = new System.Windows.Forms.Label();
            var lblHowBody = new System.Windows.Forms.Label();
            var footerPanel = new System.Windows.Forms.Panel();
            var lblEmail = new System.Windows.Forms.Label();
            var rootSplit = new System.Windows.Forms.SplitContainer();
            var rightSplit = new System.Windows.Forms.SplitContainer();
            var gridSplit = new System.Windows.Forms.SplitContainer();
            var lowerSplit = new System.Windows.Forms.SplitContainer();
            var summaryPanel = new System.Windows.Forms.Panel();
            var rolesPanel = new System.Windows.Forms.Panel();
            var teamsPanel = new System.Windows.Forms.Panel();
            var groupPanel = new System.Windows.Forms.Panel();
            var lblSummaryTitle = new System.Windows.Forms.Label();
            var lblRolesTitle = new System.Windows.Forms.Label();
            var lblTeamsTitle = new System.Windows.Forms.Label();
            var lblGroupTitle = new System.Windows.Forms.Label();

            var lblGroupSection = new System.Windows.Forms.Label();
            var lblGroupSub = new System.Windows.Forms.Label();
            this.lblTenantId = new System.Windows.Forms.Label();
            this.lblClientId = new System.Windows.Forms.Label();
            this.lblClientSecret = new System.Windows.Forms.Label();
            this.lblGroupId = new System.Windows.Forms.Label();
            var lblGroupHelp = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(rootSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(leftSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(rightSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(gridSplit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(lowerSplit)).BeginInit();
            rootSplit.Panel1.SuspendLayout();
            rootSplit.Panel2.SuspendLayout();
            leftSplit.Panel1.SuspendLayout();
            leftSplit.Panel2.SuspendLayout();
            rightSplit.Panel1.SuspendLayout();
            rightSplit.Panel2.SuspendLayout();
            gridSplit.Panel1.SuspendLayout();
            gridSplit.Panel2.SuspendLayout();
            lowerSplit.Panel1.SuspendLayout();
            lowerSplit.Panel2.SuspendLayout();
            userCard.SuspendLayout();
            groupCard.SuspendLayout();
            this.SuspendLayout();

            // 
            // rootSplit
            // 
            rootSplit.Dock = DockStyle.Fill;
            rootSplit.Orientation = Orientation.Vertical;
            rootSplit.SplitterWidth = 6;
            rootSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);
            rootSplit.SplitterDistance = 400;

            // 
            // leftSplit
            // 
            leftSplit.Dock = DockStyle.Fill;
            leftSplit.Orientation = Orientation.Horizontal;
            leftSplit.SplitterWidth = 6;
            leftSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);
            leftSplit.SplitterDistance = 360;

            // 
            // userCard
            // 
            userCard.BackColor = System.Drawing.Color.White;
            userCard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            userCard.Controls.Add(lblTitle);
            userCard.Controls.Add(lblSub);
            userCard.Controls.Add(lblEmail);
            userCard.Controls.Add(lblHowTitle);
            userCard.Controls.Add(lblHowBody);
            userCard.Controls.Add(footerPanel);
            userCard.Controls.Add(this.lblUserStatus);
            userCard.Controls.Add(this.btnSync);
            userCard.Controls.Add(this.txtEmail);
            userCard.Dock = DockStyle.Fill;
            userCard.Name = "userCard";
            userCard.TabIndex = 0;

            // 
            // groupCard
            // 
            groupCard.BackColor = System.Drawing.Color.White;
            groupCard.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            groupCard.Controls.Add(lblGroupSection);
            groupCard.Controls.Add(lblGroupSub);
            groupCard.Controls.Add(this.rbClientSecret);
            groupCard.Controls.Add(this.rbDeviceCode);
            groupCard.Controls.Add(this.chkShowAppFields);
            groupCard.Controls.Add(this.lblTenantId);
            groupCard.Controls.Add(this.txtTenantId);
            groupCard.Controls.Add(this.lblClientId);
            groupCard.Controls.Add(this.txtClientId);
            groupCard.Controls.Add(this.lblClientSecret);
            groupCard.Controls.Add(this.txtClientSecret);
            groupCard.Controls.Add(this.lblGroupId);
            groupCard.Controls.Add(this.txtGroupId);
            groupCard.Controls.Add(this.btnFetchGroup);
            groupCard.Controls.Add(this.btnBeginGroupSync);
            groupCard.Controls.Add(lblGroupHelp);
            groupCard.Controls.Add(this.lblGroupStatus);
            groupCard.Dock = DockStyle.Fill;
            groupCard.Name = "groupCard";
            groupCard.TabIndex = 1;

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
            lblSub.Location = new System.Drawing.Point(22, 46);
            lblSub.Name = "lblSub";
            lblSub.Size = new System.Drawing.Size(309, 18);
            lblSub.Text = "Impersonate a user and validate via WhoAmI.";

            // 
            // lblHowTitle
            // 
            lblHowTitle.AutoSize = true;
            lblHowTitle.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Bold);
            lblHowTitle.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblHowTitle.Location = new System.Drawing.Point(22, 178);
            lblHowTitle.Name = "lblHowTitle";
            lblHowTitle.Size = new System.Drawing.Size(125, 19);
            lblHowTitle.Text = "How it works";

            // 
            // lblHowBody
            // 
            lblHowBody.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            lblHowBody.ForeColor = System.Drawing.Color.FromArgb(75, 85, 99);
            lblHowBody.Location = new System.Drawing.Point(22, 202);
            lblHowBody.Name = "lblHowBody";
            lblHowBody.Size = new System.Drawing.Size(500, 86);
            lblHowBody.Text = "Flow: resolve user by email/UPN, impersonate, then run WhoAmI." + Environment.NewLine +
                              "Requirements: delegate privilege; user must exist in Dataverse." + Environment.NewLine +
                              "Latency: ~1-5s per sync (longer for large orgs)." + Environment.NewLine +
                              "Output: summary, roles, and team assignments shown on the right.";

            // 
            // lblEmail
            // 
            lblEmail.AutoSize = true;
            lblEmail.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Bold);
            lblEmail.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblEmail.Location = new System.Drawing.Point(22, 90);
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
            this.txtEmail.Location = new System.Drawing.Point(22, 112);
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
            this.btnSync.Location = new System.Drawing.Point(400, 110);
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
            footerPanel.Location = new System.Drawing.Point(22, 294);
            footerPanel.Name = "footerPanel";
            footerPanel.Size = new System.Drawing.Size(496, 36);
            footerPanel.TabIndex = 10;
            footerPanel.Controls.Add(this.lblWarning);

            // 
            // lblUserStatus
            // 
            this.lblUserStatus.AutoSize = true;
            this.lblUserStatus.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Italic);
            this.lblUserStatus.ForeColor = System.Drawing.Color.FromArgb(75, 85, 99);
            this.lblUserStatus.Location = new System.Drawing.Point(22, 146);
            this.lblUserStatus.Name = "lblUserStatus";
            this.lblUserStatus.Size = new System.Drawing.Size(0, 20);
            this.lblUserStatus.TabIndex = 2;

            // 
            // Group section labels
            // 
            lblGroupSection.AutoSize = true;
            lblGroupSection.Font = new System.Drawing.Font("Trebuchet MS", 10.5F, System.Drawing.FontStyle.Bold);
            lblGroupSection.ForeColor = System.Drawing.Color.FromArgb(31, 41, 55);
            lblGroupSection.Location = new System.Drawing.Point(20, 16);
            lblGroupSection.Name = "lblGroupSection";
            lblGroupSection.Size = new System.Drawing.Size(188, 22);
            lblGroupSection.Text = "Entra Group Sync";

            lblGroupSub.AutoSize = true;
            lblGroupSub.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            lblGroupSub.ForeColor = System.Drawing.Color.FromArgb(107, 114, 128);
            lblGroupSub.Location = new System.Drawing.Point(22, 42);
            lblGroupSub.Name = "lblGroupSub";
            lblGroupSub.Size = new System.Drawing.Size(360, 18);
            lblGroupSub.Text = "Fetch Entra members and bulk sync them via WhoAmI.";

            // 
            // rbClientSecret
            // 
            this.rbClientSecret.AutoSize = true;
            this.rbClientSecret.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            this.rbClientSecret.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            this.rbClientSecret.Location = new System.Drawing.Point(22, 140);
            this.rbClientSecret.Name = "rbClientSecret";
            this.rbClientSecret.Size = new System.Drawing.Size(146, 22);
            this.rbClientSecret.Text = "Client credentials";
            this.rbClientSecret.Checked = true;
            this.rbClientSecret.CheckedChanged += new System.EventHandler(this.AuthModeChanged);

            // 
            // rbDeviceCode
            // 
            this.rbDeviceCode.AutoSize = true;
            this.rbDeviceCode.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Regular);
            this.rbDeviceCode.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            this.rbDeviceCode.Location = new System.Drawing.Point(190, 140);
            this.rbDeviceCode.Name = "rbDeviceCode";
            this.rbDeviceCode.Size = new System.Drawing.Size(193, 22);
            this.rbDeviceCode.Text = "Browser login (no secret)";
            this.rbDeviceCode.CheckedChanged += new System.EventHandler(this.AuthModeChanged);

            // 
            // chkShowAppFields
            // 
            this.chkShowAppFields.AutoSize = true;
            this.chkShowAppFields.Font = new System.Drawing.Font("Trebuchet MS", 8.5F, System.Drawing.FontStyle.Regular);
            this.chkShowAppFields.ForeColor = System.Drawing.Color.FromArgb(107, 114, 128);
            this.chkShowAppFields.Location = new System.Drawing.Point(22, 164);
            this.chkShowAppFields.Name = "chkShowAppFields";
            this.chkShowAppFields.Size = new System.Drawing.Size(214, 21);
            this.chkShowAppFields.Text = "Show app registration fields";
            this.chkShowAppFields.Checked = true;
            this.chkShowAppFields.CheckedChanged += new System.EventHandler(this.AuthModeChanged);

            lblTenantId.AutoSize = true;
            lblTenantId.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            lblTenantId.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblTenantId.Location = new System.Drawing.Point(22, 192);
            lblTenantId.Text = "Tenant ID";

            lblClientId.AutoSize = true;
            lblClientId.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            lblClientId.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblClientId.Location = new System.Drawing.Point(22, 236);
            lblClientId.Text = "Client ID";

            lblClientSecret.AutoSize = true;
            lblClientSecret.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            lblClientSecret.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblClientSecret.Location = new System.Drawing.Point(22, 280);
            lblClientSecret.Text = "Client Secret";

            lblGroupId.AutoSize = true;
            lblGroupId.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            lblGroupId.ForeColor = System.Drawing.Color.FromArgb(55, 65, 81);
            lblGroupId.Location = new System.Drawing.Point(22, 74);
            lblGroupId.Text = "Group ID";

            lblGroupHelp.Font = new System.Drawing.Font("Trebuchet MS", 8.5F, System.Drawing.FontStyle.Italic);
            lblGroupHelp.ForeColor = System.Drawing.Color.FromArgb(107, 114, 128);
            lblGroupHelp.Location = new System.Drawing.Point(22, 356);
            lblGroupHelp.Size = new System.Drawing.Size(500, 56);
            lblGroupHelp.Text = "Permissions: GroupMember.Read.All and User.Read.All." + Environment.NewLine +
                               "Browser login uses delegated permissions; client credentials uses application permissions.";

            // 
            // txtTenantId
            // 
            this.txtTenantId.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Regular);
            this.txtTenantId.Location = new System.Drawing.Point(22, 210);
            this.txtTenantId.Size = new System.Drawing.Size(360, 26);

            // 
            // txtClientId
            // 
            this.txtClientId.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Regular);
            this.txtClientId.Location = new System.Drawing.Point(22, 254);
            this.txtClientId.Size = new System.Drawing.Size(360, 26);

            // 
            // txtClientSecret
            // 
            this.txtClientSecret.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Regular);
            this.txtClientSecret.Location = new System.Drawing.Point(22, 298);
            this.txtClientSecret.Size = new System.Drawing.Size(360, 26);
            this.txtClientSecret.UseSystemPasswordChar = true;

            // 
            // txtGroupId
            // 
            this.txtGroupId.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Regular);
            this.txtGroupId.Location = new System.Drawing.Point(22, 92);
            this.txtGroupId.Size = new System.Drawing.Size(360, 26);

            // 
            // btnFetchGroup
            // 
            this.btnFetchGroup.BackColor = System.Drawing.Color.FromArgb(59, 130, 246);
            this.btnFetchGroup.ForeColor = System.Drawing.Color.White;
            this.btnFetchGroup.FlatStyle = FlatStyle.Flat;
            this.btnFetchGroup.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            this.btnFetchGroup.Location = new System.Drawing.Point(400, 90);
            this.btnFetchGroup.Name = "btnFetchGroup";
            this.btnFetchGroup.Size = new System.Drawing.Size(118, 28);
            this.btnFetchGroup.TabIndex = 11;
            this.btnFetchGroup.Text = "Fetch Members";
            this.btnFetchGroup.FlatAppearance.BorderSize = 0;
            this.btnFetchGroup.Click += new System.EventHandler(this.BtnFetchGroup_Click);

            // 
            // btnBeginGroupSync
            // 
            this.btnBeginGroupSync.BackColor = System.Drawing.Color.FromArgb(15, 118, 110);
            this.btnBeginGroupSync.ForeColor = System.Drawing.Color.White;
            this.btnBeginGroupSync.FlatStyle = FlatStyle.Flat;
            this.btnBeginGroupSync.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold);
            this.btnBeginGroupSync.Location = new System.Drawing.Point(400, 128);
            this.btnBeginGroupSync.Name = "btnBeginGroupSync";
            this.btnBeginGroupSync.Size = new System.Drawing.Size(118, 28);
            this.btnBeginGroupSync.TabIndex = 12;
            this.btnBeginGroupSync.Text = "Begin Syncing";
            this.btnBeginGroupSync.FlatAppearance.BorderSize = 0;
            this.btnBeginGroupSync.Enabled = false;
            this.btnBeginGroupSync.Visible = false;
            this.btnBeginGroupSync.Click += new System.EventHandler(this.BtnBeginGroupSync_Click);

            // 
            // lblGroupStatus
            // 
            this.lblGroupStatus.AutoSize = true;
            this.lblGroupStatus.Font = new System.Drawing.Font("Trebuchet MS", 9.5F, System.Drawing.FontStyle.Italic);
            this.lblGroupStatus.ForeColor = System.Drawing.Color.FromArgb(75, 85, 99);
            this.lblGroupStatus.Location = new System.Drawing.Point(22, 424);
            this.lblGroupStatus.Name = "lblGroupStatus";
            this.lblGroupStatus.Size = new System.Drawing.Size(0, 20);

            // 
            // rightSplit
            // 
            rightSplit.Dock = DockStyle.Fill;
            rightSplit.Orientation = Orientation.Horizontal;
            rightSplit.SplitterWidth = 6;
            rightSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);
            rightSplit.SplitterDistance = 320;

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
            gridSplit.SplitterDistance = 360;

            // 
            // lowerSplit
            // 
            lowerSplit.Dock = DockStyle.Fill;
            lowerSplit.Orientation = Orientation.Horizontal;
            lowerSplit.SplitterWidth = 6;
            lowerSplit.BackColor = System.Drawing.Color.FromArgb(232, 228, 222);
            lowerSplit.SplitterDistance = 220;

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

            // 
            // groupPanel
            // 
            groupPanel.BackColor = Color.White;
            groupPanel.BorderStyle = BorderStyle.FixedSingle;
            groupPanel.Padding = new Padding(10);
            groupPanel.Dock = DockStyle.Fill;

            // 
            // lblGroupTitle
            // 
            lblGroupTitle.Text = "Group Members";
            lblGroupTitle.Font = new Font("Trebuchet MS", 10.5F, FontStyle.Bold);
            lblGroupTitle.ForeColor = Color.FromArgb(31, 41, 55);
            lblGroupTitle.Dock = DockStyle.Top;
            lblGroupTitle.Height = 22;

            // 
            // lblGroupResultsHint
            // 
            this.lblGroupResultsHint.Text = "No group members loaded yet.";
            this.lblGroupResultsHint.Font = new Font("Trebuchet MS", 8.5F, FontStyle.Italic);
            this.lblGroupResultsHint.ForeColor = Color.FromArgb(107, 114, 128);
            this.lblGroupResultsHint.Dock = DockStyle.Top;
            this.lblGroupResultsHint.Height = 20;

            // 
            // dgvGroupMembers
            // 
            ConfigureGrid(this.dgvGroupMembers, new[] { "Display Name", "UPN", "Email" });
            this.dgvGroupMembers.Dock = DockStyle.Fill;

            summaryPanel.Controls.Add(this.lblSummary);
            summaryPanel.Controls.Add(lblSummaryTitle);

            rolesPanel.Controls.Add(this.dgvRoles);
            rolesPanel.Controls.Add(lblRolesTitle);

            teamsPanel.Controls.Add(this.dgvTeams);
            teamsPanel.Controls.Add(lblTeamsTitle);

            groupPanel.Controls.Add(this.dgvGroupMembers);
            groupPanel.Controls.Add(this.lblGroupResultsHint);
            groupPanel.Controls.Add(lblGroupTitle);

            gridSplit.Panel1.Controls.Add(rolesPanel);
            gridSplit.Panel2.Controls.Add(teamsPanel);

            lowerSplit.Panel1.Controls.Add(summaryPanel);
            lowerSplit.Panel2.Controls.Add(gridSplit);

            rightSplit.Panel1.Controls.Add(lowerSplit);
            rightSplit.Panel2.Controls.Add(groupPanel);

            leftSplit.Panel1.Controls.Add(userCard);
            leftSplit.Panel2.Controls.Add(groupCard);
            rootSplit.Panel1.Controls.Add(leftSplit);
            rootSplit.Panel2.Controls.Add(rightSplit);

            // 
            // EntraSyncControl
            // 
            this.BackColor = System.Drawing.Color.FromArgb(248, 245, 240);
            this.Controls.Add(rootSplit);
            this.Name = "EntraSyncControl";
            this.Size = new System.Drawing.Size(980, 680);

            userCard.ResumeLayout(false);
            userCard.PerformLayout();
            groupCard.ResumeLayout(false);
            groupCard.PerformLayout();
            rootSplit.Panel1.ResumeLayout(false);
            rootSplit.Panel2.ResumeLayout(false);
            leftSplit.Panel1.ResumeLayout(false);
            leftSplit.Panel2.ResumeLayout(false);
            rightSplit.Panel1.ResumeLayout(false);
            rightSplit.Panel2.ResumeLayout(false);
            gridSplit.Panel1.ResumeLayout(false);
            gridSplit.Panel2.ResumeLayout(false);
            lowerSplit.Panel1.ResumeLayout(false);
            lowerSplit.Panel2.ResumeLayout(false);
            rootSplit.EndInit();
            leftSplit.EndInit();
            rightSplit.EndInit();
            gridSplit.EndInit();
            lowerSplit.EndInit();
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
                        lblUserStatus.Text = status;
                        lblUserStatus.ForeColor = Color.DarkBlue;
                    }
                },
                Work = (worker, args) =>
                {
                    var details = ResolveUserAndSync(email, worker);
                    args.Result = details;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        lblUserStatus.Text = "Error";
                        lblUserStatus.ForeColor = Color.Red;
                        MessageBox.Show(args.Error.Message, "Sync Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        var details = args.Result as SyncResultDetails;
                        lblUserStatus.Text = "Sync Complete";
                        lblUserStatus.ForeColor = Color.Green;
                        if (details != null)
                        {
                            UpdateDetails(details);
                        }
                    }
                }
            });
        }

        private void BtnFetchGroup_Click(object sender, EventArgs e)
        {
            if (!ValidateGroupInputs())
            {
                return;
            }

            btnBeginGroupSync.Visible = false;
            btnBeginGroupSync.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Fetching group members...",
                ProgressChanged = (progress) =>
                {
                    if (progress?.UserState is string status)
                    {
                        lblGroupStatus.Text = status;
                        lblGroupStatus.ForeColor = Color.DarkBlue;
                    }
                },
                Work = (worker, args) =>
                {
                    var members = FetchGroupMembers(
                        txtTenantId.Text.Trim(),
                        txtClientId.Text.Trim(),
                        txtClientSecret.Text,
                        txtGroupId.Text.Trim(),
                        worker
                    );

                    args.Result = members;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        lblGroupStatus.Text = "Error";
                        lblGroupStatus.ForeColor = Color.Red;
                        MessageBox.Show(args.Error.Message, "Group Fetch Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    groupMembers = args.Result as List<GraphGroupMember> ?? new List<GraphGroupMember>();
                    UpdateGroupMembers(groupMembers);
                    btnBeginGroupSync.Enabled = groupMembers.Count > 0;
                    btnBeginGroupSync.Visible = groupMembers.Count > 0;
                    lblGroupStatus.Text = $"Fetched {groupMembers.Count} members";
                    lblGroupStatus.ForeColor = Color.DarkGreen;
                }
            });
        }

        private void BtnBeginGroupSync_Click(object sender, EventArgs e)
        {
            if (groupMembers.Count == 0)
            {
                MessageBox.Show("Fetch group members first.", "No Members", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Syncing group members...",
                ProgressChanged = (progress) =>
                {
                    if (progress?.UserState is string status)
                    {
                        lblGroupStatus.Text = status;
                        lblGroupStatus.ForeColor = Color.DarkBlue;
                    }
                },
                Work = (worker, args) =>
                {
                    var results = new List<GroupSyncResult>();
                    int total = groupMembers.Count;
                    int index = 0;

                    foreach (var member in groupMembers)
                    {
                        index++;
                        string email = SelectMemberEmail(member);
                        if (string.IsNullOrWhiteSpace(email))
                        {
                            results.Add(new GroupSyncResult(member, false, "Missing email/UPN"));
                            worker.ReportProgress(index * 100 / total, $"Skipped {member.DisplayName} (missing email)");
                            continue;
                        }

                        try
                        {
                            ResolveUserAndSync(email, worker);
                            results.Add(new GroupSyncResult(member, true, null));
                            worker.ReportProgress(index * 100 / total, $"Synced {member.DisplayName}");
                        }
                        catch (Exception ex)
                        {
                            results.Add(new GroupSyncResult(member, false, ex.Message));
                            worker.ReportProgress(index * 100 / total, $"Failed {member.DisplayName}");
                        }
                    }

                    args.Result = results;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        lblGroupStatus.Text = "Error";
                        lblGroupStatus.ForeColor = Color.Red;
                        MessageBox.Show(args.Error.Message, "Group Sync Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var results = args.Result as List<GroupSyncResult> ?? new List<GroupSyncResult>();
                    int ok = 0;
                    foreach (var result in results)
                    {
                        if (result.Success)
                        {
                            ok++;
                        }
                    }

                    lblGroupStatus.Text = $"Group sync complete: {ok}/{results.Count} succeeded";
                    lblGroupStatus.ForeColor = Color.Green;
                }
            });
        }

        private SyncResultDetails ResolveUserAndSync(string email, System.ComponentModel.BackgroundWorker worker)
        {
            var query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("systemuserid", "fullname");
            query.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, email);

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
                throw new Exception($"User with email/UPN '{email}' not found in Dataverse.");
            }

            var targetUser = result.Entities[0];
            Guid targetUserId = targetUser.Id;
            string fullName = targetUser.GetAttributeValue<string>("fullname");

            worker?.ReportProgress(20, $"User found: {fullName} ({targetUserId})");

            var details = new SyncResultDetails
            {
                FullName = fullName,
                Email = email,
                UserId = targetUserId,
                Roles = LoadUserRoles(Service, targetUserId),
                Teams = LoadUserTeams(Service, targetUserId)
            };

            var svc = this.ConnectionDetail.GetCrmServiceClient();
            if (svc == null)
            {
                throw new Exception("Could not create an impersonated service client.");
            }

            svc.CallerId = targetUserId;
            var response = svc.Execute(new WhoAmIRequest());
            details.WhoAmIUserId = ((WhoAmIResponse)response).UserId;

            return details;
        }

        private static string SelectMemberEmail(GraphGroupMember member)
        {
            if (!string.IsNullOrWhiteSpace(member.UserPrincipalName))
            {
                return member.UserPrincipalName;
            }

            if (!string.IsNullOrWhiteSpace(member.Mail))
            {
                return member.Mail;
            }

            return null;
        }

        private List<GraphGroupMember> FetchGroupMembers(string tenantId, string clientId, string clientSecret, string groupId, System.ComponentModel.BackgroundWorker worker)
        {
            var token = AcquireToken(tenantId, clientId, clientSecret, rbDeviceCode.Checked);
            var members = new List<GraphGroupMember>();
            string url = $"https://graph.microsoft.com/v1.0/groups/{groupId}/members?$select=id,displayName,userPrincipalName,mail";

            using (var http = new HttpClient())
            {
                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                while (!string.IsNullOrWhiteSpace(url))
                {
                    var response = http.GetAsync(url).Result;
                    var payload = response.Content.ReadAsStringAsync().Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"Graph error: {response.StatusCode} {payload}");
                    }

                    var page = ParseGraphPage(payload);
                    members.AddRange(page.Members);
                    url = page.NextLink;
                    worker?.ReportProgress(0, $"Loaded {members.Count} members...");
                }
            }

            return members;
        }

        private static GraphPage ParseGraphPage(string json)
        {
            var serializer = new JavaScriptSerializer();
            var root = serializer.DeserializeObject(json) as Dictionary<string, object>;
            var page = new GraphPage();

            if (root == null)
            {
                return page;
            }

            if (root.TryGetValue("@odata.nextLink", out var next))
            {
                page.NextLink = next as string;
            }

            if (root.TryGetValue("value", out var valueObj) && valueObj is object[] values)
            {
                foreach (var item in values)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null)
                    {
                        continue;
                    }

                    page.Members.Add(new GraphGroupMember
                    {
                        Id = dict.TryGetValue("id", out var idVal) ? idVal as string : null,
                        DisplayName = dict.TryGetValue("displayName", out var dnVal) ? dnVal as string : null,
                        UserPrincipalName = dict.TryGetValue("userPrincipalName", out var upnVal) ? upnVal as string : null,
                        Mail = dict.TryGetValue("mail", out var mailVal) ? mailVal as string : null
                    });
                }
            }

            return page;
        }
        private string AcquireToken(string tenantId, string clientId, string clientSecret, bool useDeviceCode)
        {
            if (useDeviceCode)
            {
                return AcquireTokenDeviceCode(tenantId, clientId);
            }

            string url = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
            using (var http = new HttpClient())
            {
                var content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    { "client_id", clientId },
                    { "client_secret", clientSecret },
                    { "scope", "https://graph.microsoft.com/.default" },
                    { "grant_type", "client_credentials" }
                });

                var response = http.PostAsync(url, content).Result;
                var payload = response.Content.ReadAsStringAsync().Result;
                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Token error: {response.StatusCode} {payload}");
                }

                var serializer = new JavaScriptSerializer();
                var root = serializer.DeserializeObject(payload) as Dictionary<string, object>;
                if (root != null && root.TryGetValue("access_token", out var token))
                {
                    return token as string;
                }
            }

            throw new Exception("Could not acquire access token.");
        }

        private string AcquireTokenDeviceCode(string tenantId, string clientId)
        {
            string deviceUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/devicecode";
            string tokenUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

            using (var http = new HttpClient())
            {
                var deviceRequest = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    { "client_id", clientId },
                    { "scope", "https://graph.microsoft.com/.default" }
                });

                var deviceResponse = http.PostAsync(deviceUrl, deviceRequest).Result;
                var devicePayload = deviceResponse.Content.ReadAsStringAsync().Result;
                if (!deviceResponse.IsSuccessStatusCode)
                {
                    throw new Exception($"Device code error: {deviceResponse.StatusCode} {devicePayload}");
                }

                var serializer = new JavaScriptSerializer();
                var deviceRoot = serializer.DeserializeObject(devicePayload) as Dictionary<string, object>;
                if (deviceRoot == null || !deviceRoot.TryGetValue("device_code", out var deviceCodeObj))
                {
                    throw new Exception("Device code response missing device_code.");
                }

                string deviceCode = deviceCodeObj as string;
                string userCode = deviceRoot.TryGetValue("user_code", out var userCodeObj) ? userCodeObj as string : null;
                string verificationUri = deviceRoot.TryGetValue("verification_uri", out var verObj) ? verObj as string : null;
                string message = deviceRoot.TryGetValue("message", out var msgObj) ? msgObj as string : null;
                int interval = deviceRoot.TryGetValue("interval", out var intObj) ? Convert.ToInt32(intObj) : 5;

                if (!string.IsNullOrWhiteSpace(message))
                {
                    Clipboard.SetText(userCode ?? string.Empty);
                    MessageBox.Show(message + Environment.NewLine + "Code copied to clipboard.", "Browser Login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Go to {verificationUri} and enter code: {userCode}", "Browser Login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (!string.IsNullOrWhiteSpace(verificationUri))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(verificationUri) { UseShellExecute = true });
                    }
                    catch
                    {
                        // If the browser cannot be opened automatically, the MessageBox already showed the URL.
                    }
                }

                var pollContent = new Dictionary<string, string>
                {
                    { "grant_type", "urn:ietf:params:oauth:grant-type:device_code" },
                    { "client_id", clientId },
                    { "device_code", deviceCode }
                };

                while (true)
                {
                    Thread.Sleep(interval * 1000);
                    var tokenResponse = http.PostAsync(tokenUrl, new FormUrlEncodedContent(pollContent)).Result;
                    var tokenPayload = tokenResponse.Content.ReadAsStringAsync().Result;

                    if (tokenResponse.IsSuccessStatusCode)
                    {
                        var tokenRoot = serializer.DeserializeObject(tokenPayload) as Dictionary<string, object>;
                        if (tokenRoot != null && tokenRoot.TryGetValue("access_token", out var token))
                        {
                            return token as string;
                        }
                    }
                    else if (tokenPayload.Contains("authorization_pending"))
                    {
                        continue;
                    }
                    else if (tokenPayload.Contains("slow_down"))
                    {
                        interval += 5;
                        continue;
                    }
                    else
                    {
                        throw new Exception($"Token error: {tokenResponse.StatusCode} {tokenPayload}");
                    }
                }
            }
        }

        private bool ValidateGroupInputs()
        {
            if (string.IsNullOrWhiteSpace(txtTenantId.Text) ||
                string.IsNullOrWhiteSpace(txtClientId.Text) ||
                string.IsNullOrWhiteSpace(txtGroupId.Text))
            {
                if (rbDeviceCode.Checked && !chkShowAppFields.Checked)
                {
                    MessageBox.Show("Tenant ID and Client ID are required for browser login. Enable \"Show app registration fields\" to enter them.", "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Tenant ID, Client ID, and Group ID are required.", "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return false;
            }

            if (rbClientSecret.Checked && string.IsNullOrWhiteSpace(txtClientSecret.Text))
            {
                MessageBox.Show("Client Secret is required for client credentials.", "Missing Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void AuthModeChanged(object sender, EventArgs e)
        {
            if (rbClientSecret.Checked)
            {
                chkShowAppFields.Checked = true;
                chkShowAppFields.Enabled = false;
            }
            else
            {
                if (!chkShowAppFields.Enabled)
                {
                    chkShowAppFields.Checked = false;
                }
                chkShowAppFields.Enabled = true;
            }

            bool showAppFields = rbClientSecret.Checked || chkShowAppFields.Checked;
            lblTenantId.Visible = showAppFields;
            txtTenantId.Visible = showAppFields;
            lblClientId.Visible = showAppFields;
            txtClientId.Visible = showAppFields;

            bool showSecret = rbClientSecret.Checked;
            lblClientSecret.Visible = showSecret;
            txtClientSecret.Visible = showSecret;
            txtClientSecret.Enabled = showSecret;
            txtClientSecret.UseSystemPasswordChar = showSecret;
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

        private void UpdateGroupMembers(List<GraphGroupMember> members)
        {
            dgvGroupMembers.Rows.Clear();
            if (members.Count == 0)
            {
                dgvGroupMembers.Rows.Add("(none)", string.Empty, string.Empty);
                lblGroupResultsHint.Text = "No members found for this group.";
                return;
            }

            foreach (var member in members)
            {
                dgvGroupMembers.Rows.Add(member.DisplayName ?? string.Empty, member.UserPrincipalName ?? string.Empty, member.Mail ?? string.Empty);
            }

            lblGroupResultsHint.Text = $"{members.Count} members loaded.";
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

        private class GraphPage
        {
            public string NextLink { get; set; }
            public List<GraphGroupMember> Members { get; } = new List<GraphGroupMember>();
        }

        private class GroupSyncResult
        {
            public GraphGroupMember Member { get; }
            public bool Success { get; }
            public string Error { get; }

            public GroupSyncResult(GraphGroupMember member, bool success, string error)
            {
                Member = member;
                Success = success;
                Error = error;
            }
        }
    }
}
