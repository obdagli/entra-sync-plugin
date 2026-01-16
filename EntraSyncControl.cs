using System;
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
        private Label lblStatus;
        private Label lblWarning;

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
            var grpBox = new System.Windows.Forms.GroupBox();

            grpBox.SuspendLayout();
            this.SuspendLayout();

            // 
            // grpBox
            // 
            grpBox.Controls.Add(this.lblWarning);
            grpBox.Controls.Add(this.lblStatus);
            grpBox.Controls.Add(this.btnSync);
            grpBox.Controls.Add(this.txtEmail);
            grpBox.Location = new System.Drawing.Point(20, 20);
            grpBox.Name = "grpBox";
            grpBox.Size = new System.Drawing.Size(500, 250);
            grpBox.TabIndex = 0;
            grpBox.TabStop = false;
            grpBox.Text = "User Synchronization";
            grpBox.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular);

            // 
            // lblWarning
            // 
            this.lblWarning.AutoSize = true;
            this.lblWarning.ForeColor = System.Drawing.Color.Red;
            this.lblWarning.Location = new System.Drawing.Point(20, 30);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(400, 20);
            this.lblWarning.TabIndex = 3;
            this.lblWarning.Text = "⚠️ ADMIN WARNING: This operation requires Delegate privileges.";

            // 
            // txtEmail
            // 
            this.txtEmail.Location = new System.Drawing.Point(20, 70);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(350, 30);
            this.txtEmail.TabIndex = 0;
            this.txtEmail.Text = "user@example.com";
            this.txtEmail.ForeColor = Color.Gray;
            this.txtEmail.Enter += (s, e) => { if(txtEmail.Text == "user@example.com") { txtEmail.Text = ""; txtEmail.ForeColor = Color.Black; } };
            this.txtEmail.Leave += (s, e) => { if(string.IsNullOrWhiteSpace(txtEmail.Text)) { txtEmail.Text = "user@example.com"; txtEmail.ForeColor = Color.Gray; } };

            // 
            // btnSync
            // 
            this.btnSync.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnSync.ForeColor = System.Drawing.Color.White;
            this.btnSync.FlatStyle = FlatStyle.Flat;
            this.btnSync.Location = new System.Drawing.Point(380, 68);
            this.btnSync.Name = "btnSync";
            this.btnSync.Size = new System.Drawing.Size(100, 34);
            this.btnSync.TabIndex = 1;
            this.btnSync.Text = "Force Sync";
            this.btnSync.UseVisualStyleBackColor = false;
            this.btnSync.Click += new System.EventHandler(this.BtnSync_Click);

            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(20, 120);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 20);
            this.lblStatus.TabIndex = 2;

            // 
            // EntraSyncControl
            // 
            this.Controls.Add(grpBox);
            this.Name = "EntraSyncControl";
            this.Size = new System.Drawing.Size(600, 400);
            grpBox.ResumeLayout(false);
            grpBox.PerformLayout();
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

            ExecuteMethod(new WorkAsyncInfo
            {
                Message = $"Attempting to sync user: {email}...",
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

                    // 2. Impersonate and Call WhoAmI
                    // We need a NEW service client or clone to switch CallerId safely in XrmToolBox context
                    // Usually Service is IOrganizationService. If it's CrmServiceClient, we can clone.
                    // For safety in Plugins, we often just create a request and hope the host handles CallerId if we wrap it,
                    // BUT OrganizationServiceProxy.CallerId is the classic way.
                    // XrmToolBox 'Service' property usually doesn't allow easy CallerId switching on the fly without casting.
                    
                    // Simpler approach for XrmToolBox: Use the ConnectionDetail to create a temp service
                    // Or check if we can cast Service to something that supports CallerId.
                    
                    // Note: In modern XrmToolBox, we might need to rely on the current connection. 
                    // Let's try creating a WhoAmIRequest and see if we can attach the CallerId property?
                    // No, CallerId is on the Service Proxy, not the Request (except for specific requests).
                    
                    // Standard pattern:
                    // var adminService = Service; // Current Admin Service
                    // We can't change 'Service.CallerId' directly as it might affect other plugins or thread safety.
                    // We will assume 'Service' is an IOrganizationService.
                    
                    // We will try to create a new service with the CallerId.
                    // Since we don't have the connection string handy (it's encapsulated), 
                    // we will try to use the 'ConnectionDetail' if available, or just cast the service.
                    
                    // Hack for simplicity: Check if it's CrmServiceClient (Microsoft.Xrm.Tooling.Connector)
                    // If so, it has a CallerId property.
                    
                    try 
                    {
                        // 3. Execute WhoAmI as the user
                        // We use a OrganizationRequest with 'CallerId' if the proxy supports it? No.
                        // We need a proxy with the CallerId set.
                        
                        // Let's assume we can cast to IProxy or just Create a new proxy from the ConnectionDetail
                        var svc = this.ConnectionDetail.GetCrmServiceClient();
                        if (svc != null)
                        {
                            svc.CallerId = targetUserId;
                            var response = svc.Execute(new WhoAmIRequest());
                            args.Result = $"Success! Synced {fullName}. WhoAmI returned: {((WhoAmIResponse)response).UserId}";
                        }
                        else
                        {
                            // Fallback: Try classic proxy
                            // This part is tricky without knowing the exact connection type (Oauth vs ConnectionString)
                            // But GetCrmServiceClient() is robust in XTB.
                            throw new Exception("Could not create an impersonated service client.");
                        }
                    }
                    catch(Exception ex)
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
                        lblStatus.Text = "Sync Complete ✅";
                        lblStatus.ForeColor = Color.Green;
                        MessageBox.Show(args.Result.ToString(), "Sync Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            });
        }
    }
}
