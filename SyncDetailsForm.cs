using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace EntraSyncPlugin
{
    public class SyncDetailsForm : Form
    {
        public SyncDetailsForm(SyncResultDetails details)
        {
            Text = "Entra Sync Details";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(900, 620);
            BackColor = Color.FromArgb(246, 243, 238);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 1,
                RowCount = 3
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var header = new Label
            {
                Text = "Entra Sync Results",
                Font = new Font("Trebuchet MS", 18F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            root.Controls.Add(header, 0, 0);

            var userPanel = BuildUserPanel(details);
            root.Controls.Add(userPanel, 0, 1);

            var mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterWidth = 6,
                BackColor = Color.FromArgb(232, 228, 222)
            };

            var rolesSection = BuildSection(
                "Personal Security Roles",
                BuildGrid(details.Roles, new[] { "Role", "Business Unit" }, r => new[] { r.Name, r.BusinessUnit })
            );
            mainSplit.Panel1.Controls.Add(rolesSection);

            var lowerSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6,
                BackColor = Color.FromArgb(232, 228, 222)
            };

            var teamsSection = BuildSection(
                "Assigned Teams",
                BuildGrid(details.Teams, new[] { "Team", "Type", "Business Unit" }, t => new[] { t.Name, t.TeamType, t.BusinessUnit })
            );
            lowerSplit.Panel1.Controls.Add(teamsSection);

            var entraGrid = BuildGrid(details.EntraGroups, new[] { "Group Id", "Display Name" }, g => new[] { g.GroupId, g.DisplayName });
            var entraSection = BuildSection("Entra Groups", entraGrid);
            if (!string.IsNullOrWhiteSpace(details.EntraNote))
            {
                var note = new Label
                {
                    Text = details.EntraNote,
                    Font = new Font("Trebuchet MS", 8.5F, FontStyle.Italic),
                    ForeColor = Color.FromArgb(107, 114, 128),
                    Dock = DockStyle.Bottom,
                    Padding = new Padding(8, 6, 8, 6)
                };
                entraSection.Controls.Add(note);
            }
            lowerSplit.Panel2.Controls.Add(entraSection);

            mainSplit.Panel2.Controls.Add(lowerSplit);
            root.Controls.Add(mainSplit, 0, 2);

            var closeBtn = new Button
            {
                Text = "Close",
                BackColor = Color.FromArgb(15, 118, 110),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Trebuchet MS", 9.5F, FontStyle.Bold),
                Size = new Size(90, 30),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            closeBtn.FlatAppearance.BorderSize = 0;
            closeBtn.Click += (s, e) => Close();

            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            buttonPanel.Controls.Add(closeBtn);
            closeBtn.Location = new Point(buttonPanel.Width - closeBtn.Width - 4, 4);
            buttonPanel.Resize += (s, e) =>
            {
                closeBtn.Location = new Point(buttonPanel.Width - closeBtn.Width - 4, 4);
            };

            Controls.Add(root);
            Controls.Add(buttonPanel);
        }

        private static Panel BuildUserPanel(SyncResultDetails details)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(12)
            };

            var title = new Label
            {
                Text = "User Summary",
                Font = new Font("Trebuchet MS", 11.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Dock = DockStyle.Top,
                Height = 22
            };

            var body = new Label
            {
                Font = new Font("Trebuchet MS", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(55, 65, 81),
                Dock = DockStyle.Fill
            };

            body.Text =
                "Name: " + details.FullName + Environment.NewLine +
                "Email: " + details.Email + Environment.NewLine +
                "UserId: " + details.UserId + Environment.NewLine +
                "WhoAmI: " + (details.WhoAmIUserId.HasValue ? details.WhoAmIUserId.ToString() : "n/a");

            panel.Controls.Add(body);
            panel.Controls.Add(title);
            return panel;
        }

        private static Panel BuildSection(string title, Control content)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            var header = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Font = new Font("Trebuchet MS", 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Height = 22
            };

            content.Dock = DockStyle.Fill;
            panel.Controls.Add(content);
            panel.Controls.Add(header);
            return panel;
        }

        private static Control BuildGrid<T>(List<T> items, string[] headers, Func<T, string[]> rowSelector)
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                Font = new Font("Trebuchet MS", 9.5F, FontStyle.Regular)
            };

            foreach (var header in headers)
            {
                grid.Columns.Add(header, header);
            }

            if (items.Count == 0)
            {
                var row = new string[headers.Length];
                row[0] = "(none)";
                grid.Rows.Add(row);
            }
            else
            {
                foreach (var item in items)
                {
                    grid.Rows.Add(rowSelector(item));
                }
            }

            return grid;
        }
    }

    public class SyncResultDetails
    {
        public string FullName { get; set; }
        public string Email { get; set; }
        public Guid UserId { get; set; }
        public Guid? WhoAmIUserId { get; set; }
        public List<RoleInfo> Roles { get; set; }
        public List<TeamInfo> Teams { get; set; }
        public List<EntraGroupInfo> EntraGroups { get; set; }
        public string EntraNote { get; set; }
    }

    public class RoleInfo
    {
        public string Name { get; set; }
        public string BusinessUnit { get; set; }
    }

    public class TeamInfo
    {
        public string Name { get; set; }
        public string TeamType { get; set; }
        public string BusinessUnit { get; set; }
    }

    public class EntraGroupInfo
    {
        public string GroupId { get; set; }
        public string DisplayName { get; set; }
    }
}
