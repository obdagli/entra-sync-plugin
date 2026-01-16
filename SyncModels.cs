using System;
using System.Collections.Generic;

namespace EntraSyncPlugin
{
    public class SyncResultDetails
    {
        public string FullName { get; set; }
        public string Email { get; set; }
        public Guid UserId { get; set; }
        public Guid? WhoAmIUserId { get; set; }
        public List<RoleInfo> Roles { get; set; }
        public List<TeamInfo> Teams { get; set; }
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
}
