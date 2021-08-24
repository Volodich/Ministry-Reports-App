using MinistryReports.Models;
using MinistryReports.Models.S21;

namespace MinistryReports.ViewModels
{
    public class UserSettings
    {
        // Basic Information
        public string UserName { get; set; }

        // JWBook Settings
        public JWBookSettings JWBookSettings { get; set; }

        // S-21 Settings
        public S21Settings S21Settings { get; set; }
    }
}
