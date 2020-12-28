using System;
using MinistryReports.Models;

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
