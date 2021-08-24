using System;

namespace MinistryReports.Models.JWBook
{
    public class SheetSettings
    {
        public SheetSettings()
        {
            Id = new Guid();
        }
        public Guid Id { get; set; }

        public string Name { get; set; }
        public string SheetId { get; set; }
    }
}
