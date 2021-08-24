using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MinistryReports.ViewModels
{
    public class S21InfoPublisherField
    {
        public S21InfoPublisherField()
        {
            Hope144 = "Off";
            HopeOther = "Yes";
            AppointmentPastor = "Off";
            AppointmentMinistryHelp = "Off";
            Pioner = "Off";
            MenGender = "Off";
            WomenGender = "Off";
        }
        internal string Name { get; set; } // Name
        internal string DateBirthday { get; set; } // Date of birth
        internal string DateBaptism { get; set; } // Date immersed
        internal string MenGender { get; set; } // Men - CheckBox1; 
        internal string WomenGender { get; set; } // Women - CheckBox2
        internal string HopeOther { get; set; } // д.о. - CheckBox3; 
        internal string Hope144 { get; set; } // помазанник - CheckBox4;
        internal string AppointmentPastor { get; set; } // старейшина- CheckBox5; 
        internal string AppointmentMinistryHelp { get; set; } // служебный помошник - CheckBox6;             
        internal string Pioner { get; set; } // - CheckBox7;
    }
}
