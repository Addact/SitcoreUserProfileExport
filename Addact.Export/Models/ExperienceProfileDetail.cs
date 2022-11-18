using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Addact.Export.Models
{
    public class ExperienceProfileDetail
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string CompanyName { get; set; }
        public string JobTitle { get; set; }
        public List<string> PageList { get; set; }
        public List<string> PageEventList { get; set; }
    }
}