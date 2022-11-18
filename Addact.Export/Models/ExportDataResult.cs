using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Addact.Export.Models
{
    public class ExportDataResult
    {
        public string FileName { get; set; }

        public byte[] Content { get; set; }

        public string MediaType { get; set; }
    }
}