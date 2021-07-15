using System;
using System.Collections.Generic;

namespace ExcelUpload.WebApi.Models
{
    public partial class ExcelTable
    {
        public int ExcelId { get; set; }
        public string PatientId { get; set; }
        public string Name { get; set; }
        public string Age { get; set; }
        public string Address { get; set; }
    }
}
