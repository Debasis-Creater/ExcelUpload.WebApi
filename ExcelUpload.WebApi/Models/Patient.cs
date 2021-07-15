using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelUpload.WebApi.Models
{
    public class Patient
    {
        public string PatientId { get; set; }
        public string Name { get; set; }
        public string Age { get; set; }
        public string Address { get; set; }
    }
}
