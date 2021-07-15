using System;
using System.Collections.Generic;

namespace ExcelUpload.WebApi.Models
{
    public partial class EmployeeAddress
    {
        public int AddressId { get; set; }
        public string Country { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public int EmployeeId { get; set; }

        public virtual Employees Employee { get; set; }
    }
}
