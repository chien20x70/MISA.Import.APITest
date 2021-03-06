using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MISA.Import.APITest.Model
{
    public class Customer
    {
        public Guid? CustomerId { get; set; }
        public string CustomerCode { get; set; }
        public string FullName { get; set; }
        public string MemberCardCode { get; set; }
        public Guid? CustomerGroupId { get; set; }
        public string CustomerGroupName { get; set; }
        public string PhoneNumber { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string CompanyName { get; set; }
        public string CompanyTaxCode { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string Note { get; set; }
        public string Status { get; set; }

    }
}
