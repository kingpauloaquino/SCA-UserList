using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AuditSheet
{
    public class UserList
    {
        public int UserId { get; set; }
        public int WorkerId { get; set; }
        public string Role { get; set; }
        public int Active { get; set; }
        public string Username { get; set; }
        public string PasswordHint { get; set; }
        public string BusinessName { get; set; }
        public string Firstname { get; set; }
        public string Middlename { get; set; }
        public string Lastname { get; set; }
        public string Position { get; set; }
        public string EmailAddress { get; set; }
        public string MobileNumber { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Country { get; set; }
        public string Zipcode { get; set; }
        public string Type { get; set; }
        public string Created_At { get; set; }
        public string Updated_At { get; set; }

    }

    public class ResponseUser
    {
        public int Status { get; set; }
        public string Message { get; set; }
        public int Rows { get; set; }
        public List<UserList> Users { get; set; }
    }
}