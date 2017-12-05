using Compare.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Compare.Models
{
    public class Order
    {
        public string Firstname { get; set; }

        public string Lastname { get; set; }

        public MailSentStatus Status { get; set; }

        public int ChosenImage { get; set; }

        public string Email { get; set; }

        public double TotalAmount { get; set; }

        public string ImageFile { get; set; }

        public string OrderProduct { get; set; }

        public string OrderColor { get; set; }

        public string OrderImagePackage { get; set; }

        public string OrderExtraImagePackage { get; set; }
    }
}
