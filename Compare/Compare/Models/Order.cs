using Compare.Enums;

namespace Compare.Models
{
    public class Order
    {
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public string FullName { get; set; }
        public MailSentStatus Status { get; set; }
        public string ChosenImage { get; set; }
        public string Email { get; set; }
        public decimal TotalAmount { get; set; }
        public string ImageFile { get; set; }
        public string OrderProduct { get; set; }
        public string OrderImagePackage { get; set; }
        public string OrderImagePackageColor { get; set; }
        public string OrderExtraImagePackage { get; set; }
        public string OrderExtraImagePackageColor { get; set; }
    }
}
