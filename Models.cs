using System;

namespace BONDVERSE
{
    public class PortfolioEntry
    {
        public string PortfolioName { get; set; }
        public string InvestorName { get; set; }
        public DateTime TransactionDate { get; set; }
        public double FV { get; set; }
        public int Quantity { get; set; }
        public string BondName { get; set; }
        public double CouponRate { get; set; }
        public double ChequeAmount { get; set; }
        public string Frequency { get; set; }
        public DateTime MaturityDate { get; set; }
    }
}
