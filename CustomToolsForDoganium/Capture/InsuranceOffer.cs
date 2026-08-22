namespace CustomToolsForDoganium.Capture
{
    /// <summary>Sorgu ekranından ayrıştırılan tek bir sigorta şirketi teklifi.</summary>
    internal sealed class InsuranceOffer
    {
        public string CompanyName { get; set; }
        public int Price { get; set; }
        public string OfferNumber { get; set; }
    }
}
