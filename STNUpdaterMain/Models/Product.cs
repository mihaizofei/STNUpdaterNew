namespace STNUpdater.Models
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public int CategoryId { get; set; }
        public string Maker { get; set; }
        public int MakerId { get; set; }
        public string Model { get; set; }
        public string Code { get; set; }
        public string ShortDescription { get; set; }
        public string Warranty { get; set; }
    }
}
