namespace PDFProject
{
    public class Memory
    {
        public Cards[] Cards { get; set; }
        public long FreePhysicalMemory { get; set; }
        public string FreePhysicalMemoryFormat { get; set; }
        public long TotalPhysicalMemory { get; set; }
        public string TotalPhysicalMemoryFormat { get; set; }
    }
}