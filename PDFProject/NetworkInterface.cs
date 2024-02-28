namespace PDFProject
{
    public class NetworkInterface
    {
        public string Description { get; set; }
        public bool DHCP { get; set; }
        public string[] Dns { get; set; }
        public string[] Gateways { get; set; }
        public string Id { get; set; }
        public string[] IPAddresses { get; set; }
        public string MACAddress { get; set; }
        public string Name { get; set; }
        public string[] SubnetMasks { get; set; }
    }
}