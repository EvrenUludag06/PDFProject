namespace PDFProject
{
    public class RootObject
    {
        public Bios Bios { get; set; }
        public Cpu Cpu { get; set; }
        public Memory Memory { get; set; }
        public NetworkInterface[] NetworkInterface { get; set; }
        public Gpu[] Gpu { get; set; }
        public CaptureCardObject[] CaptureCard { get; set; }
    }
}