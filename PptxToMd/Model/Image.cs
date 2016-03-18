namespace PptxToMd.Model
{
    /// <summary>
    /// Represents an image on a slide.
    /// </summary>
    public class Image
    {
        public string Filename { get; set; }
        public string ContentType { get; set; }
        public System.Drawing.Image Data { get; set; }
    }
}
