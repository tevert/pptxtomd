namespace PptxToMd.Model
{
    /// <summary>
    /// Represents a bullet in a bulleted list on a slide. Level indicates indentation level.
    /// </summary>
    public class Bullet
    {
        public string Text { get; set; }
        public int Level { get; set; }
    }
}
