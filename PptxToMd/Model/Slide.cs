using System;
using System.Collections.Generic;

namespace PptxToMd.Model
{
    /// <summary>
    /// Top-level model representation of a slide
    /// </summary>
    public class Slide
    {
        public Slide()
        {
            Titles = new List<string>();
            SubTitles = new List<string>();
            Bullets = new List<Bullet>();
            Images = new List<Image>();
            ID = Guid.NewGuid();
        }

        public List<string> Titles { get; set; }
        public List<string> SubTitles { get; set; }
        public List<Bullet> Bullets { get; set; }
        public List<Image> Images { get; set; }
        public string Notes { get; set; }
        public Guid ID { get; }
    }
}
