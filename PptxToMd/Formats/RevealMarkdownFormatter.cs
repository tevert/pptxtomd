using System;
using System.Text;
using PptxToMd.Model;

namespace PptxToMd.Formats
{
    /// <summary>
    /// Outputs markdown for the reveal.js library.
    /// 
    /// <see cref="http://lab.hakim.se/reveal-js/"/>
    /// </summary>
    /// <seealso cref="PptxToMd.Formats.ISlideFormatter" />
    public class RevealJsMarkdownFormatter : ISlideFormatter
    {
        /// <summary>
        /// Outputs slide data as markdown.
        /// Some markdown formats can vary - this was generated with reveal.js as the intended consumer:
        /// 
        /// Titles begin with "#".
        /// 
        /// Subtitles begin with "##".
        /// 
        /// Bullet points begin with "*", and are indented 2 spaces for each logical indentation.
        /// 
        /// Notes are prefixed by "Notes:".
        /// 
        /// Slide ends with 3 newlines ("\n\n\n").
        /// </summary>
        /// <param name="data">The slide.</param>
        /// <param name="resourcePath">Specifies where any resources external to the slide should be referenced from (example: images)</param>
        /// <returns>A markdown formatted string of the slide.</returns>
        public string Convert(Slide data, string resourcePath)
        {
            StringBuilder output = new StringBuilder();

            foreach (var title in data.Titles)
            {
                output.Append($"# {title}\n");
            }

            foreach (var title in data.SubTitles)
            {
                output.Append($"## {title}\n");
            }

            if (data.Bullets.Count > 0)
            {
                output.Append("\n");
            }
            foreach (var bullet in data.Bullets)
            {
                output.Append($"{new string(' ', bullet.Level * 2)}* {bullet.Text}\n");
            }

            if (data.Images.Count > 0)
            {
                output.Append("\n");
            }
            for (int i = 0; i < data.Images.Count; i++)
            {
                output.Append($"![Image]({resourcePath}/{data.ID.ToString()}-img{i + 1}.jpg)\n");
            }

            if (!String.IsNullOrWhiteSpace(data.Notes))
            {
                output.Append("\n");
                output.Append($"Notes: {data.Notes}\n");
            }
            output.Append("\n\n\n");

            return output.ToString();
        }
    }
}
