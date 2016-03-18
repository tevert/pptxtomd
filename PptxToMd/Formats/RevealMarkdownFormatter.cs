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
        /// <returns>A markdown formatted string of the slide.</returns>
        public string Convert(Slide data)
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

            output.Append("\n");

            foreach (var bullet in data.Bullets)
            {
                output.Append($"*{new string(' ', bullet.Level * 2)} {bullet.Text}\n");
            }

            output.Append("\n");

            for (int i = 0; i < data.Images.Count; i++)
            {
                output.Append($"![Image](./{data.ID.ToString()}-img{i + 1}.jpg)\n");
            }

            output.Append("\n");

            if (!String.IsNullOrWhiteSpace(data.Notes))
            {
                output.Append($"Notes: {data.Notes}\n");
                output.Append("\n");
            }
            output.Append("\n\n");

            return output.ToString();
        }
    }
}
