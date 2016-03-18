using PptxToMd.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptxToMd.Parsers
{
    /// <summary>
    /// Interface for a slide parser
    /// </summary>
    public interface ISlideParser : IDisposable
    {
        /// <summary>
        /// Parses the data into a list of slides.
        /// </summary>
        /// <returns>A list of parsed slides.</returns>
        List<Slide> Parse();
    }
}
