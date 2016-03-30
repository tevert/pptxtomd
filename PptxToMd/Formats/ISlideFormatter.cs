using PptxToMd.Model;

namespace PptxToMd.Formats
{
    /// <summary>
    /// Interface for slide output formatters. 
    /// </summary>
    public interface ISlideFormatter
    {
        /// <summary>
        /// Converts the slide data to the instance-specific format.
        /// </summary>
        /// <param name="data">The slide data.</param>
        /// <param name="resourcePath">Specifies where any resources external to the slide should be referenced from (example: images)</param>
        /// <returns>A string representation of the slide.</returns>
        string Convert(Slide data, string resourcePath);
    }
}
