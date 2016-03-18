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
        /// <returns>A string representation of the slide.</returns>
        string Convert(Slide data);
    }
}
