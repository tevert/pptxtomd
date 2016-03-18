using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PptxToMd.Model;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PptxToMd.Parsers
{
    /// <summary>
    /// The parser for MS Powerpoint files (PPTX). Uses OpenXML on the backend.
    /// </summary>
    /// <seealso cref="PptxToMd.Parsers.ISlideParser" />
    public class PptxSlideParser : ISlideParser
    {
        private PresentationDocument doc;

        /// <summary>
        /// Initializes a new instance of the <see cref="PptxSlideParser"/> class.
        /// 
        /// The file path is expected to point to a valid PPTX file.
        /// 
        /// After this class is disposed, future calls to Parse() may fail. Also, attempting to
        /// access streamed data from the parsed slides after disposing may fail.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        public PptxSlideParser(string filePath)
        {
            doc = PresentationDocument.Open(filePath, true);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            doc.Dispose();
        }

        /// <summary>
        /// Parses the data into a list of slides.
        /// </summary>
        /// <returns>
        /// A list of parsed slides.
        /// </returns>
        public List<Model.Slide> Parse()
        {
            var parsedSlides = new List<Model.Slide>();
            PresentationPart presentationPart = doc.PresentationPart;
            // Verify that the presentation part and presentation exist.
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get the Presentation object from the presentation part.
                Presentation presentation = presentationPart.Presentation;

                // Verify that the slide ID list exists.
                if (presentation.SlideIdList != null)
                {
                    // Loop on the collection of slide IDs from the slide ID list.
                    foreach (SlideId slideId in presentation.SlideIdList.ChildElements)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = slideId.RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                        parsedSlides.Add(ParseSlide(slidePart));
                    }
                }
            }

            return parsedSlides;
        }

        /// <summary>
        /// Parses the slide into one of our models.
        /// </summary>
        /// <param name="slidePart">The slide part.</param>
        /// <returns>The populated model.</returns>
        /// <exception cref="ArgumentNullException">slidePart</exception>
        private static Model.Slide ParseSlide(SlidePart slidePart)
        {
            // Verify that the slide part exists.
            if (slidePart?.Slide == null)
            {
                throw new ArgumentNullException("slidePart");
            }

            // If the slide exists...
            var parsedSlide = new Model.Slide();
            if (slidePart.Slide != null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (var entity in slidePart.Slide.Descendants<Shape>())
                {
                    var placeholderShape = entity.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.PlaceholderShape;
                    if (placeholderShape != null && placeholderShape.Type != PlaceholderValues.Body)
                    {
                        // Handle titles
                        switch ((PlaceholderValues)placeholderShape.Type)
                        {
                            case PlaceholderValues.Title:
                            case PlaceholderValues.CenteredTitle:
                                parsedSlide.Titles.Add(entity.GetFirstChild<TextBody>().InnerText);
                                break;
                            case PlaceholderValues.SubTitle:
                                parsedSlide.SubTitles.Add(entity.GetFirstChild<TextBody>().InnerText);
                                break;
                            default:
                                // Skip for now
                                break;
                        }
                    }
                    else
                    {
                        // Lets just assume this is a random text box and see what we can strip from it.
                        var paragraphs = entity.GetFirstChild<TextBody>().Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>();
                        foreach (var paragraph in paragraphs)
                        {
                            parsedSlide.Bullets.Add(new Bullet()
                            {
                                Text = paragraph.InnerText,
                                Level = paragraph.ParagraphProperties?.Level != null ?
                                        paragraph.ParagraphProperties.Level.Value : 0
                            });
                        }
                    }
                }
            }

            // Now strip the images. 
            var pictures = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>();
            foreach (var img in pictures)
            {
                var embed = img.BlipFill.Blip.Embed.Value;
                var imgPart = slidePart.GetPartById(embed);
                parsedSlide.Images.Add(new Image()
                {
                    Filename = imgPart.Uri.OriginalString,
                    ContentType = imgPart.ContentType,
                    Data = System.Drawing.Image.FromStream(imgPart.GetStream())
                });
            }

            // Lastly, the speaker notes
            var notesPart = slidePart.NotesSlidePart;
            if (notesPart != null)
            {
                var noteShape = notesPart
                    .NotesSlide?
                    .CommonSlideData?
                    .ShapeTree?
                    .Descendants<Shape>()?
                    .FirstOrDefault(s =>
                        s.NonVisualShapeProperties?
                        .ApplicationNonVisualDrawingProperties?
                        .PlaceholderShape?
                        .Type == PlaceholderValues.Body);

                if (noteShape != null)
                {
                    parsedSlide.Notes = noteShape.TextBody.InnerText;
                }
            }

            return parsedSlide;
        }
    }
}
