using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PptxToMd.Formats;
using PptxToMd.Parsers;

namespace PptxToMd
{
    public class Program
    {
        /// <summary>
        /// Main method. First argument should be a filename, otherwise it will prompt the user.
        /// A path ending with ".md" will be considered a single-file output. Otherwise, the output
        /// will be a directory with a .md file for each slide. In both cases, images will be output 
        /// as separate files alongside the slide(s). If no argument was provided, the output will
        /// be printed, and no image files will be written. 
        /// </summary>
        /// <param name="args">The arguments.</param>
        public static void Main(string[] args)
        {
            string file = GetFilename(args);
            string output = GetOutput(args);
            ISlideFormatter format = GetFormat(args);
            
            using (ISlideParser parser = GetParser(file))
            {
                var slides = parser.Parse();
                GenerateOutput(output, slides, format);
            }
        }

        /// <summary>
        /// Gets the format, based on the passed in arguments. Defaults to reveal.js Markdown format.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns>The appropriate output formatter</returns>
        private static ISlideFormatter GetFormat(string[] args)
        {
            if (args.Length < 3)
            {
                return new RevealJsMarkdownFormatter();
            }
            else
            {
                // Currently no others are supported!
                return new RevealJsMarkdownFormatter();
            }
        }

        /// <summary>
        /// Gets the parser, checking to make sure the file extension is supported.
        /// </summary>
        /// <param name="file">The file path.</param>
        /// <returns>A support instance of a slide parser.</returns>
        /// <exception cref="FormatException">Unrecognized file type, currently only supports \*.PPTX\.</exception>
        private static ISlideParser GetParser(string file)
        {
            if (file.EndsWith(".PPTX", StringComparison.CurrentCultureIgnoreCase))
            {
                return new PptxSlideParser(file);
            }
            else
            {
                throw new FormatException("Unrecognized file type, currently only supports \"*.PPTX\".");
            }
        }

        /// <summary>
        /// Determines the output path, based on the second command-line argument, if provided.
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns>The resolved output string.</returns>
        private static string GetOutput(string[] args)
        {
            return (args.Length < 2 || String.IsNullOrWhiteSpace(args[1])) ? null : args[1];
        }

        /// <summary>
        /// Gets the target filename, either from the command line arguments or from the user
        /// </summary>
        /// <param name="args">The command line arguments.</param>
        /// <returns>The accepted filename</returns>
        /// <exception cref="Exception">
        /// File argument was not provided, and the program is running non-interactively!
        /// or
        /// The file doesn't exist!
        /// </exception>
        private static string GetFilename(string[] args)
        {
            string file = null;
            if (args.Length < 1 || String.IsNullOrWhiteSpace(args[0]))
            {
                if (Environment.UserInteractive)
                {
                    Console.Write("Filename: ");
                    file = Console.ReadLine();
                }
                else
                {
                    throw new Exception("File argument was not provided, and the program is running non-interactively!");
                }
            }
            else
            {
                file = args[0];
            }

            if (!File.Exists(file))
            {
                throw new Exception($"File {file} does not exist!");
            }

            return file;
        }

        /// <summary>
        /// Generates the output with the specified formatter.
        /// </summary>
        /// <param name="output">The output path. May be null to represent STDOUT.</param>
        /// <param name="parsedSlides">The parsed slides.</param>
        /// <param name="formatter">The output specifier.</param>
        private static void GenerateOutput(string output, List<Model.Slide> parsedSlides, ISlideFormatter formatter)
        {
            // Scrub empty strings
            parsedSlides.ForEach(slide => {
                slide.Titles.RemoveAll(s => String.IsNullOrWhiteSpace(s));
                slide.SubTitles.RemoveAll(s => String.IsNullOrWhiteSpace(s));
                slide.Bullets.RemoveAll(b => String.IsNullOrWhiteSpace(b.Text));
            });

            bool consoleOut = output == null;
            bool singleFileOutput = false;
            string outputDirectory = null;
            if (!consoleOut)
            {
                singleFileOutput = output.EndsWith(".MD", StringComparison.CurrentCultureIgnoreCase);
                outputDirectory = singleFileOutput ?
                                    output.Substring(0, output.LastIndexOf(Path.DirectorySeparatorChar)) :
                                    output;
                Directory.CreateDirectory(outputDirectory);
                var imgOutputDirectory = outputDirectory + Path.DirectorySeparatorChar + "img";
                Directory.CreateDirectory(imgOutputDirectory);

                // Dump the image files
                foreach (var slide in parsedSlides)
                {
                    for (int i = 0; i < slide.Images.Count; i++)
                    {
                        string fileName = imgOutputDirectory +
                                        Path.DirectorySeparatorChar +
                                        $"{slide.ID.ToString()}-img{i + 1}.jpg";
                        slide.Images[i].Data.Save(fileName);
                    }
                }
            }

            var lastDir = outputDirectory != null ?
                "/" + outputDirectory.Substring(outputDirectory.LastIndexOf(Path.DirectorySeparatorChar) + 1) :
                String.Empty;
            var imgPath = $".{lastDir}/img";
            // One long dump of markdown, put it all in a StringBuilder and vomit
            if (consoleOut || singleFileOutput)
            {
                var markdownString = new StringBuilder();
                parsedSlides.ForEach(slide => markdownString.Append(formatter.Convert(slide, imgPath)));

                if (consoleOut)
                {
                    Console.Write(markdownString.ToString());
                    // Give the user a chance to copy it all out before the shell potentially dies.
                    Console.WriteLine("END OF SLIDESHOW - PRESS ANY KEY TO EXIT");
                    Console.ReadKey();
                }
                else
                {
                    File.WriteAllText(output, markdownString.ToString());
                }
            }
            else
            {
                // Dump each slide individually
                for (int i = 0; i < parsedSlides.Count; i++)
                {
                    File.WriteAllText(output + Path.DirectorySeparatorChar + $"slide{i + 1}.MD", formatter.Convert(parsedSlides[i], imgPath));
                }
            }
        }
    }
}
