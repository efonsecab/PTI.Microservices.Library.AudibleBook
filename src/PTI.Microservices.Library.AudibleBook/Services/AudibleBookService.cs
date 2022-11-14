using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using PTI.Microservices.Library.Interceptors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace PTI.Microservices.Library.Services.Specialized
{
    /// <summary>
    /// Service in charge of converting books to audio
    /// </summary>
    public class AudibleBookService
    {
        private AzureSpeechService AzureSpeechService { get; }
        private ILogger<AudibleBookService> Logger { get; }
        private CustomHttpClient CustomHttpClient { get; }

        /// <summary>
        /// Creates a new instance of <see cref="AudibleBookService"/>
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="azureSpeechService"></param>
        /// <param name="customHttpClient"></param>
        public AudibleBookService(ILogger<AudibleBookService> logger, AzureSpeechService azureSpeechService,
            CustomHttpClient customHttpClient)
        {
            this.AzureSpeechService = azureSpeechService;
            this.Logger = logger;
            this.CustomHttpClient = customHttpClient;
        }

        /// <summary>
        /// Reads the text from the book and generates the audio into the specified output stream.
        /// Produces a maximum of 10 minutes.
        /// For longer books, you will need to split the source contents into several files
        /// </summary>
        /// <param name="outputStream"></param>
        /// <param name="bookUrl"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task ConvertBookToAudioAsync(Stream outputStream,Uri bookUrl, CancellationToken cancellationToken = default)
        {
            try
            {
                List<MemoryStream> lstOutputStreams = new List<MemoryStream>();
                var bookBytes = await this.CustomHttpClient.GetByteArrayAsync(bookUrl, cancellationToken);
                MemoryStream memoryStream = new MemoryStream(bookBytes);
                // Open a WordprocessingDocument for read-only access based on a stream.
                List<string> elementTypes = new List<string>();
                List<string> encodedInputs = new List<string>();
                StringBuilder stringBuilder = new StringBuilder();
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(memoryStream, isEditable: true))
                {
                    MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var content = mainPart.Document.Body.InnerText;
                    var allElements = mainPart.Document.Body.ToList();
                    var allTypes = allElements.Select(p => p.GetType().Name).Distinct().ToList();
                    foreach (var singleElement in allElements)
                    {
                        if (singleElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                        {
                            var paragraph = singleElement as DocumentFormat.OpenXml.Wordprocessing.Paragraph;
                            var paragraphInnerText = paragraph.InnerText;
                            stringBuilder.AppendLine(paragraphInnerText?.Replace(".", " "));
                        }
                        else
                            if (singleElement is DocumentFormat.OpenXml.Wordprocessing.SdtBlock)
                        {
                        }
                        else
                            if (singleElement is DocumentFormat.OpenXml.Wordprocessing.SectionProperties)
                        {

                        }
                        else
                            elementTypes.Add(singleElement.GetType().FullName);
                    };
                }
            }
            catch (Exception ex)
            {
                this.Logger?.LogError(ex.Message, ex);
                throw;
            }
        }

        /// <summary>
        /// Generates audio for the book in the specified url.
        /// Produce a memory stream every 20 lines of text recognized.
        /// </summary>
        /// <param name="bookUrl"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public async Task<List<MemoryStream>> ConvertBookToAudioSplittedAsync(Uri bookUrl, CancellationToken cancellationToken = default)
        {
            try
            {
                List<MemoryStream> lstOutputStreams = new List<MemoryStream>();
                var bookBytes = await this.CustomHttpClient.GetByteArrayAsync(bookUrl, cancellationToken);
                MemoryStream memoryStream = new MemoryStream(bookBytes);
                // Open a WordprocessingDocument for read-only access based on a stream.
                List<string> elementTypes = new List<string>();
                List<string> encodedInputs = new List<string>();
                StringBuilder stringBuilder = new StringBuilder();
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(memoryStream, isEditable: true))
                {
                    MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                    var content = mainPart.Document.Body.InnerText;
                    var allElements = mainPart.Document.Body.ToList();
                    var allTypes = allElements.Select(p => p.GetType().Name).Distinct().ToList();
                    int linesInBuffer = 0;
                    foreach (var singleElement in allElements)
                    {
                        if (singleElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                        {
                            var paragraph = singleElement as DocumentFormat.OpenXml.Wordprocessing.Paragraph;
                            //var allTextForParagraph =
                            //                paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
                            //foreach (var singleText in allTextForParagraph)
                            //{
                            //    stringBuilder.AppendLine(singleText.Text);
                            //}
                            //----------------
                            var paragraphInnerText = paragraph.InnerText;
                            stringBuilder.AppendLine(paragraphInnerText?.Replace(".", " "));
                            linesInBuffer++;
                            if (linesInBuffer >= 20)
                            {
                                MemoryStream pageOutputStream = new MemoryStream();
                                await this.AzureSpeechService.TalkToStreamAsync(stringBuilder.ToString(), pageOutputStream);
                                lstOutputStreams.Add(pageOutputStream);
                                stringBuilder.Clear();
                                linesInBuffer = 0;
                            }
                        }
                        else
                            if (singleElement is DocumentFormat.OpenXml.Wordprocessing.SdtBlock)
                        {
                        }
                        else
                            if (singleElement is DocumentFormat.OpenXml.Wordprocessing.SectionProperties)
                        {

                        }
                        else
                            elementTypes.Add(singleElement.GetType().FullName);
                    };
                    if (linesInBuffer > 0)
                    {
                        MemoryStream pageOutputStream = new MemoryStream();
                        await this.AzureSpeechService.TalkToStreamAsync(stringBuilder.ToString(), pageOutputStream);
                        lstOutputStreams.Add(pageOutputStream);
                        stringBuilder.Clear();
                    }
                }
                //await this.AzureSpeechService.TalkToStreamAsync(stringBuilder.ToString(), outputStream);
                return lstOutputStreams;
            }
            catch (Exception ex)
            {
                this.Logger?.LogError(ex.Message, ex);
                throw;
            }
        }
    }
}
