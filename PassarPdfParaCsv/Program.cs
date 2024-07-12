using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using PassarPdfParaCsv.Entities;

namespace PassarPdfParaCsv
{

    /// <summary>
    ///   O código está fase de teste, vai ser melhorado ao longo do tempo. (Futuramente podendo criar um front para ele) 
    ///   Ele lê PDFs de três colunas, é preciso ajustar as colunas conforme o seu tamanho e espaçamento
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Informe o caminho do pasta, por exemplo: C:\\Users\\devTeste\\Downloads\\Compass2025.pdf  \n");
                string pathOrigin = Console.ReadLine();

                ReadPdf(pathOrigin);
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Error ao gerar arquivo: {ex}");
            }
        }

        private static void ReadPdf(string PathOrigin)
        {
            using (PdfReader pdfReader = new PdfReader(PathOrigin))
            {
                using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
                {
                    Dictionary<string, string> dataObtained = new Dictionary<string, string>();
                    int numPaginas = pdfDocument.GetNumberOfPages();

                    string fullTextpage = "";

                    for (int page = 1; page < numPaginas; page++)
                    {
                        PdfPage pdfPage = pdfDocument.GetPage(page);
                        ImageRenderData tamanhoImagem = new ImageRenderData();
                        PdfCanvasProcessor processor = new PdfCanvasProcessor(tamanhoImagem);
                        processor.ProcessPageContent(pdfPage);

                        // Obtenha as dimensões da página
                        Rectangle pageSize = pdfPage.GetPageSize();
                        float pageWidth = pageSize.GetWidth();
                        float pageHeight = pageSize.GetHeight();


                        if (pageWidth == 510.235657f && pageHeight == 354.329224f)
                        {
                            // Divida a largura da página em três partes iguais
                            float columnWidth = pageWidth / 2.9f;

                            // Defina as regiões das colunas
                            Rectangle leftColumn = new Rectangle(0, 0, columnWidth, pageHeight);
                            Rectangle middleColumn = new Rectangle(columnWidth * 1.0f, 0, columnWidth * 0.86f, pageHeight);
                            Rectangle rightColumn = new Rectangle(columnWidth * 1.9f, 0, columnWidth * 0.83f, pageHeight);

                            // Extraia texto de cada coluna
                            string leftColumnText = ExtractTextFromRegion(pdfPage, leftColumn).Replace("\r", "").Replace("\n", " ");
                            string middleColumnText = ExtractTextFromRegion(pdfPage, middleColumn).Replace("\r", "").Replace("\n", " ");
                            string rightColumnText = ExtractTextFromRegion(pdfPage, rightColumn).Replace("\r", "").Replace("\n", " ");

                            fullTextpage = $"{leftColumnText}\n\n{middleColumnText}\n\n{rightColumnText}";

                        }
                        else
                        {
                            fullTextpage = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(page));
                        }
                        dataObtained.Add(page.ToString(), fullTextpage);
                    }

                    Random numberRandom = new Random();
                    int numberFile = numberRandom.Next(01, 100);

                    string fileName = $"PdfParaCsv{numberFile}.csv";
                    string directoryPath = System.IO.Path.GetDirectoryName(PathOrigin);
                    string folderPath = $"{directoryPath}\\{fileName}";

                    ConvertCsv(dataObtained, folderPath);
                }
            }
        }

        static string ExtractTextFromRegion(PdfPage pdfPage, Rectangle region)
        {
            ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), new TextRegionEventFilter(region));
            return PdfTextExtractor.GetTextFromPage(pdfPage, strategy);
        }

        private static void ConvertCsv(Dictionary<string, string> dados, string folderPath)
        {
            using (var writer = new StreamWriter(folderPath))
            {
                using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    csv.WriteField("Conteúdo");
                    csv.WriteField("Página");
                    csv.NextRecord();

                    foreach (var kvp in dados)
                    {
                        csv.WriteField(kvp.Value);
                        csv.WriteField(kvp.Key);
                        csv.NextRecord();
                    }
                }
            }

            Console.WriteLine("\nConcluído");
            Console.WriteLine($"\nCaminho do arquivo gerado: {folderPath}");
        }
    }

}

