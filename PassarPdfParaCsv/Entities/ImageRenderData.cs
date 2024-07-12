using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PassarPdfParaCsv.Entities
{
    public class ImageRenderData : IEventListener
    {
        public List<ImageData> imageData { get; private set; } = new List<ImageData>();
        public void EventOccurred(IEventData data, EventType type)
        {
            if (type == EventType.RENDER_IMAGE)
            {
                var renderInfo = (ImageRenderInfo)data;
                var image = renderInfo.GetImage();

                var dadosImage = new ImageData
                {
                    Width = image.GetWidth(),
                    Height = image.GetHeight()
                };

                imageData.Add(dadosImage);
            }
        }
        public ICollection<EventType> GetSupportedEvents()
        {
            return new HashSet<EventType> { EventType.RENDER_IMAGE };
        }
    }
}
