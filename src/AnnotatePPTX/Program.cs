using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Serilog;
using Microsoft.Extensions.DependencyInjection;
using System.IO;

namespace AnnotatePPTX
{
    class Program
    {
        private static ILogger _logger;

        static void Main(string[] args)
        {
            var services = Startup();

            var presentationFile = "test-files/slides.pptx";
            var fileCopyPath = "test-files/replacedAudio.pptx";

            File.Copy(presentationFile, fileCopyPath, true);
            _logger = Log.Logger;

            var doc = new Presentation(fileCopyPath, Log.Logger);

            var comments = doc.GetAllSlideNotes();
            doc.ReplaceSlideAudioAnnotation(comments.FirstOrDefault()?.SlideRelationshipId, File.ReadAllBytes("test-files/zh-HK.wav"));

            //// Open the presentation as read-only.
            //using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            //{
            //    ReadSlideInfo(presentationDocument);
            //}

            Console.WriteLine("Hello World!");
        }

        private static IServiceProvider Startup()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.File("runtime.log")
                .WriteTo.Console()
                .CreateLogger();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);

            var serviceProvider = serviceCollection.BuildServiceProvider();
            return serviceProvider;
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(configure => configure.AddSerilog());
                    // .AddTransient<MyClass>();
        }

        // Insert the specified slide into the presentation at the specified position.
        public static void ReadSlideInfo(PresentationDocument presentationDocument)
        {

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
// NotesSlidePart.NotesSlide.InnerXml:
// "<p:cSld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /><a:chOff x=\"0\" y=\"0\" /><a:chExt cx=\"0\" cy=\"0\" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" /></p:nvPr></p:nvSpPr><p:spPr /></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"1\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:rPr lang=\"en-US\" dirty=\"0\" /><a:t>Hey there!</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Number Placeholder 3\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:fld id=\"{8E72DC8C-08B2-4C1F-8F46-48208154C347}\" type=\"slidenum\"><a:rPr lang=\"en-US\" smtClean=\"0\" /><a:t>1</a:t></a:fld><a:endParaRPr lang=\"en-US\" /></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3488005250\" /></p:ext></p:extLst></p:cSld><p:clrMapOvr xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:masterClrMapping xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:clrMapOvr>"
            }
        }


        // // Insert the specified slide into the presentation at the specified position.
        // public static void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        // {

        //     if (presentationDocument == null)
        //     {
        //         throw new ArgumentNullException("presentationDocument");
        //     }

        //     if (slideTitle == null)
        //     {
        //         throw new ArgumentNullException("slideTitle");
        //     }

        //     PresentationPart presentationPart = presentationDocument.PresentationPart;

        //     // Verify that the presentation is not empty.
        //     if (presentationPart == null)
        //     {
        //         throw new InvalidOperationException("The presentation document is empty.");
        //     }

        //     // Declare and instantiate a new slide.
        //     Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
        //     uint drawingObjectId = 1;

        //     // Construct the slide content.
        //     // Specify the non-visual properties of the new slide.
        //     NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
        //     nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
        //     nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
        //     nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        //     // Specify the group shape properties of the new slide.
        //     slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

        //     // Declare and instantiate the title shape of the new slide.
        //     Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

        //     drawingObjectId++;

        //     // Specify the required shape properties for the title shape.
        //     titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
        //         (new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Title" },
        //         new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
        //         new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
        //     titleShape.ShapeProperties = new ShapeProperties();

        //     // Specify the text of the title shape.
        //     titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
        //             new Drawing.ListStyle(),
        //             new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() { Text = slideTitle })));

        //     // Declare and instantiate the body shape of the new slide.
        //     Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
        //     drawingObjectId++;

        //     // Specify the required shape properties for the body shape.
        //     bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(new NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
        //             new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
        //             new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
        //     bodyShape.ShapeProperties = new ShapeProperties();

        //     // Specify the text of the body shape.
        //     bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
        //             new Drawing.ListStyle(),
        //             new Drawing.Paragraph());

        //     // Create the slide part for the new slide.
        //     SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

        //     // Save the new slide part.
        //     slide.Save(slidePart);

        //     // Modify the slide ID list in the presentation part.
        //     // The slide ID list should not be null.
        //     SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

        //     // Find the highest slide ID in the current list.
        //     uint maxSlideId = 1;
        //     SlideId prevSlideId = null;

        //     foreach (SlideId slideId in slideIdList.ChildElements)
        //     {
        //         if (slideId.Id > maxSlideId)
        //         {
        //             maxSlideId = slideId.Id;
        //         }

        //         position--;
        //         if (position == 0)
        //         {
        //             prevSlideId = slideId;
        //         }

        //     }

        //     maxSlideId++;

        //     // Get the ID of the previous slide.
        //     SlidePart lastSlidePart;

        //     if (prevSlideId != null)
        //     {
        //         lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
        //     }
        //     else
        //     {
        //         lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
        //     }

        //     // Use the same slide layout as that of the previous slide.
        //     if (null != lastSlidePart.SlideLayoutPart)
        //     {
        //         slidePart.AddPart(lastSlidePart.SlideLayoutPart);
        //     }

        //     // Insert the new slide into the slide list after the previous slide.
        //     SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
        //     newSlideId.Id = maxSlideId;
        //     newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

        //     // Save the modified presentation.
        //     presentationPart.Presentation.Save();
        // }
    }
}
