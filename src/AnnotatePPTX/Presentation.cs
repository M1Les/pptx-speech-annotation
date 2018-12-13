using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace AnnotatePPTX
{
    public class Presentation
    {
        private readonly string filePath;

        public Presentation(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Argument cannot be empty", nameof(filePath));
            }

            this.filePath = filePath;
        }

        public IReadOnlyCollection<PresentationSlideComment> GetSlideComments()
        {
            var comments = new List<PresentationSlideComment>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(this.filePath, false))
            {
                SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    // NotesSlidePart.NotesSlide.InnerXml:
                    // "<p:cSld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /><a:chOff x=\"0\" y=\"0\" /><a:chExt cx=\"0\" cy=\"0\" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" /></p:nvPr></p:nvSpPr><p:spPr /></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"1\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:rPr lang=\"en-US\" dirty=\"0\" /><a:t>Hey there!</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Number Placeholder 3\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:fld id=\"{8E72DC8C-08B2-4C1F-8F46-48208154C347}\" type=\"slidenum\"><a:rPr lang=\"en-US\" smtClean=\"0\" /><a:t>1</a:t></a:fld><a:endParaRPr lang=\"en-US\" /></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3488005250\" /></p:ext></p:extLst></p:cSld><p:clrMapOvr xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:masterClrMapping xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:clrMapOvr>"
                }
            }

            return comments.AsReadOnly();
        }
    }
}
