using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using Serilog;
using CSCore.MediaFoundation;
using CSCore;

namespace AnnotatePPTX
{
    public class Presentation
    {
        private readonly string _filePath;
        private readonly ILogger _logger;
        private readonly AACEncoder _audioEncoder;

        public Presentation(string filePath, ILogger logger, AACEncoder audioEncoder)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Argument cannot be empty", nameof(filePath));
            }

            this._filePath = filePath;
            this._logger = logger;
            this._audioEncoder = audioEncoder;
        }

        public IReadOnlyCollection<PresentationSlideNote> GetAllSlideNotes()
        {
            var comments = new List<PresentationSlideNote>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(this._filePath, false))
            {
                var presentationPart = presentationDocument.PresentationPart;
                SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
                uint slideShowIndex = 0;

                _logger.Verbose($"Getting notes from {slideIdList.Count()} slides");

                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    slideShowIndex++;
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    var notesText = slidePart.NotesSlidePart?.NotesSlide?.InnerText;

                    if (!string.IsNullOrWhiteSpace(notesText))
                    {
                        comments.Add(new PresentationSlideNote(slideId.Id?.Value, slideId.RelationshipId, notesText, slideShowIndex));
                    }
                    //slidePart.Parts.Select(rel => rel.)
                    // NotesSlidePart.NotesSlide.InnerXml:
                    // "<p:cSld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /><a:chOff x=\"0\" y=\"0\" /><a:chExt cx=\"0\" cy=\"0\" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" /></p:nvPr></p:nvSpPr><p:spPr /></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"1\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:r><a:rPr lang=\"en-US\" dirty=\"0\" /><a:t>Hey there!</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Number Placeholder 3\" /><p:cNvSpPr><a:spLocks noGrp=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:cNvSpPr><p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"5\" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:lstStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /><a:p xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:fld id=\"{8E72DC8C-08B2-4C1F-8F46-48208154C347}\" type=\"slidenum\"><a:rPr lang=\"en-US\" smtClean=\"0\" /><a:t>1</a:t></a:fld><a:endParaRPr lang=\"en-US\" /></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3488005250\" /></p:ext></p:extLst></p:cSld><p:clrMapOvr xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><a:masterClrMapping xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:clrMapOvr>"
                }
            }

            return comments.AsReadOnly();
        }

        public void ReplaceSlideAudioAnnotation(string slideRelationshipId, byte[] audioFile)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(this._filePath, true))
            {
                ReplaceSlideAudioAnnotationForRelationship(presentationDocument, slideRelationshipId, audioFile);
            }
        }

        public void ReplaceAllSlideAudioAnnotations(Dictionary<string, string> slidesToAnnotations)
        {
            _logger.Verbose($"Replacing slide audio annotations in batch mode for '{this._filePath}'.");
            _logger.Verbose($"Got {slidesToAnnotations.Count} annotations for replacement.");
            using (PresentationDocument presentationDocument = PresentationDocument.Open(this._filePath, true))
            {
                foreach (var item in slidesToAnnotations)
                {
                    if (string.IsNullOrWhiteSpace(item.Value))
                    {
                        _logger.Debug($"There was no translated annotation for relationship id: '{item.Key}'. Skipping.");
                        continue;
                    }

                    _logger.Verbose($"Replacing audio annotation. Relationship id: '{item.Key}', replacement file: '{item.Value}'");
                    string slideRelationshipId = item.Key;
                    var audioFile = File.ReadAllBytes(item.Value);

                    // short files are most likely corrupt
                    if (audioFile.LongLength < 100)
                    {
                        _logger.Warning($"Empty or corrupt wav file: '{item.Value}'.");
                    }
                    else
                    {
                        _logger.Verbose($"Replacement file size: '{audioFile.LongLength}'");
                        ReplaceSlideAudioAnnotationForRelationship(presentationDocument, slideRelationshipId, audioFile);
                    }
                }
            }
        }

        private void ReplaceSlideAudioAnnotationForRelationship(PresentationDocument presentationDocument, string slideRelationshipId, byte[] audioFile)
        {
            var presentationPart = presentationDocument.PresentationPart;
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideRelationshipId);
            MediaDataPart mediaPart = (MediaDataPart)slidePart.DataPartReferenceRelationships
                .FirstOrDefault(dpr => (dpr.DataPart?.ContentType == "audio/x-wav" || dpr.DataPart?.ContentType == "audio/mp4")
                                        && dpr.RelationshipType.EndsWith("/relationships/media"))?.DataPart;

            if (mediaPart != null)
            {
                this._logger.Verbose($"Replacing contents of '{mediaPart.Uri}' media part.");

                byte[] encodedAAC = new byte[0];

                if (mediaPart.ContentType == "audio/mp4")
                {
                    this._logger.Verbose("Converting replacement audio from Wav to Mp4 AAC.");

                    encodedAAC = _audioEncoder.FromWav(audioFile);
                }

                using (var mpStream = mediaPart.GetStream())
                using (var writer = new BinaryWriter(mpStream))
                {
                    if (encodedAAC.LongLength > 0)
                    {
                        writer.Write(encodedAAC);
                    }
                    else
                    {
                        writer.Write(audioFile);
                    }
                }
            }
            else
            {
                this._logger.Warning("Current slide has no media in it.");
                this._logger.Verbose("Could not find any parts with '/relationships/media' type and 'audio/x-wav' or 'audio/mp4' content types.");
            }
        }
    }
}
