﻿using System;
using System.Collections.Generic;
using System.Text;

namespace AnnotatePPTX
{
    public class PresentationSlideNote
    {
        public PresentationSlideNote(uint? slidePartId, string slideRelationshipId, string text)
        {
            this.SlidePartId = slidePartId ?? throw new ArgumentNullException(nameof(slidePartId));
            this.SlideRelationshipId = slideRelationshipId ?? throw new ArgumentNullException(nameof(slideRelationshipId));
            this.Text = text ?? throw new ArgumentNullException(nameof(text));
        }

        public uint SlidePartId { get; }

        public string SlideRelationshipId { get; }

        public string Text { get; }
    }
}
