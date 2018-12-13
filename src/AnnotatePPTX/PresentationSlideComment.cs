using System;
using System.Collections.Generic;
using System.Text;

namespace AnnotatePPTX
{
    public class PresentationSlideComment
    {
        private readonly string slidePartId;
        private readonly string text;

        public PresentationSlideComment(string slidePartId, string text)
        {
            this.slidePartId = slidePartId;
            this.text = text;
        }
    }
}
