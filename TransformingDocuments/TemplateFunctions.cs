using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Linq;
using System.IO;

namespace TransformingDocuments
{
    public static partial class ThemeFunctions
    {
        #region overload_string_parms
        // Apply a new theme to the presentation. 
        public static void ApplyThemeToPresentation(string presentationFile, string themePresentation)
        {
            using (PresentationDocument themeDocument = PresentationDocument.Open(themePresentation, false))
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                ApplyThemeToPresentation(presentationDocument, themeDocument);
            }
        }
        #endregion

        // Apply a new theme to the presentation. 
        public static void ApplyThemeToPresentation(PresentationDocument presentationDocument, PresentationDocument themeDocument)
        {
            #region parametervalidation

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }
            if (themeDocument == null)
            {
                throw new ArgumentNullException("themeDocument");
            }

            #endregion

            // Get the presentation part of the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // 
            // Step 2.
            // Remove the exising master/theme parts from the uploaded presentation.
            //
            // Get the existing slide master part.
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.ElementAt(0);
            string relationshipId = presentationPart.GetIdOfPart(slideMasterPart);

            // Remove the existing theme part.
            // Remove the old slide master part.
            presentationPart.DeletePart(presentationPart.ThemePart);
            presentationPart.DeletePart(slideMasterPart);

            //
            // Step 3.
            // Get the new slide master part from the template.
            // Import the new slide master part, and reuse the old relationship ID.
            // Change to the new theme part.
            //
            SlideMasterPart newSlideMasterPart = themeDocument.PresentationPart.SlideMasterParts.ElementAt(0);
            newSlideMasterPart = presentationPart.AddPart(newSlideMasterPart, relationshipId);
            presentationPart.AddPart(newSlideMasterPart.ThemePart);

            // 
            // Step 4.
            // Get the collection of layout parts to pair with slides for formatting
            //
            Dictionary<string, SlideLayoutPart> newSlideLayouts = new Dictionary<string, SlideLayoutPart>();
            foreach (var slideLayoutPart in newSlideMasterPart.SlideLayoutParts)
            {
                newSlideLayouts.Add(GetSlideLayoutType(slideLayoutPart), slideLayoutPart);
            }

            string layoutType = null;
            SlideLayoutPart newLayoutPart = null;

            //
            // Step 5. 
            // Identify the names of the Title, default and closing layout parts.
            //
            //      Default layout slide
            //string defaultLayoutType = "標題及物件";
            string defaultLayoutType = "标题和内容";
            //      Title slide
            //string defaultTitleLayoutType = "標題投影片";
            string defaultTitleLayoutType = "标题幻灯片";
            //      Closing or "Thank you" slide.
            //string defaultClosingLayoutType = "空白";
            string defaultClosingLayoutType = "自定义版式";

            //
            // Step 6. 
            // Find the First and last slides and hold onto their part objects
            // 
            // Get first slide's relationship id
            var firstSlideId = ((SlideId)presentationPart.
                RootElement.
                Descendants().
                First<OpenXmlElement>(e => e.LocalName == "sldIdLst").
                Descendants().
                ElementAt<OpenXmlElement>(0)).RelationshipId;

            // Get slide count.
            #region getslidecount

            var slideIdCount = presentationPart.
                RootElement.
                Descendants().
                First<OpenXmlElement>(e => e.LocalName == "sldIdLst").
                Descendants().
                Count();

            #endregion

            // Get the last slide's relationship id
            var lastSlideId = ((SlideId)presentationPart.
                RootElement.
                Descendants().
                First<OpenXmlElement>(e => e.LocalName == "sldIdLst").
                Descendants().
                ElementAt<OpenXmlElement>(slideIdCount - 1)).RelationshipId;

            // Get the first and last slide parts (slide<x>.xml)
            var firstSlidePart = presentationPart.GetPartById(firstSlideId);
            var lastSlidePart = presentationPart.GetPartById(lastSlideId);

            // 
            // Step 7.
            // Run through all the slides (parts) doing the following:
            //      - remove the related layout part (formatting)
            //      - if first slide, then relate to the "Title" layout part
            //      - if last slide, then relate to the "Closing" or "Thank you" layout part
            //      - otherwise, if the currently related layout part's name exists in the new theme,
            //          relate to it. If not, then use the "default" layout part.
            // 
            foreach (var slidePart in presentationPart.SlideParts)
            {
                layoutType = null;

                // - Remove the slide layout relationship on all slides. 
                // If the slide has a layout already, delete it.
                if (slidePart.SlideLayoutPart != null)
                {
                    // Determine the slide layout type for each slide based on it's name.
                    layoutType = GetSlideLayoutType(slidePart.SlideLayoutPart);

                    // Delete the old layout part.
                    slidePart.DeletePart(slidePart.SlideLayoutPart);
                }

                // - If this is the first slide in the deck, set it's layout to the default Title slide layout
                if (firstSlidePart == slidePart)
                {
                    newLayoutPart = newSlideLayouts[defaultTitleLayoutType];

                    // Apply the new default layout part.
                    slidePart.AddPart(newLayoutPart);
                    continue;
                }

                // - If this is the last (i.e. closing) slide, set the layout to the default Closing layout.
                if (lastSlidePart == slidePart)
                {
                    newLayoutPart = newSlideLayouts[defaultClosingLayoutType];

                    // Apply the new default layout part.
                    slidePart.AddPart(newLayoutPart);
                    continue;
                }

                // - Here we check to see if the slide's layout name exists in the template.
                //      If so, then we apply that named layout from the new template.
                //      If not, then we use the default layout.
                if (layoutType != null && newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
                {
                    // Apply the new layout part.
                    slidePart.AddPart(newLayoutPart);
                }
                else
                {
                    newLayoutPart = newSlideLayouts[defaultLayoutType];

                    // Apply the new default layout part.
                    slidePart.AddPart(newLayoutPart);
                }
            }
        }

        // Get the slide layout type.
        public static string GetSlideLayoutType(SlideLayoutPart slideLayoutPart)
        {
            CommonSlideData slideData = slideLayoutPart.SlideLayout.CommonSlideData;

            // Remarks: If this is used in production code, check for a null reference.

            return slideData.Name;
        }
        public static string CombineUrl(string path1, string path2)
        {

            return path1.TrimEnd('/') + '/' + path2.TrimStart('/');
        }

        public static Stream ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return new MemoryStream(ms.ToArray());
            }
        }

        private static void CopyStream(Stream source, Stream destination, long length)
        {
            // byte[] buffer = new byte[32768];
            byte[] buffer = new byte[length];
            int bytesRead;
            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }
    }
}
