﻿using CsvHelper.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using CsvHelper;

namespace PowerPointGenerator
{
    public class Generator
    {

        private const string PlaceholderPattern = @"##(?'key'\S*)##";
        private const string PatternGroupKey = "key";
        private readonly string tmpDestination;
        private readonly List<Dictionary<string, string>> data;
        private readonly PresentationDocument document;

        public Generator(string sourceFile, string sourceDataFile)
        {
            this.tmpDestination = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pptx";
            File.Copy(sourceFile, tmpDestination);
            this.document = PresentationDocument.Open(tmpDestination, true);
            this.data = ReadCsvData(sourceDataFile);
        }

        private List<Dictionary<string, string>> ReadCsvData(string sourceDataFile)
        {
            var data = new List<Dictionary<string, string>>();
            using (var reader = new StreamReader(sourceDataFile))
            {
                var parser = new CsvParser(reader, new CsvConfiguration(CultureInfo.CurrentCulture));
                var headerArray = parser.Read();
                while (true)
                {
                    var dataLine = parser.Read();
                    if (dataLine == null)
                    {
                        break;
                    }
                    var dataDict = new Dictionary<string, string>();
                    for (int i = 0; i < headerArray.Length; i++)
                    {
                        var key = headerArray[i].Replace(" ", "").ToUpper();
                        if (dataDict.ContainsKey(key))
                        {
                            continue;
                        }
                        dataDict.Add(key, dataLine[i]);
                    }
                    data.Add(dataDict);
                }
            }
            return data;
        }

        public void Convert()
        {
            SlidePart templateSlide = GetTemplateSlide();
            foreach (var line in data)
            {
                CreateAndFillDataSlide(templateSlide, line);
            }
            SaveAndCloseDocument();
        }

        public void SaveInDestination(string outputPath)
        {
            File.Copy(tmpDestination, outputPath, true);
        }

        private void SaveAndCloseDocument()
        {
            document.Save();
            document.Close();
        }

        private void CreateAndFillDataSlide(SlidePart templateSlide, Dictionary<string, string> line)
        {
            SlidePart ticketSlide = CloneSlidePart(templateSlide);
            List<OpenXmlElement> placeholderElements = GetChildShapesByPlaceholderpattern(ticketSlide.Slide.ChildElements);
            foreach (OpenXmlElement placeholderElement in placeholderElements)
            {
                try
                {
                    FillPlaceholder((P.Shape)placeholderElement, line);
                }
                catch
                {
                    //TODO: Handle this
                }
            }

            InsertSlideInPresentation(ticketSlide);
        }

        private SlidePart GetTemplateSlide()
        {
            var slideParts = document.PresentationPart.SlideParts.ToList();
            var templateTicketSlide = slideParts.First(s => GetChildShapesByPlaceholderpattern(s.Slide.ChildElements) != null);
            return templateTicketSlide;
        }

        private void InsertSlideInPresentation(SlidePart ticketSlide)
        {
            SlideId slideIdLast = document.PresentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>().OrderByDescending(sid => sid.Id).First();
            SlideId newSlideId = document.PresentationPart.Presentation.SlideIdList.InsertAfter(new SlideId(), slideIdLast);
            newSlideId.Id = slideIdLast.Id + 1;
            newSlideId.RelationshipId = document.PresentationPart.GetIdOfPart(ticketSlide);
        }

        private void FillPlaceholder(P.Shape placeholderElement, Dictionary<string, string> dataRow)
        {
            Match match = Regex.Match(placeholderElement.InnerText, PlaceholderPattern);
            if (!match.Success)
            {
                throw new InvalidDataException("placeholderElement.InnerText has to match the PlaceholderPattern");
            }

            string fieldName = match.Groups[PatternGroupKey].Value;
            string value = dataRow[fieldName];

            placeholderElement.TextBody = new P.TextBody(
                                new BodyProperties(),
                                new ListStyle(),
                                new Paragraph(
                                    new Run() { Text = new D.Text(value) },
                                    new EndParagraphRunProperties() { Language = "en-US" }));

        }

        private SlidePart CloneSlidePart(SlidePart ticketSlide)
        {
            Slide currSlide = (Slide)ticketSlide.Slide.CloneNode(true);
            SlidePart sp = document.PresentationPart.AddNewPart<SlidePart>();
            currSlide.Save(sp);
            using (var stream = ticketSlide.GetStream())
            {
                sp.FeedData(stream);
                sp.AddPart(ticketSlide.SlideLayoutPart);
            }

            return sp;
        }

        private List<OpenXmlElement> GetChildShapesByPlaceholderpattern(OpenXmlElementList list)
        {
            List<OpenXmlElement> found = new List<OpenXmlElement>();
            foreach (var element in list)
            {
                found.AddRange(FindShapesIncludingElement(element));
            }
            return found;
        }

        private List<OpenXmlElement> FindShapesIncludingElement(OpenXmlElement element)
        {
            bool elementMatchesPattern = !string.IsNullOrEmpty(element.InnerText) && Regex.IsMatch(element.InnerText, PlaceholderPattern);
            if (elementMatchesPattern && element is P.Shape)
            {
                return new List<OpenXmlElement> { element };
            }
            var foundChilds = GetChildShapesByPlaceholderpattern(element.ChildElements);
            return foundChilds;
        }
    }
}
