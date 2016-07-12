using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RSGenerate
{
    public static class XMLHelper
    {
        public static string GetAttributeValue(XElement element, string name)
        {
            var attr = GetAttribute(element, name);

            if (attr == null)
                return string.Empty;

            return attr.Value.ToString();
        }

        public static XAttribute GetAttribute(XElement element, string name)
        {
            return element.Attributes().Where(a => a.Name == name).FirstOrDefault();
        }

        public static XElement GetChildElement(XElement parent, string name)
        {
            return parent.Nodes().OfType<XElement>().Where(n => n.Name == name).FirstOrDefault();
        }

        public static XElement CreateRoutineElement(string name)
        {
            return new XElement("Routine", new XAttribute("Name", name), new XAttribute("Type", "RLL"), 
                        new XElement("RLLContent"));
        }

        public static XElement GetOrCreateRoutinesElement(XElement root)
        {
            XElement routines = null;

            var programs = XMLHelper.GetChildElement((XElement)root.FirstNode, "Programs");
            var mainProgram = programs.Nodes().OfType<XElement>()
                                .Where(e => e.Attributes().Any(a => a.Name == "Name" && a.Value == "MainProgram"))
                                .FirstOrDefault();

            if (mainProgram == null)
                throw new ArgumentNullException("Template must have a Program Element with an attr of MainProgram");

            routines = XMLHelper.GetChildElement(mainProgram, "Routines");
            if (routines == null)
            {
                routines = new XElement("Routines");
                mainProgram.Add(routines);
            }

            return routines;
        }

        public static XElement GetOrCreateRoutineWithName(XElement routinesElement, string name)
        {
            var routine = routinesElement.Elements()
                                .Where(e => e.Attributes().Any(a => a.Name == "Name" && a.Value == name))
                                .FirstOrDefault();

            if (routine == null)
            {
                routine = XMLHelper.CreateRoutineElement(name);
                routinesElement.Add(routine);
            }

            return routine;
        }

        public static XElement CreateTag(string tagName, string dataType, string tagType = "Base", string externalAccess = "Read/Write", bool? constant = false)
        {
            
            return new XElement("Tag",
                        new XAttribute("Constant", constant == null ? "False" : constant.ToString()),
                        new XAttribute("DataType", dataType),
                        new XAttribute("ExternalAccess", externalAccess),
                        new XAttribute("Name", tagName),
                        new XAttribute("TagType", tagType)
                    );
        }

        public static XElement CreateAliasTag(string tagName, string alias, string externalAccess = "Read/Write")
        {

            return new XElement("Tag",
                        new XAttribute("AliasFor", alias),
                        new XAttribute("ExternalAccess", externalAccess),
                        new XAttribute("Name", tagName),
                        new XAttribute("Radix", "Binary"),
                        new XAttribute("TagType", "Alias")
                    );
        }

        public static XElement CreateRung(string rungText, string rungType = "N", string rungNumber = "0")
        {
            var rung = new XElement("Rung",
                        new XAttribute("Number", rungNumber),
                        new XAttribute("Type", rungType));
            
            var textElement = new XElement("Text");
            textElement.Add(new XCData(rungText));
            rung.Add(textElement);

            return rung;
        }

        public static void AddSourceRungsToTargetRoutine(IEnumerable<XElement> sourceRungs, XElement routinesElement, string targetRoutineName)
        {
            var targetRoutine = XMLHelper.GetOrCreateRoutineWithName(routinesElement, targetRoutineName);
            //Add Source Ladder Logic to Target Routine
            var targetRungs = targetRoutine.Descendants("Rung");
            if (targetRungs != null && targetRungs.Count() > 0)
                targetRungs.LastOrDefault().AddAfterSelf(sourceRungs); //add to end
            else
                targetRoutine.Descendants("RLLContent").LastOrDefault().Add(sourceRungs);
        }
    }
}
