using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RSGenerate
{
    public class L5XGenerator
    {
        public XDocument TemplateDoc { get; private set; }
        public IExcelDataReader SystemDefinition { get; private set; }
        public XDocument Output { get; private set; }

        private string _templatePath;

        public void LoadTemplate(string filePath)
        {
            TemplateDoc = XDocument.Load(filePath);
            _templatePath = System.IO.Path.GetDirectoryName(filePath);
        }

        public void LoadDefinition(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var fileInfo = new FileInfo(filePath);

            if (fileInfo.Extension == ".xlsx")
                SystemDefinition = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        public string GenerateAndSaveOutput()
        {
            if (TemplateDoc == null)
                throw new ArgumentNullException("You must first specify a TemplateDoc prior to calling Generate()");

            if (SystemDefinition == null)
                throw new ArgumentNullException("You must first specify a SystemDefinition prior to calling Generate()");

            var reader = SystemDefinition;
            reader.IsFirstRowAsColumnNames = true;
            DataSet dataset = reader.AsDataSet();
            string sheetName = "Sheet1";

            var routines = XMLHelper.GetOrCreateRoutinesElement((XElement)this.TemplateDoc.FirstNode);

            foreach (DataRow row in dataset.Tables[sheetName].Rows)
                this.ProcessRow(row, routines);

            //Remove the Template Program from the Target before saving
            this.TemplateDoc.Descendants("Program").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == "Templates")).FirstOrDefault().Remove();

            //once all the routines are created and/or added to renumber them all sequentially
            this.ReNumberLadder(routines);

            var fileName = _templatePath + "\\GeneratedProject.L5X";
            this.TemplateDoc.Save(fileName);

            return fileName;
        }
        private string GetCellValue(DataRow row, string col)
        {
            return row[col].ToString().Replace('-','_');
        }

        private void ProcessRow(DataRow row, XElement routines)
        {
            //read config data from row
            var conveyorNumber = GetCellValue(row, "Conveyor Number");
            var inboundConnection = GetCellValue(row, "IB Connection");
            var outboundConnection = GetCellValue(row, "OB Connection");
            var targetRoutineName = GetCellValue(row, "Target Routine Name");
            var motorTemplateRoutineName = GetCellValue(row, "Motor Control Template Routine");
            string deviceName = null;

            if (string.IsNullOrEmpty(conveyorNumber))
                return;

            conveyorNumber = "CONVEYOR_" + conveyorNumber;
            this.CreateControllerTag(conveyorNumber, "FXG_CONVEYOR");

            //Import Motor Control Code
            //this.ImportSourceLadderRungs(routines, motorTemplateRoutineName, targetRoutineName, conveyorNumber, inboundConnection, outboundConnection, null);

            //Import Disconnect Fault Code
            this.ImportSourceLadderRungs(routines, "FXG_MOTOR_DISC", "MOTOR_DISC", "CONVEYOR_XXX", conveyorNumber, inboundConnection, outboundConnection, null);

            //Import Motor Overload Code
            this.ImportSourceLadderRungs(routines, "FXG_MOTOR_OVLD", "MOTOR_OVLD", "CONVEYOR_XXX", conveyorNumber, inboundConnection, outboundConnection, null);

            //Import SE Logic
            deviceName =  GetCellValue(row, "SE1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SE");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SE", "CS_SE", "CS_SE1_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "SE2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SE");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SE", "CS_SE", "CS_SE2_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            ////Import SM Logic
            deviceName = GetCellValue(row, "SM1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SM");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SM", "CS_SM", "CS_SM_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "SM2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SM");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SM", "CS_SM", "CS_SM_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "SS1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SS");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SS", "CS_SS", "CS_SS_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "SS2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_SS");
                this.ImportSourceLadderRungs(routines, "FXG_CS_SS", "CS_SS", "CS_SS_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            ////Import S1 Logic
            //if (!string.IsNullOrEmpty(GetCellValue(row, "S1")))
            //    this.ImportSourceLadderRungs(routines, "FXG_CS_S1", "CS_S1", conveyorNumber, inboundConnection, outboundConnection, GetCellValue(row, "S1"));

            deviceName = GetCellValue(row, "EPC1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routines, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "EPC2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routines, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

            deviceName = GetCellValue(row, "EPC3");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.CreateControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routines, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName, inboundConnection, outboundConnection, deviceName);
            }

        }

        private void CreateControllerTag(string tagName, string dataType)
        {
            var tag = new XElement("Tag",
                                    new XAttribute("Constant", "False"),
                                    new XAttribute("DataType", dataType),
                                    new XAttribute("ExternalAccess", "Read/Write"),
                                    new XAttribute("Name", tagName),
                                    new XAttribute("TagType", "Base")
                                );
            this.TemplateDoc.Descendants("Controller").FirstOrDefault().Element("Tags").Add(tag);
        }

        private void ImportSourceLadderRungs(XElement routines, string sourceRountineName, string targetRoutineName, string sourceTag, string currentTag, string inboundConnection, string outboundConnection, string deviceName)
        {
            //Load Ladder Logic from Source Template
            var sourceProgramName = "Templates";

            //first we clone the source ladder rungs
            var rungs = this.GetSourceRungs(sourceProgramName, sourceRountineName);
            if (rungs == null)
                return;

            var sourceClone = new XElement("Source", rungs);
            var sourceRungs = sourceClone.Elements("Rung");

            //Perform Tag Replacement in ladder text
            foreach (var rung in sourceRungs)
            {
                this.ReplaceTagMembers(rung, sourceTag, currentTag);               
                this.ReplaceTagMembers(rung, "IB_CONNECTION", inboundConnection);
                this.ReplaceTagMembers(rung, "OB_CONNECTION", outboundConnection);

                this.ReplaceCommentText(rung, "{DEVICE_NUMBER}", currentTag);
            }

            //Find the Right PLC Routine in Target
            var targetRoutine = XMLHelper.GetOrCreateRoutineWithName(routines, targetRoutineName);

            //Add Source Ladder Logic to Target Routine
            var targetRungs = targetRoutine.Descendants("Rung");
            if (targetRungs != null && targetRungs.Count() > 0)
                targetRungs.LastOrDefault().AddAfterSelf(sourceRungs); //add to end
            else
                targetRoutine.Descendants("RLLContent").LastOrDefault().Add(sourceRungs);
        }

        private IEnumerable<XElement> GetSourceRungs(string programName, string routineName)
        {
            var sourceProgram = this.TemplateDoc.Descendants("Program").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == programName)).FirstOrDefault();
            if (sourceProgram == null)
                return null;

            var sourceRoutine = sourceProgram.Descendants("Routine").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == routineName)).FirstOrDefault();
            return sourceRoutine?.Descendants("Rung");
        }

        private XDocument LoadSourceTemplate(string fileName)
        {
            return XDocument.Load(Path.Combine(_templatePath, fileName));
        }

        private void ReplaceTagMembers(XElement rung, string fromTag, string toTag)
        {
            var textNode = rung.Descendants("Text").FirstOrDefault();
            textNode.Value = textNode.Value.ToString().Replace(fromTag, toTag);
        }

        private void ReplaceCommentText(XElement rung, string fromText, string toText)
        {
            var textNode = rung.Descendants("Comment").FirstOrDefault();
            if (textNode == null)
                return;

            textNode.Value = textNode.Value.ToString().Replace(fromText, toText);
        }

        private void ReNumberLadder(XElement routines)
        {
            //re-number ladder rungs in each routines to be sequential as they may import with a conflicting sequence
            foreach (var routine in routines.Descendants("Routine"))
            {
                var i = 0;
                foreach (var rung in routine.Descendants("Rung"))
                {
                    XMLHelper.GetAttribute(rung, "Number").Value = i.ToString();
                    i++;
                }
            }
        }
    }

}
