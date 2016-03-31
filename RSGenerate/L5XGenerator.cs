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
        public XDocument L5XTemplate { get; private set; }
        public IExcelDataReader ExcelSystemDefinition { get; private set; }
        public XDocument L5XOutput { get; private set; }

        private string _L5XTemplatePath;
        private string _OutputFileName;

        private const string _SourceProgramName = "Templates"; //Name of the Logix Program that contains the Template Routines

        public void LoadL5XTemplate(string filePath)
        {
            L5XTemplate = XDocument.Load(filePath);
            _L5XTemplatePath = System.IO.Path.GetDirectoryName(filePath);
            _OutputFileName = _L5XTemplatePath + "\\GeneratedProject.L5X";
        }

        public void LoadXLSXDefinition(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var fileInfo = new FileInfo(filePath);

            if (fileInfo.Extension == ".xlsx")
                ExcelSystemDefinition = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        public string GenerateAndSaveOutput()
        {
            if (L5XTemplate == null)
                throw new ArgumentNullException("You must first specify a TemplateDoc prior to calling Generate()");

            if (ExcelSystemDefinition == null)
                throw new ArgumentNullException("You must first specify a SystemDefinition prior to calling Generate()");

            var reader = ExcelSystemDefinition;
            reader.IsFirstRowAsColumnNames = true;
            DataSet dataset = reader.AsDataSet();
            string sheetName = "Sheet1";

            var routines = XMLHelper.GetOrCreateRoutinesElement((XElement)this.L5XTemplate.FirstNode);

            foreach (DataRow row in dataset.Tables[sheetName].Rows)
                this.ProcessRow(row, routines);

            //Remove the Template Program from the Target before saving
            this.L5XTemplate.Descendants("Program").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == "Templates")).FirstOrDefault().Remove();

            //once all the routines are created and/or added to renumber them all sequentially
            this.ReNumberLadder(routines);

            
            this.L5XTemplate.Save(_OutputFileName);

            return _OutputFileName;
        }
        private string GetCellValue(DataRow row, string col)
        {
            return row[col].ToString().Replace('-','_');
        }

        private void ProcessRow(DataRow row, XElement routinesNode)
        {
            //read config data from row
            var conveyorNumber = GetCellValue(row, "Conveyor Number");
            var inboundConnection = "CONVEYOR_" + GetCellValue(row, "IB Connection");
            var outboundConnection = "CONVEYOR_" + GetCellValue(row, "OB Connection");
            var targetRoutineName = GetCellValue(row, "Target Routine Name");
            var motorTemplateRoutineName = GetCellValue(row, "Motor Control Template Routine");
            string deviceName = null;

            if (string.IsNullOrEmpty(conveyorNumber))
                return;

            conveyorNumber = "CONVEYOR_" + conveyorNumber;
            this.AddControllerTag(conveyorNumber, "FXG_CONVEYOR");

            //Import Disconnect Fault Code
            this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_DISC", "MOTOR_DISC", "CONVEYOR_XXX", conveyorNumber);

            //Import Motor Overload Code
            this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_OVLD", "MOTOR_OVLD", "CONVEYOR_XXX", conveyorNumber);

            //Import SE Logic
            deviceName =  GetCellValue(row, "SE1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SE");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SE", "CS_SE", "CS_SE1_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "SE2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SE");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SE", "CS_SE", "CS_SE1_XXX", deviceName);
            }

            ////Import SM Logic
            deviceName = GetCellValue(row, "SM1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SM");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SM", "CS_SM", "CS_SM_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "SM2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SM");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SM", "CS_SM", "CS_SM_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "SS1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SS");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SS", "CS_SS", "CS_SS_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "SS2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_SS");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_SS", "CS_SS", "CS_SS_XXX", deviceName);
            }

            ////Import S1 Logic
            //if (!string.IsNullOrEmpty(GetCellValue(row, "S1")))
            //    this.ImportSourceLadderRungs(routines, "FXG_CS_S1", "CS_S1", conveyorNumber, inboundConnection, outboundConnection, GetCellValue(row, "S1"));

            deviceName = GetCellValue(row, "EPC1");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "EPC2");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName);
            }

            deviceName = GetCellValue(row, "EPC3");
            if (!string.IsNullOrEmpty(deviceName))
            {
                //deviceName = "CS_" + deviceName;
                this.AddControllerTag(deviceName, "CS_EPC");
                this.ImportSourceLadderRungs(routinesNode, "FXG_CS_EPC", "CS_EPC", "EPC_XXX", deviceName);
            }

            //Import Motor Control Code
           // this.ImportSourceLadderRungs(routines, motorTemplateRoutineName, targetRoutineName, conveyorNumber, inboundConnection, outboundConnection, null);
        }

        private void AddControllerTag(string tagName, string dataType)
        {
            var tagNode = XMLHelper.CreateTag(tagName, dataType);
            this.L5XTemplate.Descendants("Controller").FirstOrDefault().Element("Tags").Add(tagNode);
        }

        private void ImportSourceLadderRungs(XElement routines, string sourceRountineName, string targetRoutineName, string sourceTagName, string targetTagName, Dictionary<string, string> additionalTagReplacements = null)
        {

            //first we clone the ladder rungs from the Source Template that is specified
            var rungs = this.GetSourceRungs(_SourceProgramName, sourceRountineName);
            if (rungs == null)
                return;

            var clone = new XElement("Source", rungs);
            var sourceRungs = clone.Elements("Rung");

            //Perform Tag Replacement in ladder text
            foreach (var rung in sourceRungs)
            {
                //replace main tag for routine
                this.ReplaceTagMembers(rung, sourceTagName, targetTagName);

                //update device name in comments using token {DEVICE_NUMBER}
                this.ReplaceCommentText(rung, "{DEVICE_NUMBER}", targetTagName);

                foreach (var item in additionalTagReplacements ?? new Dictionary<string, string>())
                    this.ReplaceTagMembers(rung, item.Key, item.Value);
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
            var sourceProgram = this.L5XTemplate.Descendants("Program").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == programName)).FirstOrDefault();
            if (sourceProgram == null)
                return null;

            var sourceRoutine = sourceProgram.Descendants("Routine").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == routineName)).FirstOrDefault();
            return sourceRoutine?.Descendants("Rung");
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
