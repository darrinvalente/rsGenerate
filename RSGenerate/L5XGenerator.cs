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
        private List<string> _DevicesProcessed;

        private const string _SourceProgramName = "Templates"; //Name of the Logix Program that contains the Template Routines
        private const string _ConveyorTagToken = "CONVEYOR_XXX";
        private const string _IBConnectionTagToken = "CONVEYOR_IB_CONNECTION";
        private const string _OBConnectionTagToken = "CONVEYOR_OB_CONNECTION";
        private const string _SE1TagToken = "CS_SE1_XXX";
        private const string _SE2TagToken = "CS_SE2_XXX";
        private const string _SMTagToken = "CS_SM_XXX";
        private const string _SSTagToken = "CS_SS_XXX";
        private const string _EPCTagToken = "EPC_XXX";

        public L5XGenerator()
        {
            _DevicesProcessed = new List<string>();
        }

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

        private void ProcessRow(DataRow row, XElement routinesNode)
        {
            //read config data from row
            var conveyorNumber = ExcelHelper.GetCellValue(row, "Conveyor Number");
            if (string.IsNullOrEmpty(conveyorNumber))
                return;

            conveyorNumber = "CONVEYOR_" + conveyorNumber;
            this.AddControllerTag(conveyorNumber, "FXG_CONVEYOR");

            this.AddMotorRoutines(routinesNode, conveyorNumber, row);

            this.AddDeviceRoutine(routinesNode, row, "SE1", "CS_SE", "FXG_CS_SE", "CS_SE", _SE1TagToken);
            this.AddDeviceRoutine(routinesNode, row, "SE2", "CS_SE", "FXG_CS_SE", "CS_SE", _SE1TagToken);

            this.AddDeviceRoutine(routinesNode, row, "SM1", "CS_SM", "FXG_CS_SM", "CS_SM", _SMTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SM2", "CS_SM", "FXG_CS_SM", "CS_SM", _SMTagToken);

            this.AddDeviceRoutine(routinesNode, row, "SS1", "CS_SS", "FXG_CS_SS", "CS_SS", _SSTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SS2", "CS_SS", "FXG_CS_SS", "CS_SS", _SSTagToken);

            this.AddDeviceRoutine(routinesNode, row, "EPC1", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);
            this.AddDeviceRoutine(routinesNode, row, "EPC2", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);
            this.AddDeviceRoutine(routinesNode, row, "EPC3", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);


        }

        private void AddMotorRoutines(XElement routinesNode, string conveyorTagName, DataRow row)
        {
            //Add Disconnect Fault Code
            this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_DISC", "MOTOR_DISC", _ConveyorTagToken, conveyorTagName);

            //Import Motor Overload Code
            this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_OVLD", "MOTOR_OVLD", _ConveyorTagToken, conveyorTagName);

            //Import Motor Control Code
            var motorRoutine = ExcelHelper.GetCellValue(row, "Motor Control Template Routine");
            var targetRoutineName = ExcelHelper.GetCellValue(row, "Target Routine Name");

            var dict = new Dictionary<string, string>();
            dict.Add(_IBConnectionTagToken, "CONVEYOR_" + ExcelHelper.GetCellValue(row, "IB Connection"));
            dict.Add(_OBConnectionTagToken, "CONVEYOR_" + ExcelHelper.GetCellValue(row, "OB Connection"));
            this.ImportSourceLadderRungs(routinesNode, motorRoutine, targetRoutineName, _ConveyorTagToken, conveyorTagName, dict);
        }

        private void AddDeviceRoutine(XElement routinesNode, DataRow row, string columnName, string tagDataType, string templateRoutineName, string targetRoutineName, string tagToken)
        {
            var deviceName = ExcelHelper.GetCellValue(row, columnName);
            if (string.IsNullOrEmpty(deviceName))
                return;

            deviceName = "CS_" + deviceName;
            if (_DevicesProcessed.Contains(deviceName))
                return;

            this.AddControllerTag(deviceName, tagDataType);
            this.ImportSourceLadderRungs(routinesNode, templateRoutineName, targetRoutineName, tagToken, deviceName);
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
