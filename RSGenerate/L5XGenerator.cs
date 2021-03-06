﻿using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private string _ControllerName;

        private List<string> _DevicesProcessed;
        private List<string> _RoutinesAdded;

        private const string _SourceProgramName = "Templates"; //Name of the Logix Program that contains the Template Routines
        private const string _ConveyorTagToken = "CONVEYOR_XXX";
        private const string _IBConnectionTagToken = "CONVEYOR_IB_CONNECTION";
        private const string _OBConnectionTagToken = "CONVEYOR_OB_CONNECTION";
        private const string _S1TagToken = "CS_S1_XXX";
        private const string _SE1TagToken = "CS_SE1_XXX";
        private const string _SE2TagToken = "CS_SE2_XXX";
        private const string _SE3TagToken = "CS_SE3_XXX";
        private const string _SMTagToken = "CS_SM_XXX";
        private const string _SSTagToken = "CS_SS_XXX";
        private const string _EPCTagToken = "EPC_XXX";
        private const string _PETagToken = "PE_XXX";
        private const string _TipChuteTagToken = "TIP_CHUTE_XXX";
        private const string _ShuttleChuteTagToken = "SHUTTLE_CHUTE_XXX";
        private const string _GateTagToken = "GATE_XXX";
        private const string _ControllerNameTemplate = "PLC_XXXX_MCP01_Rack0_Slot0";
        private const string _SheetNameLayout = "System Layout";
        private const string _SheetNameOverview = "System Overview";
        private const string _SheetNameIODetails = "IO Details";

        private const int _ColumnIndexLabel = 2;
        private const int _ColumnIndexValue = 3;

        private readonly string[] _ValidCardTypes = { "1756-IA16", "1756-OW16I" };


        public L5XGenerator()
        {
            _DevicesProcessed = new List<string>();
            _RoutinesAdded = new List<string>();
        }

        public void LoadL5XTemplate(string filePath)
        {
            L5XTemplate = XDocument.Load(filePath);
            _L5XTemplatePath = System.IO.Path.GetDirectoryName(filePath);
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

            this.ExcelSystemDefinition.IsFirstRowAsColumnNames = true;
            DataSet dataset = this.ExcelSystemDefinition.AsDataSet();

            SetControllerAttributes(dataset);

            var routines = XMLHelper.GetOrCreateRoutinesElement((XElement)this.L5XTemplate.FirstNode);

            ConfigureIOData(dataset, routines);

            foreach (DataRow row in dataset.Tables[_SheetNameLayout].Rows)
                this.ProcessRow(row, routines);

            //Remove the Template Program from the Target before saving
            this.L5XTemplate.Descendants("Program").Where(p => p.Attributes().Any(a => a.Name == "Name" && a.Value == "Templates")).FirstOrDefault().Remove();

            //once all the routines are created and/or added to renumber them all sequentially
            this.ReNumberLadder(routines);

            _OutputFileName = _L5XTemplatePath + "\\" + _ControllerName + "_generated.L5X"; 
            this.L5XTemplate.Save(_OutputFileName);

            return _OutputFileName;
        }

        private void ConfigureIOData(DataSet dataset, XElement routines)
        {
            string currentRack = string.Empty;
            string currentSlot = string.Empty;
            string currentType = string.Empty;
            string lastSlot = string.Empty;

            var rows = dataset.Tables[_SheetNameIODetails].Rows;
            int emptyRowCount = 0;

            

            for (int currentRow = 0; currentRow < rows.Count; currentRow++)
            {
                var row = rows[currentRow];
                currentType = row[0].ToString();
                currentRack = row[1].ToString();
                lastSlot = currentSlot;
                currentSlot = row[2].ToString();

                var routineName = "IO_MAP";  // Default Routine name - used if nothing is specified in each row.
                if (!string.IsNullOrEmpty(row["Map Routine"].ToString()))
                    routineName = row["Map Routine"].ToString();
                
                if (!_ValidCardTypes.Contains(currentType) || string.IsNullOrEmpty(currentType))
                {
                    emptyRowCount++;
                    if (emptyRowCount >= 2) break;  //bail if there are more than 2 rows with no data - this signifies the end of the listing
                    continue;
                }

                emptyRowCount = 0;

                if (lastSlot != currentSlot)  //new Card
                {
                    var card = new IOCard
                    {
                        CardType = currentType,
                        Module = currentSlot,
                        Rack = currentRack
                    };

                    var endRow = currentRow + 16;
                    for (int ioRow = currentRow; ioRow < endRow; ioRow++ )
                    {
                        row = rows[ioRow];
                        card.Points.Add(new IOPoint
                        {
                            Bit = row[3].ToString(),
                            Description = row[4].ToString(), 
                            DeviceTag = row[5].ToString(),
                            Member = row[6].ToString()
                        });

                        currentRow = ioRow;
                    }

                    this.AddIOCardDetails(card);
                    var routine = XMLHelper.GetOrCreateRoutineWithName(routines, routineName);
                    this.CreateIOMappings(routine, card);
                }

            }
        }

        private void CreateIOMappings(XElement routine, IOCard card)
        {

            if (card.Points.All(p => string.IsNullOrEmpty(p.Member)))
                return;


            string cardComment =
@"********************************************************************************************************************************************************************************************************
{RACK_NUMBER}
{SLOT_NUMBER}
********************************************************************************************************************************************************************************************************";

            var rack = card.Rack.PadLeft(2, '0');
            var module = card.Module.PadLeft(2, '0');

            cardComment = cardComment.Replace("{RACK_NUMBER}", "RACK " + rack);
            cardComment = cardComment.Replace("{SLOT_NUMBER}", "SLOT " + module);

            var rungs = routine.Element("RLLContent");
            var commentRung = new XElement("Rung", new XAttribute("Number", "0"), new XAttribute("Type", "N"));
            commentRung.Add(new XElement("Comment", cardComment));
            commentRung.Add(new XElement("Text", "NOP();"));
            rungs.Add(commentRung);

            foreach (var point in card.Points)
            {
                var destTag = string.Empty;
                var sourceTag = "IO_SLOT_R" + rack + "_S" + module + "." + point.Bit;
                var mapRung = new XElement("Rung", new XAttribute("Number", "0"), new XAttribute("Type", "N"));
                
                var nopInputMap = "XIC({0})NOP({1});"; 
                var nopOutputMap = "NOP({0})OTE({1});";
                var ioMap = "XIC({0})OTE({1});";  

                var expression = string.Empty;

                if (!string.IsNullOrEmpty(point.Member) && !string.IsNullOrEmpty(point.DeviceTag))
                    destTag = point.DeviceTag + "." + point.Member;


                if (card.CardType.Contains("IA"))  //Input Card
                    expression = string.Format(string.IsNullOrEmpty(destTag) ? nopInputMap : ioMap, sourceTag, destTag);
                else if (card.CardType.Contains("OW"))  //Output Card
                    expression = string.Format(string.IsNullOrEmpty(destTag) ? nopOutputMap : ioMap, destTag, sourceTag);

                mapRung.Add(new XElement("Text", expression));
                rungs.Add(mapRung);
            }
        }

        private void AddIOCardDetails(IOCard card, bool autoAlias = true)
        {
            var rack = card.Rack.PadLeft(2, '0');
            var module = card.Module.PadLeft(2, '0');
            var type = card.CardType == "1756-IA16" ? "I" : "O";
            XElement tag = null;

            if (autoAlias)
            {            
                //we may auto-alias this - but for now we are not b/c the IO Tree will not be created and you can't go online to create the IO Tree with unresolveable references.
                string alias = string.Format("{0}:{1}:{2}.Data", rack == "00" ? "Local" : "ENET_RACK_" + rack, card.Module, type);
                tag = this.AddAliasControllerTag("IO_SLOT_R" + rack + "_S" + module, alias);
            }

            tag.Add(new XElement("Description", "IO Module Point Descriptions"));

            var comments = new XElement("Comments");
            foreach (var point in card.Points)
            {
                var comment = new XElement("Comment", new XAttribute("Operand", "." + point.Bit));
                comment.SetValue(point.Description);
                comments.Add(comment);
            }
            tag.Add(comments);


        }
        private void SetControllerAttributes(DataSet dataset)
        {
            //string controllerName = string.Empty;
            string controllerDesc = string.Empty;
            string siteNumber = string.Empty;
            string procLocation = string.Empty;
            foreach (DataRow row in dataset.Tables[_SheetNameOverview].Rows)
            {
                switch (row[_ColumnIndexLabel].ToString())
                {
                    case "Site Number":
                        siteNumber = row[_ColumnIndexValue].ToString();
                        break;
                    case "Project Name":
                        controllerDesc = row[_ColumnIndexValue].ToString() + Environment.NewLine + "MCS AUTOMATION" ;
                        break;
                    case "Processor Location":
                        procLocation = row[_ColumnIndexValue].ToString();
                        break;

                }
            }

            _ControllerName = _ControllerNameTemplate.Replace("XXXX", siteNumber).Replace("MCP01", procLocation);

            var contentNode = (XElement)this.L5XTemplate.FirstNode;
            XMLHelper.GetAttribute(contentNode, "TargetName").SetValue(_ControllerName);

            var controllerNode = XMLHelper.GetChildElement((XElement)contentNode, "Controller");
            XMLHelper.GetAttribute(controllerNode, "Name").SetValue(_ControllerName);

            var descNode = XMLHelper.GetChildElement((XElement)controllerNode, "Description");
            if (descNode != null)
                descNode.SetValue(controllerDesc);

        }

        private void ProcessRow(DataRow row, XElement routinesNode)
        {
            //read config data from row
            var deviceNumber = ExcelHelper.GetCellValue(row, "Device Number");
            var deviceType = ExcelHelper.GetCellValue(row, "Device Type");
            bool isVFD = ExcelHelper.GetCellValue(row, "VFD").ToUpper() == "X"; 
            string deviceName;

            if (string.IsNullOrEmpty(deviceNumber))
                return;

            switch (deviceType)
            {
                case "Conveyor":
                    deviceName = "CONVEYOR_" + deviceNumber;
                    this.AddControllerTag(deviceName, "FXG_CONVEYOR");
                    this.AddMotorLogic(routinesNode, deviceName, row);
                    //Add Disconnect Fault Code
                    this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_DISC", "MOTOR_DISC", _ConveyorTagToken, deviceName);
                    //Import Motor Overload Code
                    this.ImportSourceLadderRungs(routinesNode, "FXG_MOTOR_OVLD" + (isVFD ? "_VFD" : ""), "MOTOR_OVLD", _ConveyorTagToken, deviceName);
                    break;

                case "Tip Chute":
                    deviceName = "TIP_CHUTE_" + deviceNumber;
                    this.AddControllerTag(deviceName, "TIP_CHUTE");
                    this.ImportSourceLadderRungs(routinesNode, "FXG_TIP_CHUTE", "TIP_CHUTES", _TipChuteTagToken, deviceName);
                    break;

                case "Shuttle Chute":
                    deviceName = "SHUTTLE_CHUTE_" + deviceNumber;
                    this.AddControllerTag(deviceName, "SHUTTLE_CHUTE");
                    this.ImportSourceLadderRungs(routinesNode, "FXG_SHUTTLE_CHUTE", "SHUTTLE_CHUTES", _ShuttleChuteTagToken, deviceName);
                    break;

                case "Gate":
                    deviceName = "GATE_" + deviceNumber;
                    this.AddControllerTag(deviceName, "GATE");
                    this.ImportSourceLadderRungs(routinesNode, "FXG_GATE", "GATES", _GateTagToken, deviceName);
                    break;
            }
                 
            //Special handling for SE stations
            //If there is both an SE1 and an SE2 station listed, assume they work in pairs and use the _DUAL version of the template
            if (!string.IsNullOrEmpty(ExcelHelper.GetCellValue(row, "SE1")) && !string.IsNullOrEmpty(ExcelHelper.GetCellValue(row, "SE2")))
            {
                this.AddDeviceRoutine(routinesNode, row, "SE1", "CS_SE", "FXG_CS_SE_DUAL", "CS_SE", _SE1TagToken);
                this.AddDeviceRoutine(routinesNode, row, "SE2", "CS_SE", "FXG_CS_SE_DUAL", "CS_SE", _SE1TagToken);
            }
            else
            {
                this.AddDeviceRoutine(routinesNode, row, "SE1", "CS_SE", "FXG_CS_SE_SINGLE", "CS_SE", _SE1TagToken);
                this.AddDeviceRoutine(routinesNode, row, "SE2", "CS_SE", "FXG_CS_SE_SINGLE", "CS_SE", _SE1TagToken);
            }

            this.AddDeviceRoutine(routinesNode, row, "SE3", "CS_SE", "FXG_CS_SE_SINGLE", "CS_SE", _SE1TagToken);

            this.AddDeviceRoutine(routinesNode, row, "S1", "CS_S1", "FXG_CS_S1", "CS_S1", _S1TagToken);

            this.AddDeviceRoutine(routinesNode, row, "SM1", "CS_SM", "FXG_CS_SM", "CS_SM", _SMTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SM2", "CS_SM", "FXG_CS_SM", "CS_SM", _SMTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SM3", "CS_SM", "FXG_CS_SM", "CS_SM", _SMTagToken);

            this.AddDeviceRoutine(routinesNode, row, "SS1", "CS_SS", "FXG_CS_SS", "CS_SS", _SSTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SS2", "CS_SS", "FXG_CS_SS", "CS_SS", _SSTagToken);
            this.AddDeviceRoutine(routinesNode, row, "SS3", "CS_SS", "FXG_CS_SS", "CS_SS", _SSTagToken);

            this.AddDeviceRoutine(routinesNode, row, "EPC1", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);
            this.AddDeviceRoutine(routinesNode, row, "EPC2", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);
            this.AddDeviceRoutine(routinesNode, row, "EPC3", "CS_EPC", "FXG_CS_EPC", "CS_EPC", _EPCTagToken);

            this.AddDeviceRoutine(routinesNode, row, "PE", "PE", "FXG_PE", "PE", _PETagToken);

        }

        private void AddMotorLogic(XElement routinesNode, string conveyorTagName, DataRow row)
        {
            string ibInterlockExpression = "XIC(SYSTEM.IB_MODE)XIC(CONVEYOR_IB_CONNECTION.RUN_INTERLOCK_READY)";
            string obInterlockExpression = "XIC(SYSTEM.OB_MODE)XIC(CONVEYOR_OB_CONNECTION.RUN_INTERLOCK_READY)";
            string fwdInterlockCoilExpression = "OTE(CONVEYOR_XXX.RUN_FWD_INTERLOCK_READY);";
            string revInterlockCoilExpression = "OTE(CONVEYOR_XXX.RUN_REV_INTERLOCK_READY);";

            string ibUsedInModeExpression = "XIC(SYSTEM.IB_MODE)";
            string obUsedInModeExpression = "XIC(SYSTEM.OB_MODE)";
            string usedInModeCoilExpression = "OTE(CONVEYOR_XXX.USED_IN_MODE);";

            //extract motor info from sheet
            bool inbound = ExcelHelper.GetCellValue(row, "Inbound").ToUpper() == "X";
            bool outbound = ExcelHelper.GetCellValue(row, "Outbound").ToUpper() == "X";
            bool reversing = ExcelHelper.GetCellValue(row, "Reversing").ToUpper() == "X";
            string ibConnection = ExcelHelper.GetCellValue(row, "IB Connection");
            string obConnection = ExcelHelper.GetCellValue(row, "OB Connection");
            string motorRoutine = ExcelHelper.GetCellValue(row, "Motor Control Template Routine");
            string targetRoutineName = ExcelHelper.GetCellValue(row, "Target Routine Name");

            //Set up special case Tag Name replacements - this is passed to import routine
            var replaceTags = EnumerateTags(row, null);
            replaceTags.Add(_IBConnectionTagToken, "CONVEYOR_" + ibConnection);
            replaceTags.Add(_OBConnectionTagToken, "CONVEYOR_" + obConnection);

            //set up special case Logic replacements - this is passed to import routine.
            var customLogicList = new Dictionary<string, string>();

            string fwdInterlockExpression = string.Empty;
            string revInterlockExpression = string.Empty;
            string usedInModeExpression = string.Empty;


            if (ibConnection.ToUpper() == "END")
                ibInterlockExpression = "XIC(SYSTEM.IB_MODE)";

            if (obConnection.ToUpper() == "END")
                obInterlockExpression = "XIC(SYSTEM.OB_MODE)";

            if (inbound && outbound)
            {
                if (reversing)
                {
                    fwdInterlockExpression = ibInterlockExpression + fwdInterlockCoilExpression;
                    revInterlockExpression = obInterlockExpression + revInterlockCoilExpression;
                }
                else
                    fwdInterlockExpression = "[" + ibInterlockExpression + "," + obInterlockExpression + "]" + fwdInterlockCoilExpression;

                usedInModeExpression = "[" + ibUsedInModeExpression + "," + obUsedInModeExpression + "]" + usedInModeCoilExpression;
            }
            else if (inbound)
            {
                fwdInterlockExpression = ibInterlockExpression + fwdInterlockCoilExpression;
                usedInModeExpression = ibUsedInModeExpression + usedInModeCoilExpression;
            }
            else if (outbound)
            {
                fwdInterlockExpression = obInterlockExpression + fwdInterlockCoilExpression;
                usedInModeExpression = obUsedInModeExpression + usedInModeCoilExpression;
            }

            customLogicList.Add("{{RUN_FWD_INTERLOCK_READY}}", fwdInterlockExpression);
            customLogicList.Add("{{RUN_REV_INTERLOCK_READY}}", revInterlockExpression);
            customLogicList.Add("{{USED_IN_MODE}}", usedInModeExpression);


            this.ImportSourceLadderRungs(routinesNode, motorRoutine, targetRoutineName, _ConveyorTagToken, conveyorTagName, replaceTags, customLogicList);
        }

        private void AddDeviceRoutine(XElement routinesNode, DataRow row, string columnName, string tagDataType, string templateRoutineName, string targetRoutineName, string tagToken, Dictionary<string, string> replaceTags = null)
        {
            var deviceName = ExcelHelper.GetCellValue(row, columnName);
            if (string.IsNullOrEmpty(deviceName))
                return;

            if (!deviceName.StartsWith("PE"))
                deviceName = "CS_" + deviceName;

            if (_DevicesProcessed.Contains(deviceName))
                return;

            _DevicesProcessed.Add(deviceName);

            this.AddControllerTag(deviceName, tagDataType);

            replaceTags = EnumerateTags(row, replaceTags);
                
            this.ImportSourceLadderRungs(routinesNode, templateRoutineName, targetRoutineName, tagToken, deviceName, replaceTags);
        }

        private Dictionary<string, string> EnumerateTags(DataRow row, Dictionary<string, string> replaceTags)
        {
            Debug.Print("Enumerating row " + row["Device Number"].ToString());

            if (replaceTags == null)
                replaceTags = new Dictionary<string, string>();

            for (int i = 12; i <= 26; i++)
            {
                if (row.ItemArray.Length > i && !string.IsNullOrEmpty(row[i].ToString()))
                {
                    var col = row.Table.Columns[i].ColumnName.Trim();
                    string tagName = string.Empty;
                    string tagValue = string.Empty;

                    if (col == "TIP_CHUTE")
                    {
                        tagName = "TIP_CHUTE_XXX";
                        tagValue = ExcelHelper.GetCellValue(row, i);
                    }
                    else
                    {
                        tagName = "CS_" + col + "_XXX";
                        tagValue = "CS_" + ExcelHelper.GetCellValue(row, i);
                    }


                    if (!replaceTags.ContainsKey(tagName))
                        replaceTags.Add(tagName, tagValue);
                }
            }
            return replaceTags;
        }
        private XElement AddControllerTag(string tagName, string dataType)
        {
            var tagNode = XMLHelper.CreateTag(tagName, dataType);
            this.L5XTemplate.Descendants("Controller").FirstOrDefault().Element("Tags").Add(tagNode);
            return tagNode;
        }

        private XElement AddAliasControllerTag(string tagName, string alias)
        {
            var tagNode = XMLHelper.CreateAliasTag(tagName, alias);
            this.L5XTemplate.Descendants("Controller").FirstOrDefault().Element("Tags").Add(tagNode);
            return tagNode;
        }

        private void ImportSourceLadderRungs(XElement routines, string sourceRountineName, string targetRoutineName, string sourceTagName, string targetTagName, Dictionary<string, string> additionalTagReplacements = null, Dictionary<string, string> customRungText = null, bool ensureInMainroutineScanList = true)
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
                var comment = XMLHelper.GetChildElement(rung, "Comment");
                var text = XMLHelper.GetChildElement(rung, "Text");

                this.ApplyCustomRungLogic(customRungText, comment, text);

                //replace main tag for routine
                this.ReplaceTagMembers(rung, sourceTagName, targetTagName);

                foreach (var item in additionalTagReplacements ?? new Dictionary<string, string>())
                    this.ReplaceTagMembers(rung, item.Key, item.Value);

                //update device name in comments using token {DEVICE_NUMBER}
                this.ReplaceCommentText(rung, "{DEVICE_NUMBER}", targetTagName);

            }

            XMLHelper.AddSourceRungsToTargetRoutine(sourceRungs, routines, targetRoutineName);

            if (ensureInMainroutineScanList)
                EnsureRoutineInMainRoutineScanList(routines, targetRoutineName);
        }

        private void ApplyCustomRungLogic(Dictionary<string, string> customRungText, XElement comment, XElement text)
        {
            //Check to see if any Custom Rung Logic applies
            foreach (var item in customRungText ?? new Dictionary<string, string>())
            {
                if (comment != null && text != null && comment.Value.Contains(item.Key))
                {
                    //replace logic in template with Custom Logic from List
                    text.Value = item.Value;

                    //sanitize comment to remove the {{token}} from the text
                    comment.Value = Regex.Replace(comment.Value, "{{.*?}}", "");
                }
            }
        }

        private void EnsureRoutineInMainRoutineScanList(XElement routines, string targetRoutineName)
        {
            if (_RoutinesAdded.Contains(targetRoutineName))
                return;

            string rungText = string.Format("JSR({0},0);", targetRoutineName);

            var sourceRungs = new List<XElement>();
            sourceRungs.Add(XMLHelper.CreateRung(rungText));
            XMLHelper.AddSourceRungsToTargetRoutine(sourceRungs, routines, "MainRoutine");
            _RoutinesAdded.Add(targetRoutineName);
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
