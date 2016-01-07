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
        public XDocument ProjectTemplate { get; private set; }
        public IExcelDataReader SystemDefinition { get; private set; }
        public XDocument Output { get; private set; }

        private string _templatePath;


        public void LoadTemplate(string filePath)
        {
            ProjectTemplate = XDocument.Load(filePath);
            _templatePath = System.IO.Path.GetDirectoryName(filePath);
        }

        public void LoadDefinition(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var fileInfo = new FileInfo(filePath);

            if (fileInfo.Extension == ".xlsx")
                SystemDefinition = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        public void Generate()
        {
            if (ProjectTemplate == null)
                throw new ArgumentNullException("You must first specify a ProjectTemplate prior to calling Generate()");

            if (SystemDefinition == null)
                throw new ArgumentNullException("You must first specify a SystemDefinition prior to calling Generate()");

            var reader = SystemDefinition;
            reader.IsFirstRowAsColumnNames = true;
            DataSet dataset = reader.AsDataSet();
            string sheetName = "Sheet1";
            var programRoutines = XMLHelper.GetOrCreateRoutinesElement((XElement)this.ProjectTemplate.FirstNode);

            foreach (DataRow row in dataset.Tables[sheetName].Rows)
            {
                var motorControlTemplate = row["Motor Control Template"].ToString();
                var routineName = row["RoutineName"].ToString();
                var conveyorNumber = row["Conveyor Number"].ToString();
                var downstreamNumber = row["Downstream Conveyor"].ToString();

                if (string.IsNullOrEmpty(conveyorNumber))
                    continue;

                //Load the control template from the local L5X file.
                XDocument ladderTemplate = XDocument.Load(Path.Combine(_templatePath, motorControlTemplate));

                var targetRoutine = XMLHelper.GetOrCreateRoutineWithName(programRoutines, routineName);

                var sourceRungs = ladderTemplate.Descendants("RLLContent").FirstOrDefault().Descendants("Rung");
                foreach(var rung in sourceRungs)
                {
                    var textNode = rung.Descendants("Text").FirstOrDefault();
                    var rungLadder = textNode.Value.ToString();
                    rungLadder = rungLadder.Replace("CONV_CURRENT", conveyorNumber);
                    rungLadder = rungLadder.Replace("CONV_DOWNSTREAM", downstreamNumber);
                    textNode.Value = rungLadder;
                }

                var targetRungs = targetRoutine.Descendants("RLLContent").FirstOrDefault().Descendants("Rung");
                if (targetRungs?.Count() > 0)
                    targetRungs.LastOrDefault().AddAfterSelf(sourceRungs);
                else
                    targetRoutine.Descendants("RLLContent").FirstOrDefault().Add(sourceRungs);
            }

            this.RenumberLadder(programRoutines);

            this.ProjectTemplate.Save(_templatePath + "\\TestOutput.L5X");

        }

        private void RenumberLadder(XElement programRoutines)
        {
            //renumber ladder
            foreach (var routine in programRoutines.Descendants())
            {
                var i = 0;
                foreach (var rung in routine.Descendants("Rung"))
                {
                    rung.Attributes().Where(a => a.Name == "Number").First().Value = i.ToString();
                    i++;
                }
            }
        }
    }

}
