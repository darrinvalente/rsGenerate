using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Diagnostics;

namespace RSGenerate
{
    public partial class Form1 : Form
    {
        private L5XGenerator _generator;

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            _generator = new L5XGenerator();
        }

        private void btnChooseTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.InitialDirectory = "Z:\\SharePoint\\AI STRAT - Documents\\PLC\\Code Generation";   //CHANGE THIS HARDCODED PATH
            openFileDialog1.Filter = "RSLogix Export Files (.L5X)|*.L5X|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;
            var result = openFileDialog1.ShowDialog();
            string fileName = null;
            if (result == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label1.Text = fileName;
            }

            _generator.LoadL5XTemplate(fileName);

            txtLog.Text += string.Format("Template File {0} chosen for processing.", fileName) + "\r\n";

            //var root = (XElement)_generator.ProjectTemplate.FirstNode;
            //Debug.Print("Read file.  Examining root node " + root.Name.ToString());
            //Debug.Print("File created with RSLogix5000 version " + XMLHelper.GetAttribute(root, "SoftwareRevision"));

            //var controller = (XElement)root.FirstNode;
            //Debug.Print("Found Controller Node.  Contoller is of Type " + XMLHelper.GetAttribute(controller, "ProcessorType"));

            //var dataTypes = controller.Nodes().OfType<XElement>().Where(n => n.Name == "DataTypes").Nodes().OfType<XElement>();
            //Debug.Print(dataTypes?.Count() + " DataType Nodes Found.");

            //foreach (var node in dataTypes)
            //    Debug.Print(XMLHelper.GetAttribute(node, "Name"));

            //var modules = controller.Nodes().OfType<XElement>().Where(n => n.Name == "Modules").Nodes().OfType<XElement>();
            //Debug.Print(modules?.Count() + " Module Nodes Found.");

            //foreach (var node in modules)
            //    Debug.Print(XMLHelper.GetAttribute(node, "Name"));

            //var tags = controller.Nodes().OfType<XElement>().Where(n => n.Name == "Tags").Nodes().OfType<XElement>();
            //Debug.Print(tags?.Count() + " Tag Nodes Found.");

            //foreach (var node in tags)
            //    Debug.Print(XMLHelper.GetAttribute(node, "Name"));
        }

        private void btnChooseDefinition_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.InitialDirectory = "Z:\\SharePoint\\AI STRAT - Documents\\PLC\\Code Generation";   //CHANGE THIS HARDCODED PATH
            openFileDialog1.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;
            var result = openFileDialog1.ShowDialog();
            string fileName = null;
            if (result == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label2.Text = fileName;
            }

            _generator.LoadXLSXDefinition(fileName);

            txtLog.Text += string.Format("Definition File {0} chosen for processing.", fileName) + "\r\n";
        }

        private void btnGenerateOutput_Click(object sender, EventArgs e)
        {
            txtLog.Text += "Starting Template Generation." + "\r\n";
            var outputFile = _generator.GenerateAndSaveOutput();
            txtLog.Text += string.Format("Output file generation complete.  Saved to {0}", outputFile) + "\r\n";

        }
    }
}
