using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using System.Collections;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model.Operations;

namespace Export
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            MyModel = new Model();
        }

        private readonly Model MyModel;

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox3.Text = "00";
            textBox2.Text = "1101";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ObjectFilterExpressions.Type obj1 = new ObjectFilterExpressions.Type();
            NumericConstantFilterExpression type1 = new NumericConstantFilterExpression(Tekla.Structures.TeklaStructuresDatabaseTypeEnum.ASSEMBLY);
            NumericConstantFilterExpression type2 = new NumericConstantFilterExpression(Tekla.Structures.TeklaStructuresDatabaseTypeEnum.PART);

            string faza_wpisana = textBox2.Text;
            AssemblyFilterExpressions.Phase exp1 = new AssemblyFilterExpressions.Phase();
            StringConstantFilterExpression faza = new StringConstantFilterExpression(faza_wpisana);
            BinaryFilterExpression bexp = new BinaryFilterExpression(exp1, StringOperatorType.IS_EQUAL, faza);
            BinaryFilterExpression bexp2 = new BinaryFilterExpression(obj1, NumericOperatorType.IS_EQUAL, type1);
            BinaryFilterExpression bexp3 = new BinaryFilterExpression(obj1, NumericOperatorType.IS_EQUAL, type2);
            BinaryFilterExpressionCollection bexpcoll = new BinaryFilterExpressionCollection();
            bexpcoll.Add(new BinaryFilterExpressionItem(bexp, BinaryFilterOperatorType.BOOLEAN_AND));
            bexpcoll.Add(new BinaryFilterExpressionItem(bexp2, BinaryFilterOperatorType.BOOLEAN_AND));

            BinaryFilterExpressionCollection bexpcoll2 = new BinaryFilterExpressionCollection();
            bexpcoll2.Add(new BinaryFilterExpressionItem(bexp, BinaryFilterOperatorType.BOOLEAN_AND));
            bexpcoll2.Add(new BinaryFilterExpressionItem(bexp3, BinaryFilterOperatorType.BOOLEAN_AND));

            ModelObjectEnumerator objekty = MyModel.GetModelObjectSelector().GetObjectsByFilter(bexpcoll);

            ModelObjectEnumerator objekty2 = MyModel.GetModelObjectSelector().GetObjectsByFilter(bexpcoll2);
            //textBox5.Text = Convert.ToString(objekty.GetSize());
            //textBox4.Text = Convert.ToString(objekty2.GetSize());

            DrawingHandler DrawingHandler2 = new DrawingHandler();
            DrawingEnumerator rysunki = DrawingHandler2.GetDrawings();

            List<Drawing> drawings = new List<Drawing>();
            foreach (Drawing drawing in rysunki)
                drawings.Add(drawing);
            List<Drawing> assembly_drawings = new List<Drawing>(drawings.OfType<AssemblyDrawing>());
            List<Drawing> part_drawings = new List<Drawing>(drawings.OfType<SinglePartDrawing>());


            textBox4.Text = Convert.ToString(part_drawings.Count);

            ArrayList assembly_list = new ArrayList();
            foreach (Tekla.Structures.Model.ModelObject modelob in objekty)
                assembly_list.Add(modelob);

            ArrayList part_list = new ArrayList();
            foreach (Tekla.Structures.Model.ModelObject modelob2 in objekty2)
                part_list.Add(modelob2);

            string TSBinaryDir = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XSBIN", ref TSBinaryDir);
            //DrawingHandler2.SetActiveDrawing(assembly_drawings.ElementAt(1));
            var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
            var printerLocation = installLoc + @"nt\bin\applications\Tekla\Drawings\DwgExport\Dwg.exe";
            string dwgxportParams = "D:\\";
            string attributeFile = "bez osadzania";
            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = printerLocation;
            process.StartInfo.Arguments = "export outputDirectory=" + dwgxportParams + " " + "settingFile=bez_osadzania";
            process.StartInfo.WorkingDirectory = "D:\\";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();






            MyModel.CommitChanges();
        }

        private void druk_pdf(string name)
        {

            //DrawingHandler2.SetActiveDrawing(assembly_drawings.ElementAt(1));
            //var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
            //var printerLocation = installLoc + @"nt\bin\applications\Tekla\Model\DPMPrinter\DPMPrinterCommand.exe";
           // var arg = string.Format("printActive:true printer:pdf colormode:BlackAndWhite paper:Auto out:\"D:\\{0}.pdf\"", assembly_drawings.ElementAt(1).Mark + "_" + assembly_drawings.ElementAt(1).Name);
           // Console.WriteLine("Args = '{0}'", arg);
            //var process = new System.Diagnostics.Process();
            //process.StartInfo.FileName = printerLocation;
            //process.StartInfo.Arguments = arg;
           // process.StartInfo.WorkingDirectory = "D:\\";
            //process.StartInfo.UseShellExecute = false;
            //process.StartInfo.RedirectStandardOutput = true;
            //process.OutputDataReceived += (sender2, args) => Console.WriteLine("received output: {0}", args.Data);
           // process.Start();
           // process.BeginOutputReadLine();
           // process.WaitForExit();

            //DPMPrinterAttributes dpmprint = new DPMPrinterAttributes();
            // dpmprint.OutputFileName = "D:\\test.pdf";
            //dpmprint.OutputType = DotPrintOutputType.PDF;
            //dpmprint.ColorMode = DotPrintColor.BlackAndWhite;
            //DrawingHandler2.PrintDrawing(part_drawings.ElementAt(1), dpmprint);


        }

        private void raporty (string reklama)

        {
            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            //MS.Select(assembly_list);
            Operation.CreateReportFromSelected("003 - Lista elementów wysyłkowych.pdf", "003 - Lista elementów wysyłkowych.pdf", "22", "22", "22");
        }

        private void nc_export(string nc)
        {
            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            //MS.Select(part_list);
            bool test = Operation.CreateNCFilesFromSelected("Blachy", "D:\\");
            Operation.CreateNCFilesFromSelected("Profile", "D:\\");


        }

    }
}
