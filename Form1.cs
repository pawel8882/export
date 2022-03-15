using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using System.Collections;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model.Operations;
using ModelObject = Tekla.Structures.Model.ModelObject;

namespace Export
{
    public partial class Form1 : Form
    {

        private string location;
        private string rev;
        private string faza;
        private string nazwa;
        private string folder_dwa;
        private ArrayList part_list;
        private ArrayList assembly_list;
        private List<Tekla.Structures.Model.ModelObject> zespoly_faza;
        private List<Tekla.Structures.Model.ModelObject> party_faza;
        private List<Drawing> zespoly_rysunki;
        private List<Drawing> party_rysunki;
        private DrawingHandler DrawingHandler2;
        private List<Drawing> zespoly_faza_rysunki;
        private List<Drawing> party_faza_rysunki;

        public Form1()
        {
            InitializeComponent();
            MyModel = new Model();

            textBox3.Text = "0";
            textBox2.Text = "1101";
            textBox6.Text = "Kratownice";
        }

        private readonly Model MyModel;

        private void Form1_Load(object sender, EventArgs e)
        {

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

        private void button3_Click(object sender, EventArgs e)
        {
            DateTime teraz = DateTime.Today;
            string data = teraz.ToString("yyyy.MM.dd");
            string folder = location + "\\" + faza + " - " + nazwa;
            folder_dwa = folder + "\\" + faza + "_" + "rev" + rev + "_" + data;
            Directory.CreateDirectory(folder);
            Directory.CreateDirectory(folder_dwa);
            Directory.CreateDirectory(folder_dwa + "\\DWG");
            Directory.CreateDirectory(folder_dwa + "\\DWG\\Zespoły");
            Directory.CreateDirectory(folder_dwa + "\\DWG\\Elementy");
            Directory.CreateDirectory(folder_dwa + "\\NC");
            Directory.CreateDirectory(folder_dwa + "\\LISTY");
            Directory.CreateDirectory(folder_dwa + "\\PDF");
            Directory.CreateDirectory(folder_dwa + "\\DXF");
            Directory.CreateDirectory(folder_dwa + "\\PDF\\Zespoły");
            Directory.CreateDirectory(folder_dwa + "\\PDF\\Elementy");


            //List<Drawing> zespoly_faza_rysunki_beta = new List<Drawing>(lista_rysunkow_zespoly(zespoly_faza, zespoly_rysunki));
            //List<Drawing> party_faza_rysunki_beta = new List<Drawing>(lista_rysunkow_elementy(party_faza, party_rysunki));

            //zespoly_faza_rysunki = zespoly_faza_rysunki_beta;
            //party_faza_rysunki = party_faza_rysunki_beta;



            //tworzy raporty
            raporty(party_faza, zespoly_faza);

            //tworzy NC
            nc_export(part_list);

            

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

  

            textBox5.Text = Convert.ToString(objekty.GetSize());
            textBox4.Text = Convert.ToString(objekty2.GetSize());

            DrawingHandler2 = new DrawingHandler();
            DrawingEnumerator rysunki = DrawingHandler2.GetDrawings();


           /* List<Drawing> drawings = new List<Drawing>();
            foreach (Drawing drawing in rysunki)
                drawings.Add(drawing);
            List<Drawing> assembly_drawings = new List<Drawing>(drawings.OfType<AssemblyDrawing>()); zespoly_rysunki = assembly_drawings;
            List<Drawing> part_drawings = new List<Drawing>(drawings.OfType<SinglePartDrawing>()); party_rysunki = part_drawings;
           */

            List<Tekla.Structures.Model.ModelObject> zespoly_wmodelu = new List<Tekla.Structures.Model.ModelObject>();
            foreach (Tekla.Structures.Model.ModelObject zespol in objekty)
                zespoly_wmodelu.Add(zespol); zespoly_faza = zespoly_wmodelu;

            List<Tekla.Structures.Model.ModelObject> party_wmodelu = new List<Tekla.Structures.Model.ModelObject>();
            foreach (Tekla.Structures.Model.ModelObject part in objekty2)
                party_wmodelu.Add(part); party_faza = party_wmodelu;

            ModelInfo info = MyModel.GetInfo();
            int faza_spr = info.CurrentPhase;

            if (faza_wpisana != Convert.ToString(faza_spr))
                MessageBox.Show("Zmień aktualną fazę!");
            else
                MessageBox.Show("Faza wczytana.");


            MyModel.CommitChanges();
        }

        private void druk_pdf(List<Drawing> zespoly, List<Drawing> elementy)
        {
            string p = "test";
            p = MyModel.GetProjectInfo().ProjectNumber;

            string pattern = "(\\W)";
            /*
            DrawingHandler DHP = new DrawingHandler();
            DPMPrinterAttributes dpmprint = new DPMPrinterAttributes();
            dpmprint.OutputType = DotPrintOutputType.PDF;
            dpmprint.ColorMode = DotPrintColor.BlackAndWhite;
            //zespoly
            foreach (Drawing pdf in zespoly)
            {
                dpmprint.OutputFileName = folder_dwa + "\\PDF\\Zespoły\\" + p +"_A_" + Regex.Replace(pdf.Mark, pattern, String.Empty) + " - " + pdf.Name + ".pdf";
                DHP.PrintDrawing(pdf, dpmprint);
            }

            elementy
            foreach (Drawing pdf in elementy)
            {
                dpmprint.OutputFileName = folder_dwa + "\\PDF\\Elementy\\" + p + "_W_" + Regex.Replace(pdf.Mark, pattern, String.Empty) +" - " + pdf.Name + ".pdf";
                DHP.PrintDrawing(pdf, dpmprint);
            }


            dpmprint.PrintToMultipleSheet = DotPrintToMultipleSheet.LeftToRightAndTopToBottom;
            dpmprint.OutputFileName = folder_dwa + "\\PDF\\Zespoły\\" + "FAZA " + faza + " - rysunki zespołów" + ".pdf";
            DHP.PrintDrawings(zespoly, dpmprint);

            
            dpmprint.PrintToMultipleSheet = DotPrintToMultipleSheet.LeftToRightAndTopToBottom;
            dpmprint.OutputFileName = folder_dwa + "\\PDF\\Elementy\\" + "FAZA " + faza + " - rysunki elementów" + ".pdf";
            DHP.PrintDrawings(elementy, dpmprint);
            */


            //foreach (Drawing pdf in zespoly)
            
                //DrawingHandler2.SetActiveDrawing(pdf);
                var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
                var printerLocation = installLoc + @"nt\bin\applications\Tekla\Model\DPMPrinter\DPMPrinterCommand.exe";
                var arg = string.Format("dpm:selected colormode:BlackAndWhite printer:PDF paper:Auto out:\"{0}.pdf\" settingsFile:D:\\pdf.PdfPrintOptions.xml", "D:\\test");
                Console.WriteLine("Args = '{0}'", arg);
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = printerLocation;
                process.StartInfo.Arguments = arg;
                process.StartInfo.WorkingDirectory = "D:\\";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.OutputDataReceived += (sender2, args) => Console.WriteLine("received output: {0}", args.Data);
                process.Start();
                process.BeginOutputReadLine();
                process.WaitForExit();
            


            /*foreach (Drawing pdf in elementy)
            {
                DrawingHandler2.SetActiveDrawing(pdf);
                var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
                var printerLocation = installLoc + @"nt\bin\applications\Tekla\Model\DPMPrinter\DPMPrinterCommand.exe";
                var arg = string.Format("printActive:true printer:pdf colormode:BlackAndWhite paper:Auto out:\"{0}.pdf\" settingsFile:D:\\pdf.PdfPrintOptions.xml", folder_dwa + "\\PDF\\Elementy\\" + p + "_W_" + Regex.Replace(pdf.Mark, pattern, String.Empty) + " - " + pdf.Name);
                Console.WriteLine("Args = '{0}'", arg);
                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = printerLocation;
                process.StartInfo.Arguments = arg;
                process.StartInfo.WorkingDirectory = folder_dwa + "\\PDF\\Elementy\\";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.OutputDataReceived += (sender2, args) => Console.WriteLine("received output: {0}", args.Data);
                process.Start();
                process.BeginOutputReadLine();
                process.WaitForExit();
            }*/



        }

        private void raporty(List<ModelObject> party, List<ModelObject> assembly)

        {

            ArrayList assembly2 = new ArrayList();
            foreach (Tekla.Structures.Model.ModelObject modelob in assembly)
                assembly2.Add(modelob);

            ArrayList party2 = new ArrayList();
            foreach (Tekla.Structures.Model.ModelObject modelob2 in party)
                party2.Add(modelob2);

            part_list = party2;

            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MS.Select(assembly2);
            string folder_listy = folder_dwa + "\\LISTY\\";
            Operation.CreateReportFromSelected("003 - Lista elementów wysyłkowych", folder_listy + "R-" + faza + "_003 - Lista elementów wysyłkowych" + " - FAZA " + faza + ".xsr", "", faza, rev);
            Operation.CreateReportFromSelected("003 - Lista elementów wysyłkowych.pdf", folder_listy + "R-" + faza + "_003 - Lista elementów wysyłkowych" + " - FAZA " + faza + ".pdf", "", faza, rev);
            Operation.CreateReportFromSelected("003 - Lista elementów wysyłkowych.xls", folder_listy + "R-" + faza + "_003 - Lista elementów wysyłkowych" + " - FAZA " + faza + ".xls", "", faza, rev);

            Operation.CreateReportFromSelected("002 - Elementy wysyłkowe z pozycjami", folder_listy + "R-" + faza + "_002 - Elementy wysyłkowe z pozycjami" + " - FAZA " + faza + ".xsr", "", faza, rev);
            Operation.CreateReportFromSelected("002 - Elementy wysyłkowe z pozycjami.pdf", folder_listy + "R-" + faza + "_002 - Elementy wysyłkowe z pozycjami" + " - FAZA " + faza + ".pdf", "", faza, rev);
            Operation.CreateReportFromSelected("002 - Elementy wysyłkowe z pozycjami.xls", folder_listy + "R-" + faza + "_002 - Elementy wysyłkowe z pozycjami" + " - FAZA " + faza + ".xls", "", faza, rev);

            MS.Select(party2);
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych", folder_listy + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".xsr", "", faza, rev);
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych.pdf", folder_listy + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".pdf", "", faza, rev);
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych.xls", folder_listy + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".xls", "", faza, rev);


        }

        private void nc_export(ArrayList nc)
        {
            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MS.Select(nc);
            string folder_nc = folder_dwa + "\\NC";
            string folder_dxf = MyModel.GetInfo().ModelPath + "\\NC files all\\";
            Operation.CreateNCFilesFromSelected("Blachy", folder_nc + "\\NC-Blachy\\");
            Operation.CreateNCFilesFromSelected("Profile", folder_nc + "\\NC-Profile\\");
            Operation.CreateNCFilesFromSelected("Blachy do DXF", folder_dxf);
            Operation.CreateMISFileFromSelected(Operation.MISExportTypeEnum.DSTV, folder_nc + "\\MIS_LIST_" + faza);
        }

        private void dwg_export_zespoly(List<Drawing> zespoly)
        {
            //DrawingHandler DHP = new DrawingHandler();
            //DHP.SetActiveDrawing(zespoly.ElementAt(1));
            var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
            var printerLocation = installLoc + @"nt\bin\applications\Tekla\Drawings\DwgExport\Dwg.exe";
            string dwgxportParams = "\"" + folder_dwa + "\\DWG\\Zespoły" + "\"";
            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = printerLocation;
            process.StartInfo.Arguments = "export outputDirectory=" + dwgxportParams + " " + "\"settingFile=bez osadzania";
            process.StartInfo.WorkingDirectory = folder_dwa+"\\DWG\\Zespoły";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();
        }

        private void dwg_export_elementy(List<Drawing> elementy)
        {

            //DrawingHandler DHP = new DrawingHandler();
            //DHP.SetActiveDrawing(elementy.ElementAt(1));
            var installLoc = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
            var printerLocation = installLoc + @"nt\bin\applications\Tekla\Drawings\DwgExport\Dwg.exe";
            string dwgxportParams = "\"" + folder_dwa + "\\DWG\\Elementy" + "\"";
            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = printerLocation;
            process.StartInfo.Arguments = "export outputDirectory=" + dwgxportParams + " " + "\"settingFile=bez osadzania";
            process.StartInfo.WorkingDirectory = folder_dwa + "\\DWG\\Elementy";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();
        }


        private static List<Drawing> lista_rysunkow_zespoly(List<Tekla.Structures.Model.ModelObject> zespoly, List<Drawing> rysunki)
        {
            string pattern = "(\\W)";
            string p = "tak";
            List<string> zespoly_id = new List<string>();
            foreach (Tekla.Structures.Model.ModelObject id in zespoly)
            {
                (id as Tekla.Structures.Model.Assembly).GetReportProperty("ASSEMBLY_POS", ref p);
                zespoly_id.Add(Convert.ToString(p));
            }

            List<Drawing> lista_zespolow = new List<Drawing>();
            foreach (Drawing id in rysunki)
                if (zespoly_id.Contains(Regex.Replace(id.Mark, pattern, String.Empty)))
                {
                    lista_zespolow.Add(id);
                }
            return lista_zespolow;
        }

        private static List<Drawing> lista_rysunkow_elementy(List<Tekla.Structures.Model.ModelObject> party, List<Drawing> rysunki)
        {
            string pattern = "(\\W)";
            string p = "tak";
            List<string> part_mark_model = new List<string>();
            foreach (Tekla.Structures.Model.ModelObject part in party)
            {
                (part as Tekla.Structures.Model.Part).GetReportProperty("PART_POS", ref p);
                part_mark_model.Add(Convert.ToString(p));
            }

            List<Drawing> lista_part = new List<Drawing>();
            foreach (Drawing id in rysunki)
                if (part_mark_model.Contains(Regex.Replace(id.Mark, pattern, String.Empty)))
                {
                    lista_part.Add(id);
                }

            return lista_part;
        }

        private void lista_materialowa(string faza)
        {
            ObjectFilterExpressions.Type obj1 = new ObjectFilterExpressions.Type();
            NumericConstantFilterExpression type2 = new NumericConstantFilterExpression(Tekla.Structures.TeklaStructuresDatabaseTypeEnum.PART);
            string faza_wpisana = faza;
            AssemblyFilterExpressions.Phase exp1 = new AssemblyFilterExpressions.Phase();
            StringConstantFilterExpression faza_filter = new StringConstantFilterExpression(faza_wpisana);
            BinaryFilterExpression bexp = new BinaryFilterExpression(exp1, StringOperatorType.IS_EQUAL, faza_filter);
            BinaryFilterExpression bexp3 = new BinaryFilterExpression(obj1, NumericOperatorType.IS_EQUAL, type2);

            BinaryFilterExpressionCollection bexpcoll2 = new BinaryFilterExpressionCollection();
            bexpcoll2.Add(new BinaryFilterExpressionItem(bexp, BinaryFilterOperatorType.BOOLEAN_AND));
            bexpcoll2.Add(new BinaryFilterExpressionItem(bexp3, BinaryFilterOperatorType.BOOLEAN_AND));

            ModelObjectEnumerator objekty2 = MyModel.GetModelObjectSelector().GetObjectsByFilter(bexpcoll2);

            textBox4.Text = Convert.ToString(objekty2.GetSize());

            ArrayList party_wmodelu = new ArrayList();
            foreach (Tekla.Structures.Model.ModelObject part in objekty2)
                party_wmodelu.Add(part);

            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MS.Select(party_wmodelu);

            MS.Select(party_wmodelu);
            Directory.CreateDirectory(location + "\\" + faza + " - " + nazwa + "\\");
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych", location + "\\" + faza + " - " + nazwa + "\\" + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".xsr", "", faza, rev);
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych.pdf", location + "\\" + faza + " - " + nazwa + "\\" + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".pdf", "", faza, rev);
            Operation.CreateReportFromSelected("001 - Raport elementów konstrukcji stalowych.xls", location + "\\" + faza + " - " + nazwa + "\\" + "R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".xls", "", faza, rev);

            System.Threading.Thread.Sleep(3000);

            File.Delete(location + "\\" + faza + " - " + nazwa + "\\" +"R-" + faza + "_001 - Raport elementów konstrukcji stalowych" + " - FAZA " + faza + ".pdf" + ".dpm");

            
        }




        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            location = textBox1.Text;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            faza = textBox2.Text;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            rev = textBox3.Text;
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            nazwa = textBox6.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dwg_export_zespoly(zespoly_faza_rysunki);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dwg_export_elementy(party_faza_rysunki);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string folder_listy = folder_dwa + "\\LISTY\\";
            Operation.CreateReportFromSelected("004 - Lista rysunków", folder_listy + "R-" + faza + "_004 - Lista rysunków" + " - FAZA " + faza + ".xsr", "", faza, rev);
            Operation.CreateReportFromSelected("004 - Lista rysunków.pdf", folder_listy + "R-" + faza + "_004 - Lista rysunków" + " - FAZA " + faza + ".pdf", "", faza, rev);
            Operation.CreateReportFromSelected("004 - Lista rysunków.xls", folder_listy + "R-" + faza + "_004 - Lista rysunków" + " - FAZA " + faza + ".xls", "", faza, rev);
        }

        private void button7_Click(object sender, EventArgs e)
        {

            druk_pdf(zespoly_faza_rysunki, party_faza_rysunki);

        }

        private void button8_Click(object sender, EventArgs e)
        {
            lista_materialowa(faza);
        }

        private void button9_Click(object sender, EventArgs e)
        {

            string folder_listy = folder_dwa + "\\";
            Operation.CreateReportFromSelected("006 - Raport zmian", folder_listy + "R-" + faza + "_006 - Raport zmian" + " - FAZA " + faza + ".xsr", "", faza, rev);
        }

    }
   
}
