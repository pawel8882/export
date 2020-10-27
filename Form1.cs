using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using System.Collections;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;

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

            bexp.IsEnable = true;
            //Filter filtr1 = new Filter(bexp);
            //ModelObjectEnumerator.AutoFetch = true;
            ModelObjectEnumerator objekty = MyModel.GetModelObjectSelector().GetObjectsByFilter(bexpcoll);
            ModelObjectEnumerator objekty2 = MyModel.GetModelObjectSelector().GetObjectsByFilter(bexpcoll2);
            textBox5.Text = Convert.ToString(objekty.GetSize());
            textBox4.Text = Convert.ToString(objekty2.GetSize());




            MyModel.CommitChanges();
            //Filter faza = new Filter();
        }
    }
}
