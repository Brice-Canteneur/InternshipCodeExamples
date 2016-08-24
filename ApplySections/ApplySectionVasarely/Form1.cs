using System;
using System.Windows.Forms;
using RobotOM;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ApplySectionVasarely
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"EXCELPATH...\excelFile.xlsx", 0, false);
            Excel.Sheets xlSheets = xlWorkbook.Worksheets;
            string currentSheet = "Sheet1";
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlSheets.get_Item(currentSheet);
            IRobotApplication robotApp = new RobotApplication();
            IRobotLabelServer lab_serv = robotApp.Project.Structure.Labels;

            progressBar1.Value = 0;
            progressBar1.Maximum = 151;
            progressBar1.Minimum = 0;
            progressBar1.Step = 1;


            for (int i = 1; i < 152; i++)
            {
                string secName = xlWorksheet.Cells[i+1, 1].Value.ToString();
                double secAx = double.Parse(xlWorksheet.Cells[i+1, 2].Value.ToString());
                double secIx = double.Parse(xlWorksheet.Cells[i + 1, 3].Value.ToString());
                double secIy = double.Parse(xlWorksheet.Cells[i + 1, 4].Value.ToString());
                double secIz = double.Parse(xlWorksheet.Cells[i + 1, 5].Value.ToString());
                double secVy = double.Parse(xlWorksheet.Cells[i + 1, 6].Value.ToString());
                double secVpy = double.Parse(xlWorksheet.Cells[i + 1, 7].Value.ToString());
                double secVz = double.Parse(xlWorksheet.Cells[i + 1, 8].Value.ToString());
                double secVpz = double.Parse(xlWorksheet.Cells[i + 1, 9].Value.ToString());


                IRobotLabel sec = lab_serv.Create(IRobotLabelType.I_LT_BAR_SECTION, secName);
                IRobotBarSectionData data = sec.Data;
                data.Type = IRobotBarSectionType.I_BST_STANDARD;
                data.ShapeType = IRobotBarSectionShapeType.I_BSST_UNKNOWN;

                data.SetValue(IRobotBarSectionDataValue.I_BSDV_AX, secAx);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_IX, secIx);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_IY, secIy);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_IZ, secIz);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_VY, secVy);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_VPY, secVpy);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_VZ, secVz);
                data.SetValue(IRobotBarSectionDataValue.I_BSDV_VPZ, secVpz);

                lab_serv.Store(sec);

                IRobotBar bar = (IRobotBar)robotApp.Project.Structure.Bars.Get(i);

                bar.SetLabel(IRobotLabelType.I_LT_BAR_SECTION, secName);

                progressBar1.PerformStep();

            }

            Marshal.FinalReleaseComObject(xlWorksheet);
            Marshal.FinalReleaseComObject(xlSheets);
            xlWorkbook.Close(false);
            Marshal.FinalReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);

            MessageBox.Show("All the profiles have been applied!", "Work done !", MessageBoxButtons.OK);
        }
    }
}
