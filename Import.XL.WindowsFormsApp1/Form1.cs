using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Import.XL.WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var DT = new DataTable();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (var stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read))
                {
                    using (var RD = ExcelReaderFactory.CreateReader(stream))
                    {
                        var I = 0;
                        var fromRow = 0;
                        var fromCol = 0;
                        var CNFG = new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                FilterRow = rowReader => fromRow <= ++I - 1,
                                FilterColumn = (rowReader, colIndex) => fromCol <= colIndex,
                                UseHeaderRow = true,
                            }
                        };

                        var RES = RD.AsDataSet(CNFG);
                        DT = RES.Tables[0];
                    }
                }
            }
            dataGridView1.DataSource = DT;
        }
    }
}
