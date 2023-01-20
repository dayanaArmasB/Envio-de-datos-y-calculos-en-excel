using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace tarea
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_ejecutar_Click(object sender, EventArgs e)
        {
            var excelApp = new Excel.Application();

           //cantidad de celdas
            const int max_cel = 19;
            
            Random x = new Random();
            List<int> llegadas = new List<int>();

            List<int> servicios = new List<int>();


            for (int i = 0; i <= max_cel; i++)
            {

                llegadas.Add(x.Next(1, 21));
                servicios.Add(x.Next(1, 21));
            }


            excelApp.Visible= true;
            excelApp.Workbooks.Add();

            Excel._Worksheet Hoja = (Excel.Worksheet)excelApp.ActiveSheet;
            Hoja.Cells[1, "A"] = "CLIENTES";
            Hoja.Cells[1, "B"] = "LLEGADAS";
            Hoja.Cells[1, "C"] = "SERVICIOS";

            var fila = 1;
            for (int n = 0; n <= max_cel ; n++)
            {
                fila++;
                Hoja.Cells[fila, "A"] = n+1;
                Hoja.Cells[fila, "B"] = llegadas[n];
                Hoja.Cells[fila, "C"] = servicios[n];
                Hoja.Columns[1].Autofit();
                Hoja.Columns[2].Autofit();
                Hoja.Columns[3].Autofit();

                Hoja.Cells[22, "A"] = "suma";
                Hoja.Cells[23, "A"] = "promedio";
                Hoja.Cells[24, "A"] = "desviacion estandar";
                Hoja.Cells[25, "A"] = "viaranza";
                Hoja.Cells[26, "A"] = "asimetria";
                Hoja.Cells[27, "A"] = "curtosis";
                Hoja.Cells[22, "B"].Formula = "=SUM(B2:B21";
                Hoja.Cells[23, "B"].Formula = "=AVERAGE(B2:B21";
                Hoja.Cells[24, "B"].Formula = "=STDEV.P(B2:B21";
                Hoja.Cells[25, "B"].Formula = "=VAR.P(B2:B21";
                Hoja.Cells[26, "B"].Formula = "=SKEW.P(B2:B21";
                Hoja.Cells[27, "B"].Formula = "=KURT(B2:B21)";
                Hoja.Cells[22, "C"].Formula = "=SUM(C2:C21";
                Hoja.Cells[23, "C"].Formula = "=AVERAGE(C2:C21";
                Hoja.Cells[24, "C"].Formula = "=STDEV.P(C2:C21)";
                Hoja.Cells[25, "C"].Formula = "=VAR.P(C2:C21";
                Hoja.Cells[26, "C"].Formula = "=SKEW.P(C2:C21";
                Hoja.Cells[27, "C"].Formula = "=KURT(C2:C21)";
            }
        }
    }
}
