using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;

namespace avaliar
{

       
    public partial class Form1 : Form
    {
        string Date = DateTime.Now.ToShortDateString();
        string dia = DateTime.Now.Day.ToString();
        string mes = DateTime.Now.Month.ToString();
        string ano = DateTime.Now.Year.ToString();
        

        public Form1()
        {
            InitializeComponent();

          
        }
        public void enviarExcel()
        {
            Excel.Application XlApp;
            Excel.Workbook XlWorkBook;
            Excel.Worksheet XlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            //cria planilia temporaria
            XlApp = new Excel.Application();
            XlWorkBook = XlApp.Workbooks.Add(misValue);
            XlWorkSheet = (Excel.Worksheet)XlWorkBook.Worksheets.get_Item(1);

            XlWorkSheet.Cells[1, 1] = "Dados da avaliação";
            XlWorkSheet.Cells[2, 1] = "Muito Ruim";
            XlWorkSheet.Cells[3, 1] = "Ruim";
            XlWorkSheet.Cells[4, 1] = "Médio";
            XlWorkSheet.Cells[5, 1] = "Bom";
            XlWorkSheet.Cells[6, 1] = "Muito Bom";
            XlWorkSheet.Cells[7, 1] = "Ótimo";
            XlWorkSheet.Cells[8, 1] = "QTD Total votos";
            XlWorkSheet.Cells[9, 1] = "Média";
            XlWorkSheet.Cells[1, 2] = "QTD Votos";
            XlWorkSheet.Cells[2, 2] = mtruim.Text;
            XlWorkSheet.Cells[3, 2] = label1.Text;
            XlWorkSheet.Cells[4, 2] = label2.Text;
            XlWorkSheet.Cells[5, 2] = label3.Text;
            XlWorkSheet.Cells[6, 2] = mtbom.Text;
            XlWorkSheet.Cells[7, 2] = otimo.Text;
            XlWorkSheet.Cells[8, 2] = label8.Text;
            XlWorkSheet.Cells[9, 2] = label5.Text;
            XlWorkSheet.Cells[1, 3] = "Relatorio NTP";
            XlWorkSheet.Cells[2, 3] = "Valor NTP";
            XlWorkSheet.Cells[2, 4] = label16.Text;
            XlWorkSheet.Cells[6, 3] = "% de avaliação";
            XlWorkSheet.Cells[7, 3] = "Baixo";
            XlWorkSheet.Cells[8, 3] = "medio";
            XlWorkSheet.Cells[9, 3] = "alto";
            XlWorkSheet.Cells[7, 4] = label17.Text;
            XlWorkSheet.Cells[8, 4] = label18.Text;
            XlWorkSheet.Cells[9, 4] = label19.Text;
            XlWorkSheet.Cells[1, 5] = "Estrelas";
            XlWorkSheet.Cells[2, 5] = label6.Text;
            XlWorkSheet.Cells[3, 5] = label7.Text;
            XlWorkSheet.Cells[1, 6] = label4.Text;
            XlWorkSheet.Cells[11, 11] = "Os graficos estao sobreposto, arraste para separalos.";
            //erro grafico 1 n aparece e grafico 2 sim como resolver.
            //grafico 2
            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)XlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(100, 180, 500, 300);
            Excel.Chart chartPage = myChart.Chart;


            chartRange = XlWorkSheet.get_Range("A2", "B7");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
            //fim do grafico 1
            //grafico1
            Excel.Range chartRange1;

            Excel.ChartObjects xlCharts1 = (Excel.ChartObjects)XlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart1 = (Excel.ChartObject)xlCharts1.Add(100, 180, 500, 300);
            Excel.Chart chartPage1 = myChart1.Chart;


            chartRange1 = XlWorkSheet.get_Range("C7", "d9");
            chartPage1.SetSourceData(chartRange1, misValue);
            chartPage1.ChartType = Excel.XlChartType.xlColumnClustered;


           

            XlWorkBook.SaveAs(dia +"0" + mes+ ano + ".xls",
            Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            XlWorkBook.Close(true, misValue, misValue);
            XlApp.Quit();
            string folder1 = @"C:/Relatorios/";
           
          
            if (!Directory.Exists(folder1))
            {

                //Criamos um com o nome folder
                Directory.CreateDirectory(folder1);

            }
         
            MessageBox.Show(folder1+ dia + "0" + mes + ano + "xls"+" "+" caso não esteja la verifique em documentos");
        }


        private void ruim_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(label1.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

            label1.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label4.Text = Convert.ToString(Date);
          
        }

        private void salvarToolStripMenuItem_Click(object sender, EventArgs e)
        {
             

           

           

        }

        private void salvarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string folder = @"C:/Relatorios";
            ////declarando a variável do tipo StreamWriter para
            //abrir ou criar um arquivo para escrita
            StreamWriter x;
            if (!Directory.Exists(folder))
            {

                //Criamos um com o nome folder
                Directory.CreateDirectory(folder);

            }
            //Colocando o endereço físico (caminho do arquivo texto)
            string Caminho = folder+"/"+dia +"0" + mes+ ano + ".docs";

            if (MessageBox.Show("Deseja Salvar?", "Atenção",
           MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation,
         MessageBoxDefaultButton.Button2) == DialogResult.No)
            {
             

            }
            else
            {
                //usando o metodo e abrindo o arquivo texto

                x = File.AppendText(Caminho);


                //continuando escrevendo neste arquivo
                //escrevendo a partir da ultima linha
                calcular();
                x.WriteLine("Sistema de Avaliação de estabelecimento");
                x.WriteLine("Relatorio Data: " + label4.Text);
                x.WriteLine("");
                x.WriteLine("Index:");
                x.WriteLine("");
                x.WriteLine("Muito Ruim = 0");
                x.WriteLine("Ruim = 3");
                x.WriteLine("Medio = 5");
                x.WriteLine("Bom = 7");
                x.WriteLine("Muito Bom = 9");
                x.WriteLine("Otimo = 10");
                x.WriteLine("");
                x.WriteLine("Descrição________quantidade de votos");
                x.WriteLine("quantidade total de votos: " + label8.Text);
                x.WriteLine("");
                x.WriteLine("Muito Ruim:__________ " + mtruim.Text);
                x.WriteLine("Ruim:________________ " + label1.Text);
                x.WriteLine("Medio:_______________ " + label2.Text);
                x.WriteLine("Bom:_________________ " + label3.Text);
                x.WriteLine("Muito Bom:___________ " + mtbom.Text);
                x.WriteLine("Otimo:_______________ " + otimo.Text);
                x.WriteLine("");
                x.WriteLine("Nota final do estabelecimento(média): " + label5.Text);
                x.WriteLine("");
                x.WriteLine("isso equivale a um estabelecimento: " + label6.Text);
                x.WriteLine("Estrelas: " + label7.Text);
                x.WriteLine("");
                x.WriteLine("");
                x.WriteLine("Relatorio NTP");
                x.WriteLine("");
                x.WriteLine("valor NTP= "+ label16.Text);
                x.WriteLine("Seu nivel de NTP esta:" + label20.Text);
                x.WriteLine("");
                x.WriteLine("Porcentagens de avaliação:");
                x.WriteLine("");
                x.WriteLine("(Baixo)de 0 a 6: " + label17.Text);
                x.WriteLine("(médio)de 7 a 8: " + label18.Text);
                x.WriteLine("(alto)de 9 a 10: " + label19.Text);
                x.WriteLine("");
                x.WriteLine("Indicativo: "+label21.Text);
                x.WriteLine("");
                x.WriteLine("");
                x.WriteLine("__________________________________________________");
                x.WriteLine("assinatura do dia");
                x.WriteLine();
                enviarExcel();
                x.Close();
                MessageBox.Show("Salvo com sucesso \n\n diretorio: "+Caminho);
            }
        }

        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(label2.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

            label2.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(label3.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

            label3.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void bmtruim_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(mtruim.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

            mtruim.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void bmtbom_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(mtbom.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

           mtbom.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void botimo_Click(object sender, EventArgs e)
        {
            int num1 = Convert.ToInt32(otimo.Text);
            int resultado;
            //adicionar 1 todas vez que clicar no botao

            resultado = num1 + 1;

            otimo.Text = Convert.ToString(resultado);
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        public void calcular()
        {
            int valor1 = Convert.ToInt32(mtruim.Text);
            int valor2 = Convert.ToInt32(label1.Text);
            int valor3 = Convert.ToInt32(label2.Text);
            int valor4 = Convert.ToInt32(label3.Text);
            int valor5 = Convert.ToInt32(mtbom.Text);
            int valor6 = Convert.ToInt32(otimo.Text);
            int valorfinal;
            int valor11;
            int valor22;
            int valor33;
            int valor44;
            int valor55;
            int valor66;
            int resultado1;
            int ntp;
            float vbaixo;
            float vmedio;
            float valto;
            ntp = (valor5 + valor6) - (valor1 + valor2 + valor3);
            valorfinal = valor1 + valor2 + valor3 + valor4 + valor5 + valor6;
           


            valor11 = valor1 * 0;
            valor22 = valor2 * 3;
            valor33 = valor3 * 5;
            valor44 = valor4 * 7;
            valor55 = valor5 * 9;
            valor66 = valor6 * 10;

            int notafinal = valor11 + valor22 + valor33 + valor44 + valor55 + valor66;

            if (valorfinal == 0)
            {
                label5.Text = "Não houve votação";
                label6.Text = "Não houve votação";
                label7.Text = "Não houve votação";
                label8.Text = "Não houve votação";
                label16.Text = "Não houve votação";
                label17.Text = "Não houve votação";
                label18.Text = "Não houve votação";
                label19.Text = "Não houve votação";
                label20.Text = "Não houve votação";
                label21.Text = "Não houve votação";
            }
            else
            {
                
                vbaixo = (float)(valor1 + valor2 + valor3) / valorfinal;
                vmedio = (float)valor4 / valorfinal;
                valto = (float)(valor5 + valor6) / valorfinal;

                float resultadob = vbaixo * 100;
                float resultadom = vmedio * 100;
                float resultadoa = valto * 100;
                label8.Text = Convert.ToString(valorfinal);
                
              
                label16.Text = Convert.ToString(ntp);
                label17.Text = Convert.ToString(resultadob+"%");
                label18.Text = Convert.ToString(resultadom+"%");
                label19.Text = Convert.ToString(resultadoa+"%");
                if (ntp < 0)
                {
                    label20.Text = "Negativo, Significa que ha muitas avaliações negativas. isso pode afetar os negócios.";
                    
                }
                else
                {
                    label20.Text = "Positivo, Sgnifica que ha muitas avalições positivas.";
                }
                if (ntp == 0)
                {
                    label20.Text = "Neutro, houve a mesma quantidade de avaliações positivas e negativas, estar no neutro não é bom e nem ruim.";
                }
                if (resultadob>resultadom&&resultadob>resultadoa) {
                    label21.Text = "(Baixo)seu estabelecimento está com pessima avaliação.";
                }
                if (resultadom > resultadob && resultadom > resultadoa)
                {
                    label21.Text = "(medio)seu estabelecimento está com uma avaliação mediana.";
                }
                if (resultadoa > resultadob && resultadoa > resultadom)
                {
                    label21.Text = "(alto)seu estabelecimento está com uma otima avalição.";
                }
                if (resultadoa== resultadob && resultadoa == resultadom)
                {
                    label21.Text = "(Empate)seu estabelecimento teve avaliaçoes iguais, significa equivalencia.";
                }
                if (resultadoa == resultadob && resultadoa > resultadom)
                {
                    label21.Text = "(Neutro)seu estabelecimento teve mais avaliações tanto positiva quanto negativa.";
                }
                if (resultadob == resultadom && resultadob > resultadoa)
                {
                    label21.Text = "(Baixo para médio)seu estabelecimento está negativo para mediano.";
                }
                if (resultadoa == resultadom && resultadoa > resultadob)
                {
                    label21.Text = "(medio para alto)seu estabelecimento está mediano para positivo.";
                }
                resultado1 = notafinal / valorfinal;
                label5.Text = Convert.ToString(resultado1);
                if (resultado1 == 0)
                {
                    label6.Text = "0 estrelas";
                    label7.Text = "☆☆☆☆☆";

                }
                if (resultado1 == 1)
                {
                    label6.Text = "0.5 estrelas";
                    label7.Text = "✪☆☆☆☆";

                }
                if (resultado1 == 2)
                {
                    label6.Text = "1 estrelas";
                    label7.Text = "★☆☆☆☆";

                }
                if (resultado1 == 3)
                {
                    label6.Text = "1.5 estrelas";
                    label7.Text = "★✪☆☆☆";
                }

                    if (resultado1 == 4)
                    {
                        label6.Text = "2 estrelas";
                        label7.Text = "★★☆☆☆";

                    }
                    if (resultado1 == 5)
                    {
                        label6.Text = "2.5 estrelas";
                        label7.Text = "★★✪☆☆";

                    }
                    if (resultado1 == 6)
                    {
                        label6.Text = "3 estrelas";
                        label7.Text = "★★★☆☆";

                    }
                    if (resultado1 == 7)
                    {
                        label6.Text = "3.5 estrelas";
                        label7.Text = "★★★✪☆";

                    }
                    if (resultado1 == 8)
                    {
                        label6.Text = "4 estrelas";
                        label7.Text = "★★★★☆";

                    }
                    if (resultado1 == 9)
                    {
                        label6.Text = "4.5 estrelas";
                        label7.Text = "★★★★✪";


                    }
                    if (resultado1 == 10)
                    {
                        label6.Text = "5 estrelas";
                        label7.Text = "★★★★★";

                    }
                }

            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //talves criar uma senha
      

        
          

        }

        private void tutorialToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
         


        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
    }

}
