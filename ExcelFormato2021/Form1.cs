using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using xls = Microsoft.Office.Interop.Excel;

namespace ExcelFormato2021
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        xls.Application a = new xls.Application();
        int i = 6;
        string num;
        string matricula;
        string apellidop;
        string apellidom;
        string nombre;
        string especialidad;
        string semestre;
        string servicio;
        string practicas;
        string residencias;
        string certificaciones;
        string toefl;

        int totalprimersemestre = 0;
        int totalsegundosemestre = 0;
        int totaltecermestre = 0;
        int totalcuartosemestre = 0;
        int totalquintosemestre = 0;
        int totalsextosemestre = 0;
        int totalseptimosemestre = 0;
        int totaloctavosemestre = 0;
        int totalnovenosemestre = 0;
        int totalinformatica = 0;
        int totalmecanica = 0;
        int totalelectronica = 0;
        int totalindustrial = 0;
        int totalgestion = 0;
        int totalenergias = 0;
        int totalservicio = 0;
        int totalpracticas = 0;
        int totalresidencias = 0;
        int totalcertificaciones = 0;
        int totaltoefl = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            a.Workbooks.Open(Application.StartupPath + @"\formato2021");
            while (a.ActiveWorkbook.ActiveSheet.Cells(i, 1).Value != null)
            {
                i++;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView2.Items.Clear();
            int x = 6;
            while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
            {
                num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                apellidop = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                apellidom = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                switch (especialidad)
                {
                    case "Informatica":
                        totalinformatica++;
                        break;
                    case "Mecanica":
                        totalmecanica++;
                        break;
                    case "Electronica":
                        totalelectronica++;
                        break;
                    case "Industrial":
                        totalindustrial++;
                        break;
                    case "Gestion empresarial":
                        totalgestion++;
                        break;
                    case "Energias renovables":
                        totalenergias++;
                        break;
                }
                semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                switch (semestre)
                {
                    case "1":
                        totalprimersemestre++;
                        break;
                    case "2":
                        totalsegundosemestre++;
                        break;
                    case "3":
                        totaltecermestre++;
                        break;
                    case "4":
                        totalcuartosemestre++;
                        break;
                    case "5":
                        totalquintosemestre++;
                        break;
                    case "6":
                        totalsextosemestre++;
                        break;
                    case "7":
                        totalseptimosemestre++;
                        break;
                    case "8":
                        totaloctavosemestre++;
                        break;
                    case "9":
                        totalnovenosemestre++;
                        break;
                }
                servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                if(servicio == "Si")
                {
                    totalservicio++;
                }
                practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                if(practicas == "Si")
                {
                    totalpracticas++;
                }
                residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                if(residencias == "Si")
                {
                    totalresidencias++;
                }
                certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                if(certificaciones == "Si")
                {
                    totalcertificaciones++;
                }
                toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();
                if(toefl == "Si")
                {
                    totaltoefl++;
                }

                ListViewItem lista = new ListViewItem(num);

                lista.SubItems.Add(matricula);
                lista.SubItems.Add(apellidop);
                lista.SubItems.Add(apellidom);
                lista.SubItems.Add(nombre);
                lista.SubItems.Add(especialidad);
                lista.SubItems.Add(semestre);
                lista.SubItems.Add(servicio);
                lista.SubItems.Add(practicas);
                lista.SubItems.Add(residencias);
                lista.SubItems.Add(certificaciones);
                lista.SubItems.Add(toefl);
                listView2.Items.Add(lista);
                x++;
            }
            decimal total = x-6;
            lblAlumnosInformatica.Text = (((totalinformatica / total) * 100) + "%");
            lblAlumnosMecanica.Text = (((totalmecanica / total) * 100) + "%");
            lblAlumnosElectronica.Text = (((totalelectronica / total) * 100) + "%");
            lblAlumnosGestion.Text = (((totalgestion / total) * 100) + "%");
            lblAlumnosEnergias.Text = (((totalenergias / total) * 100) + "%");
            AlumnosPrimerSemestre.Text = (((totalprimersemestre / total) * 100) + "%");
            lblAlumnosSegundoSemestre.Text = (((totalsegundosemestre / total) * 100) + "%");
            lblAlumnosTercerSemestre.Text = (((totaltecermestre / total) * 100) + "%");
            lblAlumnosCuartoSemestre.Text = (((totalcuartosemestre / total) * 100) + "%");
            lblAlumnosQuintoSemestre.Text = (((totalquintosemestre / total) * 100) + "%");
            lblAlumnosSextoSemestre.Text = (((totalsextosemestre / total) * 100) + "%");
            lblAlumnosSeptimoSemestre.Text = (((totalseptimosemestre / total) * 100) + "%");
            lblAlumnosOctavoSemestre.Text = (((totaloctavosemestre / total) * 100) + "%");
            lblAlumnosNovenoSemestre.Text = (((totalnovenosemestre / total) * 100) + "%");
            lblServicioSocial.Text = (((totalservicio / total) * 100) + "%");
            lblPracticasProfesionales.Text = (((totalpracticas / total) * 100) + "%");
            lblResidencias.Text = (((totalresidencias / total) * 100) + "%");
            lblCertificaciones.Text = (((totalcertificaciones / total) * 100) + "%");
            lblToefl.Text = (((totaltoefl / total) * 100) + "%");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {


                matricula = txtMatricula.Text;
                apellidop = txtApellidoPaterno.Text;
                apellidom = txtApellidoMaterno.Text;
                nombre = txtNombre.Text;
                especialidad = cmbxEspecialidad.Text;
                semestre = cmbxSemestre.Text;
                servicio = cmbxServicio.Text;
                practicas = cmbxPracticas.Text;
                residencias = cmbxResidencias.Text;
                certificaciones = txtCertificaciones.Text;
                toefl = cmbxToefl.Text;

                a.ActiveWorkbook.Worksheets[1].Cells(i, 1).Value = i - 5;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 2).Value = matricula;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 3).Value = apellidop;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 4).Value = apellidom;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 5).Value = nombre;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 6).Value = especialidad;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 7).Value = semestre;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 8).Value = servicio;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 9).Value = practicas;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 10).Value = residencias;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 11).Value = certificaciones;
                a.ActiveWorkbook.Worksheets[1].Cells(i, 12).Value = toefl;

                a.ActiveWorkbook.Save();
                i++;
                MessageBox.Show("Se agregaron los datos al archivo de Excel.", "Lectura y Escritura", MessageBoxButtons.OK);
                txtMatricula.Clear();
                txtApellidoPaterno.Clear();
                txtApellidoMaterno.Clear();
                txtNombre.Clear();
                txtCertificaciones.Clear();
            }
            catch
            {
                MessageBox.Show("Introdduzca los datos correctamente");
            }

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void btnBuscarEspecialidad_Click(object sender, EventArgs e)
        {
            string especialidadbuscada = cmbxEspecialidad.Text;
            if (especialidadbuscada != "")
            {
                listView2.Items.Clear();
                int x = 6;
                while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
                {
                    especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                    if (especialidad == especialidadbuscada)
                    {
                        num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                        matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                        apellidop = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                        apellidom = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                        nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                        semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                        servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                        practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                        residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                        certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                        toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();

                        ListViewItem lista = new ListViewItem(num);

                        lista.SubItems.Add(matricula);
                        lista.SubItems.Add(apellidop);
                        lista.SubItems.Add(apellidom);
                        lista.SubItems.Add(nombre);
                        lista.SubItems.Add(especialidad);
                        lista.SubItems.Add(semestre);
                        lista.SubItems.Add(servicio);
                        lista.SubItems.Add(practicas);
                        lista.SubItems.Add(residencias);
                        lista.SubItems.Add(certificaciones);
                        lista.SubItems.Add(toefl);
                        listView2.Items.Add(lista);
                    }
                    x++;
                }
            }
        }

        private void btnBuscarSemestre_Click(object sender, EventArgs e)
        {
            string semestrebuscado = cmbxSemestre.Text;
            if (semestrebuscado != "")
            {
                listView2.Items.Clear();
                int x = 6;
                while (a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value != null)
                {
                    semestre = a.ActiveWorkbook.ActiveSheet.Cells(x, 7).Value.ToString();
                    if (semestre == semestrebuscado)
                    {
                        num = a.ActiveWorkbook.ActiveSheet.Cells(x, 1).Value.ToString();
                        matricula = a.ActiveWorkbook.ActiveSheet.Cells(x, 2).Value.ToString();
                        apellidop = a.ActiveWorkbook.ActiveSheet.Cells(x, 3).Value.ToString();
                        apellidom = a.ActiveWorkbook.ActiveSheet.Cells(x, 4).Value.ToString();
                        nombre = a.ActiveWorkbook.ActiveSheet.Cells(x, 5).Value.ToString();
                        especialidad = a.ActiveWorkbook.ActiveSheet.Cells(x, 6).Value.ToString();
                        servicio = a.ActiveWorkbook.ActiveSheet.Cells(x, 8).Value.ToString();
                        practicas = a.ActiveWorkbook.ActiveSheet.Cells(x, 9).Value.ToString();
                        residencias = a.ActiveWorkbook.ActiveSheet.Cells(x, 10).Value.ToString();
                        certificaciones = a.ActiveWorkbook.ActiveSheet.Cells(x, 11).Value.ToString();
                        toefl = a.ActiveWorkbook.ActiveSheet.Cells(x, 12).Value.ToString();

                        ListViewItem lista = new ListViewItem(num);

                        lista.SubItems.Add(matricula);
                        lista.SubItems.Add(apellidop);
                        lista.SubItems.Add(apellidom);
                        lista.SubItems.Add(nombre);
                        lista.SubItems.Add(especialidad);
                        lista.SubItems.Add(semestre);
                        lista.SubItems.Add(servicio);
                        lista.SubItems.Add(practicas);
                        lista.SubItems.Add(residencias);
                        lista.SubItems.Add(certificaciones);
                        lista.SubItems.Add(toefl);
                        listView2.Items.Add(lista);
                    }
                    x++;
                }
            }
        }
    }
}
