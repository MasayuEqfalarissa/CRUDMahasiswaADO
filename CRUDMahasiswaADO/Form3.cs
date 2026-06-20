using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class Form3 : Form
    {
        static string ConnectionString = "Data Source=MASAYUU-DN5OL92\\MASAYU;Initial Catalog=DBAkademikADO;Integrated Security=True";
        SqlConnection conn;
        SqlDataAdapter da;
        DataTable dtMahasiswa;

        DAL dbLogic = new DAL();

        public void simpanLog(string message)
        {
            dbLogic.InsertLog(message);
        }

        CrystalReport2 classMahasiswa = new CrystalReport2();

        string Prodi { get; set; }
        DateTime TglMasuk { get; set; }

        // PERBAIKAN: Tambahkan parameter penerima di dalam kurung
        public Form3(string prodi, DateTime tglMasuk)
        {
            InitializeComponent();
           
            prodi = Prodi;
            tglMasuk = TglMasuk;

            try
            {
                DataTable dtMahasiswa = dbLogic.getDataRekap(Prodi, TglMasuk);

                classMahasiswa.SetDataSource(dtMahasiswa);
                crystalReportViewer1.ReportSource = classMahasiswa;
                crystalReportViewer1.Refresh();
            }
            catch (Exception ex)
            {
                simpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
        }
    }
}