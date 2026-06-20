using ExcelDataReader;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq.Expressions;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace CRUDMahasiswaADO
{
    public partial class Form1 : Form
    {
        private BindingSource bindingSource = new BindingSource();
        private DataTable dtMahasiswa = new DataTable();
        private readonly SqlConnection conn;
        private readonly string ConnectionString =
            "Data Source=MASAYUU-DN5OL92\\MASAYU;Initial Catalog=DBAkademikADO;Integrated Security=True";

        DAL dbLogic = new DAL();

        private void SimpanLog(string message)
        {
            dbLogic.InsertLog(message);
        }

        public Form1()
        {
            InitializeComponent();
            conn = new SqlConnection(ConnectionString);
        }
        private void ConnectDataBase()
        {
            try
            {
                if (conn.State == System.Data.ConnectionState.Closed)
                {
                    conn.Open();
                }

                MessageBox.Show("Koneksi Berhasil");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Koneksi Gagal: " + ex.Message);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataRowView)bindingSource[e.RowIndex]).Row;

                txtNIM.Text = row[0].ToString();
                txtNama.Text = row[1].ToString();
                cmbJK.Text = row[2].ToString();
                dtpTanggalLahir.Value = Convert.ToDateTime(row[3]);
                txtAlamat.Text = row[4].ToString();
                txtKodeProdi.Text = row[6].ToString();

                if (row[5] != DBNull.Value)
                {
                    byte[] imgBytes = (byte[])row[5];
                    using (MemoryStream ms = new MemoryStream(imgBytes))
                    {
                        pbFoto.Image = System.Drawing.Image.FromStream(ms);
                        pbFoto.SizeMode = PictureBoxSizeMode.StretchImage;
                    }
                }
                else
                {
                    pbFoto.Image = null;
                }

                txtNIM.Enabled = false;
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(dbLogic.GetConnectionString()))
                {
                    conn.Open();
                    MessageBox.Show("Koneksi Berhasil");
                }
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Eror :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Eror :" + ex.Message);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] ConvertImageToBytes(PictureBox pb)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        return ms.ToArray();
                    }
                }
                byte[] imgBytes = ConvertImageToBytes(pbFoto);
                dbLogic.InsertMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);
                MessageBox.Show("Data mahasiswa berhaisl ditambahkan");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog("Rollback insert :" + ex.Message);
                MessageBox.Show("SQL eror :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog("General eror :" + ex.Message);
                MessageBox.Show("General eror :" + ex.Message);
            }
        }

        private void ClearForm()
        {
            txtNIM.Enabled = true;
            txtNIM.Clear();
            txtNama.Clear();
            cmbJK.SelectedIndex = -1;
            txtAlamat.Clear();
            txtKodeProdi.Clear();
            dtpTanggalLahir.Value = DateTime.Now;
            pbFoto.Image = null;
            txtNIM.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] ConvertImageToBytes(PictureBox pb)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pb.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        return ms.ToArray();
                    }
                }
                byte[] imgBytes = ConvertImageToBytes(pbFoto);
                dbLogic.UpdateMhs(txtNIM.Text, txtNama.Text, txtAlamat.Text, cmbJK.Text, dtpTanggalLahir.Value.Date, txtKodeProdi.Text, imgBytes);
                MessageBox.Show("Data mahasiswa berhaisl diubah");
                ClearForm();
                btnLoad.PerformClick();
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL eror :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General eror :" + ex.Message);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dg = MessageBox.Show(
                    "Yakin ingin menghapus data?",
                    "Konfirmasi",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (dg == DialogResult.Yes)
                {
                    dbLogic.DeleteMhs(txtNIM.Text);
                    MessageBox.Show("Data mahasiswa berhasil dihapus");
                    ClearForm() ;
                    btnLoad.PerformClick();
                }
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Eror :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Eror :" + ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            // TODO: This line of code loads data into the 'dBAkademikADODataSet.Mahasiswa' table. You can move, or remove it, as needed.
            this.mahasiswaTableAdapter.Fill(this.dBAkademikADODataSet.Mahasiswa);
            cmbJK.Items.Clear();
            cmbJK.Items.Add("L");
            cmbJK.Items.Add("P");

            cmbJK.DataSource = new string[] { "L", "P" };

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            bindingNavigator1.BindingSource = bindingSource;

            dataGridView1.CellClick += dataGridView1_CellContentClick;

            LoadData();
        }


        private void BindControls()
        {
            txtNIM.DataBindings.Clear();
            txtNama.DataBindings.Clear();
            cmbJK.DataBindings.Clear();
            dtpTanggalLahir.DataBindings.Clear();
            txtAlamat.DataBindings.Clear();
            txtKodeProdi.DataBindings.Clear();

            txtNIM.DataBindings.Add("Text", bindingSource, "NIM");
            txtNama.DataBindings.Add("Text", bindingSource, "Nama");
            cmbJK.DataBindings.Add("Text", bindingSource, "JenisKelamin");
            dtpTanggalLahir.DataBindings.Add("Value", bindingSource, "TanggalLahir");
            txtAlamat.DataBindings.Add("Text", bindingSource, "Alamat");
            txtKodeProdi.DataBindings.Add("Text", bindingSource, "KodeProdi");
        }


        private void btnResetData_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.resetData();
                MessageBox.Show("Data berhasil direset");
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("SQL Eror: " + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General Eror: " + ex.Message);
            }
        }

        private void btnTestInjection_Click(object sender, EventArgs e)
        {
            try
            {
                dbLogic.testInject(txtNIM.Text);
                LoadData() ;
            }
            catch (SqlException ex)
            {
                if(ex.Message.Contains("safe"))
                {
                    SimpanLog(ex.Message);
                    MessageBox.Show("SQL Eror : Unsafe UPDATE operation not allowed");
                }
                else
                {
                    SimpanLog(ex.Message);
                    MessageBox.Show("SQL Eror :" + ex.Message);
                }  
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("General eror :" + ex.Message);
            }
        }


        private void HitungTotal() //hitung total (output mahasiswa)
        {
            try
            {
                int total = (dbLogic.CountMhs().Equals(DBNull.Value)) ? 0 : dbLogic.CountMhs();

                lblTotal.Text = "Total Mahasiswa : " + total;
            }

            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }
        }


        private void LoadData()
        {
            try
            {
                bindingSource.DataSource = dbLogic.GetMhs();
                dataGridView1.DataSource = bindingSource;
                DataGridViewImageColumn fotoColumn = (DataGridViewImageColumn)dataGridView1.Columns["Foto"];
                fotoColumn.ImageLayout = DataGridViewImageCellLayout.Stretch;

                HitungTotal();
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    Console.WriteLine("Name: " + col.Name + " | DataPropertyName: " + col.DataPropertyName);
                }
                dataGridView1.Enabled = true;
                btnImportDatabase.Enabled = false;
                btnCreate.Enabled = true;
                btnUpdate.Enabled = true;
                btnDelete.Enabled = true;
                btnCari.Enabled = true;
                btnLoad.Enabled = true;
                btnResetData.Enabled = true;
                btnTestInjection.Enabled = true;
            }
            catch (Exception ex)
            {
                SimpanLog(ex.Message);
                MessageBox.Show("Gagal load data: " + ex.Message);
            }

        }

        private void btnRekapData_Click(object sender, EventArgs e)
        {
            Form2 fm3 = new Form2();
            fm3.Show(); ;
            this.Hide();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pbFoto.Image = System.Drawing.Image.FromFile(ofd.FileName);
                pbFoto.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }

        private void btnInportExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel workbook| *. xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            DataTable dt = result.Tables[0];
                            dataGridView1.DataSource = dt;
                            dataGridView1.Enabled = false;

                            btnImportDatabase.Enabled = true;
                            btnCreate.Enabled = false;
                            btnUpdate.Enabled = false;
                            btnDelete.Enabled = false;
                            btnCari.Enabled = false;
                            btnLoad.Enabled = false;
                            btnResetData.Enabled = false;
                            btnTestInjection.Enabled = false;
                        }
                    }
                }
            }
        }

        private void btnImportDatabase_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)dataGridView1.DataSource;

                if (dt == null || dt.Rows.Count == 0)
                {
                    MessageBox.Show("Tidak ada data untuk diimport.");
                    return;
                }

                int sukses = 0;

                foreach (DataRow row in dt.Rows)
                {
                    string Nim = row["NIM"].ToString().Trim();
                    string Nama = row["Nama"].ToString().Trim();
                    string Jk = row["JenisKelamin"].ToString().Trim();
                    string Alamat = row["Alamat"].ToString().Trim();
                    string KodeProdi = row["NamaProdi"].ToString().Trim();
                    string FotoPath = row.Table.Columns.Contains("FotoPath")
                                        ? row["FotoPath"].ToString().Trim() : string.Empty;

                    if (string.IsNullOrEmpty(Nim) || string.IsNullOrEmpty(Nama))
                        continue;

                    DateTime tglLahir;
                    if (!DateTime.TryParse(row["TanggalLahir"].ToString(), out tglLahir))
                        continue;

                    byte[] ConvertImageFromPath(string path)
                    {
                        if (string.IsNullOrWhiteSpace(path))
                            return null;

                        if (!File.Exists(path))
                            return null;

                        return File.ReadAllBytes(path);
                    }
                    byte[] fotoBytes = ConvertImageFromPath(FotoPath);

                    dbLogic.InsertMhs(Nim, Nama, Alamat, Jk, tglLahir, KodeProdi, fotoBytes);

                    sukses++;
                }

                MessageBox.Show("Data mahasiswa berhasil ditambahkan");
                ClearForm();
                LoadData();
            }
            catch (SqlException ex)
            {
                SimpanLog("Rollback Insert :" + ex.Message);
                MessageBox.Show("SQL Error :" + ex.Message);
            }
            catch (Exception ex)
            {
                SimpanLog("General eror :" + ex.Message);
                MessageBox.Show("General Error :" + ex.Message);
            }
        }
    }
}




