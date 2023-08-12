using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using QA_TVN2_REPORT_MONTHLY.FUNCTION;
using QA_TVN2_REPORT_MONTHLY.MODEL;

namespace QA_TVN2_REPORT_MONTHLY
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        TypeWrite typeGet = new TypeWrite();
        DataConfigLoi configFileError = new DataConfigLoi();
        DataConfigDD configFileDD = new DataConfigDD();
        private void Main_Load(object sender, EventArgs e)
        {

            this.actionButton(false);
            this.updateLable("Lấy dữ liệu config....");
            this.setTime.Start();

        }


        

        private void GetCheckType()
        {
            this.typeGet.km = this.chkKM.Checked;
            this.typeGet.riso = this.chkRiso.Checked;
            this.typeGet.oki = this.chkOkidenki.Checked;
            this.typeGet.kyo = this.chkKyocera.Checked;

            if (this.typeGet.km == false && this.typeGet.riso == false && this.typeGet.kyo == false && this.typeGet.oki == false)
            {
                throw new Exception("Phải Chọn ít nhất 1 loại báo cáo!");
            }

        }
        private void btnActionMain_Click(object sender, EventArgs e)
        {
            try
            {
                this.actionButton(false);
                this.updateLable("Check dữ liệu đầu vào....");


                //Thuc hien lay check cua tung loai
                this.GetCheckType();
                //Check su ton tai cua cac file
                this.CheckFile();

                DateTime monthGet;
                try
                {
                    monthGet = new DateTime(int.Parse(this.txtYear.Text), int.Parse(this.txtMonth.Text), 1);
                }
                catch (Exception ex)
                {
                    throw new Exception("Kiểm tra lại nhập: Tháng, Năm báo cáo! " + ex.Message);
                }
                this.updateLable("Lấy dữ liệu điểm dán....");
                List<DataDD> listDD = new List<DataDD>();
                ActionDD.GetValueDD(ref listDD, this.configFileDD, monthGet);

                this.updateLable("Lấy dữ liệu file Lỗi....");
                List<DataError> listError = new List<DataError>();
                ActionLoi.GetValueError(ref listError, this.configFileError, monthGet);

                this.updateLable("Check dữ liệu và Ghép lỗi....");
                ActionLoi.ParingError(listDD, ref listError);

                this.updateLable("Ghi dữ liệu....");
                string fileName = "";
               ActionWriteDDNew.WriteData(listDD, listError, this.typeGet, ref fileName, this.txtMonth.Text);

                MessageBox.Show($"Tạo file thành công: {fileName} !", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Program", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.actionButton(true);
            }

        }

        private void CheckFile()
        {
            if (!File.Exists(this.txtLoiFile.Text))
            {
                throw new Exception($"File ID:{this.txtLoiFile.Text} => Không tồn tại!");
            }

            if (!File.Exists(this.txtDiemDanFile.Text))
            {
                throw new Exception($"File Assy:{this.txtDiemDanFile.Text} => Không tồn tại!");
            }
            this.configFileDD.pathFile = this.txtDiemDanFile.Text;
            this.configFileDD.sheetName = this.txtDiemDanSheetName.Text;
            this.configFileError.pathFile = this.txtLoiFile.Text;
            this.configFileError.sheetName = this.txtLoiSheetName.Text;

        }

        #region Action Style
        /// <summary>
        /// Thuc hien set nut hanh dong trang thai
        /// </summary>
        /// <param name="action"></param>
        /// CreatedBy: HoaiPT(?/?/2022)
        private void actionButton(bool action)
        {
            if (action == true)
            {
                this.picExecute.Visible = false;
                this.picDone.Visible = true;
                this.tabMain.Enabled = true;

                this.updateLable("Sẵn sàng thực hiện");
            }
            else
            {
                this.tabMain.Enabled = false;

                this.picDone.Visible = false;
                this.picExecute.Visible = true;
            }
            this.picExecute.Update();
            this.picDone.Update();
        }
        /// <summary>
        /// Thuc hien update label 
        /// </summary>
        /// <param name="nameText">Ten label muon cap nhat</param>
        /// CreatedBy: HoaiPT(?/?/2022)
        private void updateLable(string nameText)
        {
            this.lblDisplay.Text = nameText;
            this.lblDisplay.Update();
        }

        #endregion

        private void btnSelectFileG2_Mat_Click(object sender, EventArgs e)
        {
            string temp = MyFuntion2.SelectFile();
            if (temp != "")
            {
                this.txtDiemDanFile.Text = temp;
            }
        }

        private void btnSelectFileG2_AOI_Click(object sender, EventArgs e)
        {
            string temp = MyFuntion2.SelectFile();
            if (temp != "")
            {
                this.txtLoiFile.Text = temp;
            }
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            //Thuc hien clear du lieu
            ClearTextBoxes();

            this.txtMonth.Text = DateTime.Now.AddMonths(-1).ToString("MM");
            this.txtYear.Text = DateTime.Now.ToString("yyyy");

        }
        private void ClearTextBoxes()
        {
            Action<Control.ControlCollection> func = null;

            func = (controls) =>
            {
                foreach (Control control in controls)
                    if (control is TextBox)
                        (control as TextBox).Clear();
                    else
                        func(control.Controls);
            };

            func(Controls);
        }

        private void setTime_Tick(object sender, EventArgs e)
        {
            try
            {
                this.setTime.Stop();

                //this.txtDiemDanFile.Text = @"D:\hoai\Hoai_Daotao\vs\SMT\SMT2_REPORT_MONTHLY\Tai_lieu\TVN2-QUẢN LÍ ĐIỂM DÁN HÀNG NGÀY T7.2023.xlsx";
               // this.txtLoiFile.Text = @"D:\hoai\Hoai_Daotao\vs\SMT\SMT2_REPORT_MONTHLY\Tai_lieu\SMT2- LOI TT TRONG CONG DOAN T07.xlsx";
                this.txtLoiSheetName.Text = "Dữ liệu";
                this.txtDiemDanSheetName.Text = DateTime.Now.AddMonths(-1).ToString("yyyy");

                this.txtMonth.Text = DateTime.Now.AddMonths(-1).ToString("MM");
                this.txtYear.Text = DateTime.Now.AddMonths(-1).ToString("yyyy");

                ActionGetConfig.GetConfigAll();
                ActionGetConfig.GetConfig(ref this.configFileError);
                ActionGetConfig.GetConfig(ref this.configFileDD);

                this.txtComment.Text += Environment.NewLine + ConfigurationManager.AppSettings["TEXT_COMMENT"];

                this.actionButton(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Program", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       
    }
}

