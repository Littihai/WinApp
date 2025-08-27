using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace WinApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            // ตั้งชื่อหน้าต่าง
            this.Text = "Run Advice";

            // สร้างปุ่ม
            Button btnRunMacro = new Button();
            btnRunMacro.Text = "Run Advice";
            btnRunMacro.Width = 190;
            btnRunMacro.Height = 30;
            btnRunMacro.Top = 50;
            btnRunMacro.Left = 50;

            // ผูก event เมื่อคลิก
            btnRunMacro.Click += BtnRunMacro_Click;

            // เพิ่มปุ่มเข้า form
            this.Controls.Add(btnRunMacro);
        }

        private void BtnRunMacro_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                // 🔹 Step 1: เรียก Task Scheduler ที่ Server (ให้ Server ไปรันสคริปต์/จ็อบเอง)
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "schtasks";
                psi.Arguments = @"/run /s TSEDB /u tse\administrator /p scsadmin /tn ""Run_SQL_Job_QCC""";
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;

                using (Process cmdProcess = Process.Start(psi))
                {
                    cmdProcess.WaitForExit(); // ✅ รอจน Task ถูกสั่งรัน
                }

                // 🔹 Step 2: เปิด Excel แล้วรัน Macro
                // NOTE: พาธเริ่มต้น (แก้กลับให้ถูก ต้องมีเว้นวรรคก่อน _Qcc)
                string workbookPath = @"D:\Format _Qcc\Summary Forecast Advics.xlsm";
                //@"T:\PO\"

                // ถ้าไม่เจอไฟล์ ให้เปิดกล่องให้ผู้ใช้เลือกไฟล์ .xlsm
                if (!File.Exists(workbookPath))
                {
                    using (OpenFileDialog ofd = new OpenFileDialog()
                    {
                        Title = "เลือกไฟล์ Summary Forecast Advics.xlsm",
                        Filter = "Excel Macro-Enabled (*.xlsm)|*.xlsm|All files (*.*)|*.*",
                        CheckFileExists = true,
                        Multiselect = false
                    })
                    {
                        if (ofd.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show(
                                "ไม่พบไฟล์: " + workbookPath + Environment.NewLine +
                                "และผู้ใช้ไม่ได้เลือกไฟล์ใหม่",
                                "ไม่พบไฟล์",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning
                            );
                            return;
                        }
                        workbookPath = ofd.FileName;
                    }
                }

                excelApp = new Excel.Application();
                excelApp.Visible = true; // แสดง Excel บนหน้าจอ

                workbook = excelApp.Workbooks.Open(workbookPath);

                // ✅ รัน Macro แรก
                excelApp.Run("InsertData");

                // ✅ รัน Macro ที่สอง
                excelApp.Run("Macro1");

                workbook.Save();

                MessageBox.Show("สำเร็จแล้ว", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("❌ เกิดข้อผิดพลาด: " + ex.Message, "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Cleanup Excel
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                }
                catch { /* ignore */ }

                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { /* ignore */ }

                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
