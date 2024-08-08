using ClosedXML.Excel;
using Microsoft.Win32;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;


namespace WpfApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string filePath;

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
                txtPath.Text = filePath;
                LoadData(filePath);
            }
        }

        private void LoadData(string filePath)
        {
            var projects = ReadExcelFile(filePath);

            dtgProject.Columns.Clear();

            dtgProject.Columns.Add(new DataGridTextColumn { Header = "STT", Binding = new System.Windows.Data.Binding("ID") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Mã Đề Tài", Binding = new System.Windows.Data.Binding("Code") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Tên đề tài Tiếng Anh/ Tiếng Nhật", Binding = new System.Windows.Data.Binding("ForgeinName") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Tên đề tài Tiếng Việt", Binding = new System.Windows.Data.Binding("VietName") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Department", Binding = new System.Windows.Data.Binding("Department") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "GVHD", Binding = new System.Windows.Data.Binding("GVHD") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Email GVHD1", Binding = new System.Windows.Data.Binding("Email1") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Email GVHD2", Binding = new System.Windows.Data.Binding("Email2") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Result 1", Binding = new System.Windows.Data.Binding("Result1") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Result 2", Binding = new System.Windows.Data.Binding("Result2") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Comment 1", Binding = new System.Windows.Data.Binding("Cmt1") });
            dtgProject.Columns.Add(new DataGridTextColumn { Header = "Comment 2", Binding = new System.Windows.Data.Binding("Cmt2") });

            dtgProject.Columns.Add(new DataGridTextColumn
            {
                Header = "Final Result",
                Binding = new System.Windows.Data.Binding("Final")
            });

            dtgProject.ItemsSource = projects;
        }


        public List<CapstoneProject> ReadExcelFile(string filePath)
        {
            var projects = new List<CapstoneProject>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RowsUsed().Skip(2);

                foreach (var row in rows)
                {
                    var result1 = row.Cell(9).GetString();
                    var result2 = row.Cell(10).GetString();

                    string finalResult = string.Empty;

                    if (!string.IsNullOrWhiteSpace(result1) && !string.IsNullOrWhiteSpace(result2))
                    {
                        finalResult = CheckResult(result1, result2);
                    }

                    var project = new CapstoneProject
                    {
                        ID = row.Cell(1).GetString(),
                        Code = row.Cell(2).GetString(),
                        ForgeinName = row.Cell(3).GetString(),
                        VietName = row.Cell(4).GetString(),
                        Department = row.Cell(5).GetString(),
                        GVHD = row.Cell(6).GetString(),
                        Email1 = row.Cell(7).GetString(),
                        Email2 = row.Cell(8).GetString(),
                        Result1 = result1,
                        Result2 = result2,
                        Cmt1 = row.Cell(11).GetString(),
                        Cmt2 = row.Cell(12).GetString(),
                        Final = finalResult
                    };
                    projects.Add(project);
                }
            }

            return projects;
        }


        public class CapstoneProject
        {
            public string ID { get; set; }
            public string Code { get; set; }
            public string ForgeinName { get; set; }
            public string VietName { get; set; }
            public string Department { get; set; }
            public string GVHD { get; set; }
            public string Email1 { get; set; }
            public string Email2 { get; set; }
            public string Result1 { get; set; }
            public string Result2 { get; set; }
            public string Cmt1 { get; set; }
            public string Cmt2 { get; set; }
            public string Final { get; set; }
        }

        public string CheckResult(string res1, string res2)
        {
            if (string.IsNullOrEmpty(res1) || string.IsNullOrEmpty(res2)) return "";

            if (res1.Equals(res2))
            {
                return res1;
            }

            if ((res1 == "Pass" && res2 == "Fail") || (res1 == "Fail" && res2 == "Pass"))
            {
                return "Consider";
            }

            if ((res1 == "Fail" && res2 == "Consider") || (res1 == "Consider" && res2 == "Fail"))
            {
                return "Fail";
            }

            if ((res1 == "Pass" && res2 == "Consider") || (res1 == "Consider" && res2 == "Pass"))
            {
                return "Pass";
            }

            return "Consider";
        }

        private void btnSendMail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var projects = new List<CapstoneProject>();

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheets.First();
                    var rows = worksheet.RowsUsed().Skip(2);

                    foreach (var row in rows)
                    {
                        var result1 = row.Cell(9).GetString();
                        var result2 = row.Cell(10).GetString();

                        string finalResult = string.Empty;

                        if (!string.IsNullOrWhiteSpace(result1) && !string.IsNullOrWhiteSpace(result2))
                        {
                            finalResult = CheckResult(result1, result2);
                        }

                        var project = new CapstoneProject
                        {
                            ID = row.Cell(1).GetString(),
                            Code = row.Cell(2).GetString(),
                            ForgeinName = row.Cell(3).GetString(),
                            VietName = row.Cell(4).GetString(),
                            Department = row.Cell(5).GetString(),
                            GVHD = row.Cell(6).GetString(),
                            Email1 = row.Cell(7).GetString(),
                            Email2 = row.Cell(8).GetString(),
                            Result1 = result1,
                            Result2 = result2,
                            Cmt1 = row.Cell(11).GetString(),
                            Cmt2 = row.Cell(12).GetString(),
                            Final = finalResult
                        };

                        if (!string.IsNullOrEmpty(project.Email1))
                        {
                            SendEmail(project.GVHD.Split('/')[0].Trim(), project.Email1, project);
                        }

                        if (!string.IsNullOrEmpty(project.Email2) && project.GVHD.Split('/').Length > 1)
                        {
                            SendEmail(project.GVHD.Split('/')[1].Trim(), project.Email2, project);
                        }
                    }

                    MessageBox.Show("Emails sent successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        private async void SendEmail(string name, string email, CapstoneProject project)
        {
            try
            {
                string subject = $"Thông báo kết quả dự án: ";
                string body = $@"
            <p>Kính gửi {name},</p>
            <p>Bộ phận Capstone Project gửi kết quả review các đề tài của thầy cô như sau:</p>
            <p><strong>{project.VietName}</strong></p>
            <table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%;'>
                <tr>
                    <th>STT</th>
                    <th>Mã Đề Tài</th>
                    <th>Tên đề tài Tiếng Anh/ Tiếng Nhật</th>
                    <th>Tên đề tài Tiếng Việt</th>
                    <th>Department</th>
                    <th>GVHD</th>
                    <th>Email GVHD1</th>
                    <th>Email GVHD2</th>
                    <th>Result 1</th>
                    <th>Result 2</th>
                    <th>Comment 1</th>
                    <th>Comment 2</th>
                    <th>Final</th>
                </tr>
                <tr>
                    <td>{project.ID}</td>
                    <td>{project.Code}</td>
                    <td>{project.ForgeinName}</td>
                    <td>{project.VietName}</td>
                    <td>{project.Department}</td>
                    <td>{project.GVHD}</td>
                    <td>{project.Email1}</td>
                    <td>{project.Email2}</td>
                    <td>{project.Result1}</td>
                    <td>{project.Result2}</td>
                    <td>{project.Cmt1}</td>
                    <td>{project.Cmt2}</td>
                    <td>{project.Final}</td>
                </tr>
            </table>
            <p>Các đề tài Pass sẽ được công bố cho các nhóm SV chọn ở đầu Học kì.</p>
            <p>Các đề tài Not Pass sẽ không được công bố và quí thầy cô vui lòng không gửi đăng ký lại cho đợt 2.</p>
            <p>Các đề tài Consider, quí Thầy Cô sẽ cập nhật lại theo 10 tiêu chí đánh giá (trong file đính kèm) để hoàn thiện hơn và gửi lại ở đợt 2 nếu thấy phù hợp.</p>
            <p>Trân trọng,</p>";

                MailMessage mail = new MailMessage()
                {
                    From = new MailAddress("nextintern.corp@gmail.com", "CapstoneProjectReview"),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true
                };

                mail.To.Add(email);

                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential("nextintern.corp@gmail.com", "wflm cyhu ifww lnbz"),
                    EnableSsl = true
                };

                await smtpClient.SendMailAsync(mail);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error sending email to {name} ({email}): {ex.Message}");
            }
        }

    }
}
