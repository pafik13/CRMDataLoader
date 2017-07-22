using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using RestSharp;
using System.Net;

using CRMLite.Entities;
using CRMLite.DaData;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SBL.DataLoader
{
    public partial class MainForm : Form
    {
        //private Excel.Application ExcelApp = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                ReadAndUploadNet(fileDialog.FileName);
            }
        }

        void ReadAndUploadNet(string filePath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("АС");
            workSheet.Select(Type.Missing);

            for (int i = 1; i < 255; i++)
            {
                string val = workSheet.Cells[i, "A"].Value;
                if (string.IsNullOrEmpty(val)) break;
                Console.WriteLine("{0}: {1}", i, val);
                UploadData(new Net()
                {
                    name = val
                });
            }
        }

        void UploadData<T>(T data)
        {
            var client = new RestClient(Global.URL_C9);
            var path = typeof(T).Name;

            var request = new RestRequest(path, Method.POST).AddJsonBody(data);
            var response = client.Execute(request);
            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                case HttpStatusCode.Created:
                    break;
                default:
                    string message = string.Format("StatusCode: {0}, StatusDescription: {1}, Content: {2}", response.StatusCode, response.StatusDescription, response.Content);
                    //MessageBox.Show(message);
                    Console.WriteLine(message);
                    return;
                    //break;
            }
            Console.WriteLine(response.StatusDescription);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var client = new RestClient(Global.URL);

            var request = new RestRequest("Subway?limit=300", Method.GET);
            var response = client.Execute<List<Subway>>(request);
            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    foreach (var item in response.Data)
                    {
                        UploadData(new Subway()
                        {
                            name = item.name,
                            city = item.city
                        });
                    }
                    break;
                default:
                    string message = string.Format("StatusCode: {0}, StatusDescription: {1}, Content: {2}", response.StatusCode, response.StatusDescription, response.Content);
                    MessageBox.Show(message);
                    break;
            }
            Console.WriteLine(response.StatusDescription);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var client = new RestClient(Global.URL);

            var request = new RestRequest("Region?limit=300", Method.GET);
            var response = client.Execute<List<Region>>(request);
            switch (response.StatusCode)
            {
                case HttpStatusCode.OK:
                    foreach (var item in response.Data)
                    {
                        UploadData(new Region()
                        {
                            name = item.name,
                            city = item.city
                        });
                    }
                    break;
                default:
                    string message = string.Format("StatusCode: {0}, StatusDescription: {1}, Content: {2}", response.StatusCode, response.StatusDescription, response.Content);
                    MessageBox.Show(message);
                    break;
            }
            Console.WriteLine(response.StatusDescription);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                FindSubways(fileDialog.FileName);
                //FindRegions(fileDialog.FileName);
            }
        }

        void FindSubways(string filePath)
        {
            if (string.IsNullOrEmpty(AgentUUID.Text))
            {
                throw new Exception(@"Не заполнен Город");
            }

            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);


            var client = new RestClient(Global.JohnsonURL);
            IRestRequest request;
            IRestResponse<List<Subway>> responseSubway;
            IRestResponse<List<Region>> responseRegion;

            for (int i = 2; i <= 201; i++)
            {
                string val = workSheet.Cells[i, "K"].Value;
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);
                string pathSubway = "Subway?where={\"name\":\"" + val + "\", \"city\":\"" + AgentUUID.Text + "\"}";
                request = new RestRequest (pathSubway, Method.GET);
                responseSubway = client.Execute<List<Subway>>(request);
                if (responseSubway.StatusCode == HttpStatusCode.OK)
                {
                    switch(responseSubway.Data.Count)
                    {
                        case 0:
                            Console.WriteLine(@"Data.Count==0: {0}", val);
                            break;
                        case 1:
                            workSheet.Cells[i, "L"].Value = responseSubway.Data[0].uuid;
                            break;
                        default:
                            Console.WriteLine(@"Data.Count>1: {0}", val);
                            break;
                    }
                }
                else
                {
                    string message = string.Format(
                        "StatusCode: {0}, StatusDescription: {1}, Content: {2}", 
                        responseSubway.StatusCode, responseSubway.StatusDescription, responseSubway.Content
                        );
                    Console.WriteLine(message);
                }
                Console.WriteLine(
                    "StatusCode: {0}, Path: {1}",
                    responseSubway.StatusDescription, pathSubway
                    );

                string pathRegion = "Region?where={\"name\":\"" + val + "\", \"city\":\"" + AgentUUID.Text + "\"}";
                request = new RestRequest(pathRegion, Method.GET);
                responseRegion = client.Execute<List<Region>>(request);
                if (responseRegion.StatusCode == HttpStatusCode.OK)
                {
                    switch (responseRegion.Data.Count)
                    {
                        case 0:
                            Console.WriteLine(@"Data.Count==0: {0}", val);
                            break;
                        case 1:
                            workSheet.Cells[i, "M"].Value = responseRegion.Data[0].uuid;
                            break;
                        default:
                            Console.WriteLine(@"Data.Count>1: {0}", val);
                            break;
                    }
                }
                else
                {
                    string message = string.Format(
                        "StatusCode: {0}, StatusDescription: {1}, Content: {2}", 
                        responseRegion.StatusCode, responseRegion.StatusDescription, responseRegion.Content
                        );
                    Console.WriteLine(message);
                }
                Console.WriteLine(
                    "StatusCode: {0}, Path: {1}",
                    responseRegion.StatusDescription, pathRegion
                    );
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //ExcelApp.Workbooks.Close();

            //ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        void FindRegions(string filePath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);


            var client = new RestClient(Global.URL);
            IRestRequest request;
            IRestResponse<List<Region>> response;

            for (int i = 2; i <= 201; i++)
            {
                string val = workSheet.Cells[i, "K"].Value;
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);
                string path = "Region?where={\"name\":\"" + val + "\"}";
                request = new RestRequest(path, Method.GET);
                response = client.Execute<List<Region>>(request);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    switch (response.Data.Count)
                    {
                        case 0:
                            Console.WriteLine(@"Data.Count==0: {0}", val);
                            break;
                        case 1:
                            workSheet.Cells[i, "M"].Value = response.Data[0].uuid;
                            break;
                        default:
                            Console.WriteLine(@"Data.Count>1: {0}", val);
                            break;
                    }
                }
                else
                {
                    string message = string.Format("StatusCode: {0}, StatusDescription: {1}, Content: {2}", response.StatusCode, response.StatusDescription, response.Content);
                    Console.WriteLine(message);
                }
                Console.WriteLine(response.StatusDescription);
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //ExcelApp.Workbooks.Close();

            //ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                FindAddress(fileDialog.FileName);
            }
        }

        void FindAddress(string filePath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);


            var api = new SuggestClient(Global.DadataApiToken, Global.DadataApiURL);


            for (int i = 2; i <= 201; i++)
            {
                string val = workSheet.Cells[i, "E"].Value;
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);
                var suggestions = api.QueryAddress(val);
                if (suggestions.suggestionss.Count > 0)
                {
                    var sugg = suggestions.suggestionss[0];
                    workSheet.Cells[i, "F"].Value = sugg.value;
                    workSheet.Cells[i, "G"].Value = sugg.data.qc_geo;
                    workSheet.Cells[i, "H"].Value = sugg.data.geo_lat;
                    workSheet.Cells[i, "I"].Value = sugg.data.geo_lon;
                }
                else
                {
                    Console.WriteLine("suggestions.suggestionss.Count == 0");
                }
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //ExcelApp.Workbooks.Close();

            //ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                FindNet(fileDialog.FileName);
            }
        }

        void FindNet(string filePath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);

            var client = new RestClient(Global.JohnsonURL);
            IRestRequest request;
            IRestResponse<List<Net>> response;

            var cache = new Dictionary<string, string>();
            for (int i = 2; i <= 201; i++)
            {
                string val = workSheet.Cells[i, "O"].Value;
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);

                if (cache.ContainsKey(val))
                {
                    workSheet.Cells[i, "P"].Value = cache[val];
                    continue;
                }

                string path = "Net?where={\"name\":\"" + val.Replace("+", "%2B") + "\"}";
                request = new RestRequest(path, Method.GET);
                response = client.Execute<List<Net>>(request);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    switch (response.Data.Count)
                    {
                        case 0:
                            Console.WriteLine(@"Data.Count==0: {0}", val);
                            break;
                        case 1:
                            workSheet.Cells[i, "P"].Value = response.Data[0].uuid;
                            cache.Add(val, response.Data[0].uuid);
                            break;
                        default:
                            Console.WriteLine(@"Data.Count>1: {0}", val);
                            break;
                    }
                }
                else
                {
                    string message = string.Format(
                        "StatusCode: {0}, StatusDescription: {1}, Content: {2}",
                        response.StatusCode, response.StatusDescription, response.Content
                    );
                    Console.WriteLine(message);
                }
                Console.WriteLine(
                    "StatusCode: {0}, Path: {1}",
                    response.StatusDescription, path
                    );
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            ExcelApp.Workbooks.Close();

            ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                FindCategory(fileDialog.FileName);
            }
        }

        void FindCategory(string filePath)
        {
            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);

            var client = new RestClient(Global.JohnsonURL);
            IRestRequest request;
            IRestResponse<List<Net>> response;

            var cache = new Dictionary<string, string>();
            for (int i = 2; i <= 201; i++)
            {
                object obj = workSheet.Cells[i, "S"].Value2;
                string val = obj == null ? string.Empty : obj.ToString();
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);

                if (cache.ContainsKey(val))
                {
                    workSheet.Cells[i, "T"].Value = cache[val];
                    continue;
                }

                string path = "Category?where={\"name\":\"" + val + "\"}";
                request = new RestRequest(path, Method.GET);
                response = client.Execute<List<Net>>(request);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    switch (response.Data.Count)
                    {
                        case 0:
                            Console.WriteLine(@"Data.Count==0: {0}", val);
                            break;
                        case 1:
                            workSheet.Cells[i, "T"].Value = response.Data[0].uuid;
                            cache[val] = response.Data[0].uuid;
                            break;
                        default:
                            Console.WriteLine(@"Data.Count>1: {0}", val);
                            break;
                    }
                }
                else
                {
                    string message = string.Format(
                        "StatusCode: {0}, StatusDescription: {1}, Content: {2}",
                        response.StatusCode, response.StatusDescription, response.Content
                        );
                    Console.WriteLine(message);
                }
                Console.WriteLine(
                    "StatusCode: {0}, Path: {1}",
                    response.StatusDescription, path
                    );
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            ExcelApp.Workbooks.Close();

            ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                UploadPharmacies(fileDialog.FileName);
            }
        }

        void UploadPharmacies(string filePath)
        {
            if (string.IsNullOrEmpty(AgentUUID.Text))
            {
                throw new Exception(@"Не заполнен AgentUUID");
            }

            Guid.Parse(AgentUUID.Text);

            if (string.IsNullOrEmpty(AccessToken.Text))
            {
                throw new Exception(@"Не заполнен AccessToken");
            }

            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист1");
            workSheet.Select(Type.Missing);

            var client = new RestClient(Global.JohnsonURL);
            IRestRequest request;
            IRestResponse response;

            for (int i = 2; i <= 201; i++)
            {
                string val = workSheet.Cells[i, "D"].Value;
                if (string.IsNullOrEmpty(val)) continue;

                Console.WriteLine("{0}: {1}", i, val);

                string uuid = workSheet.Cells[i, "B"].Value;
                if (!string.IsNullOrEmpty(uuid))
                {
                    Console.WriteLine("{0}: uuid={1}", i, uuid);
                    continue;
                }


                var pharmacy = new Pharmacy();
                pharmacy.UUID = Guid.NewGuid().ToString();
                //Console.WriteLine("{0}: uuid={1}", i, uuid);
                //pharmacy.UUID = uuid;
                pharmacy.CreatedBy = AgentUUID.Text;
                pharmacy.CreatedAt = DateTimeOffset.UtcNow;
                pharmacy.UpdatedAt = DateTimeOffset.UtcNow;
                pharmacy.isNeedLC = cbIsNeedLC.Checked;
                pharmacy.SetState(PharmacyState.psActive);
                //pharmacy.NumberName = string.Format("Номер в Excel форме: {0}", workSheet.Cells[i, "A"].Value);
                pharmacy.LegalName = workSheet.Cells[i, "C"].Value == null ? string.Empty : workSheet.Cells[i, "C"].Value;
                var address = workSheet.Cells[i, "F"].Value;
                if (string.IsNullOrEmpty(address)) throw new Exception("Пустой адрес!");
                pharmacy.Address = address;
                pharmacy.Phone = workSheet.Cells[i, "J"].Value == null ? string.Empty : workSheet.Cells[i, "J"].Value.ToString();
                pharmacy.Subway = workSheet.Cells[i, "L"].Value == null ? string.Empty : workSheet.Cells[i, "L"].Value.ToString();
                pharmacy.Region = workSheet.Cells[i, "M"].Value == null ? string.Empty : workSheet.Cells[i, "M"].Value.ToString();
                pharmacy.Net = workSheet.Cells[i, "P"].Value == null ? string.Empty : workSheet.Cells[i, "P"].Value.ToString();
                pharmacy.Brand = workSheet.Cells[i, "Q"].Value == null ? string.Empty : workSheet.Cells[i, "Q"].Value.ToString();
                pharmacy.Category = workSheet.Cells[i, "T"].Value == null ? string.Empty : workSheet.Cells[i, "T"].Value.ToString();
                pharmacy.Comment = workSheet.Cells[i, "U"].Value == null ? string.Empty : workSheet.Cells[i, "U"].Value.ToString();

                string path = "Pharmacy?access_token=" + AccessToken.Text;
                request = new RestRequest(path, Method.POST);
                request.AddJsonBody(pharmacy);
                response = client.Execute(request);
                switch (response.StatusCode)
                {
                    case HttpStatusCode.OK:
                    case HttpStatusCode.Created:
                        workSheet.Cells[i, "B"].Value = pharmacy.UUID;
                        break;
                    default:
                        string message = string.Format(
                            "StatusCode: {0}, StatusDescription: {1}, Content: {2}",
                            response.StatusCode, response.StatusDescription, response.Content
                            );
                        Console.WriteLine(message);
                        break;
                }
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            ExcelApp.Workbooks.Close();

            ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (fileDialog.ShowDialog(this) == DialogResult.OK)
            {
                //filePath = oFileDlg.FileName;
                //txtFile.Text = Path.GetFileName(oFileDlg.FileName);
                //MessageBox.Show(fileDialog.FileName);
                UploadEmployees(fileDialog.FileName);
            }
        }

        void UploadEmployees(string filePath)
        {
            if (string.IsNullOrEmpty(AgentUUID.Text))
            {
                throw new Exception(@"Не заполнен AgentUUID");
            }

            Guid.Parse(AgentUUID.Text);

            if (string.IsNullOrEmpty(AccessToken.Text))
            {
                throw new Exception(@"Не заполнен AccessToken");
            }

            var ExcelApp = new Excel.Application();
            ExcelApp.Workbooks.Open(filePath,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
               );
            ExcelApp.Visible = true;
            Excel.Worksheet workSheet = ExcelApp.Worksheets.get_Item("Лист2");
            workSheet.Select(Type.Missing);

            var client = new RestClient(Global.JohnsonURL);
            IRestRequest request;
            IRestResponse response;

            for (int i = 5; i <= 204; i++)
            {
                object val = workSheet.Cells[i, "E"].Value;
                if (val == null) continue;
                string pharmacyUUID = val.ToString();
                if (string.IsNullOrEmpty(pharmacyUUID)) continue;
                if (pharmacyUUID == "0") continue;

                Guid.Parse(pharmacyUUID);
                Console.WriteLine("{0}: {1}", i, pharmacyUUID);

                for (int j = 0; j < 6; j++)
                {
                    string FIO = workSheet.Cells[i, 6 + j * 2].Value;

                    if (string.IsNullOrEmpty(FIO)) continue;

                    string pos = workSheet.Cells[i, 6 + j * 2 + 1].Value;
                    Guid.Parse(pos);

                    var employee = new Employee();
                    employee.UUID = Guid.NewGuid().ToString();
                    employee.CreatedBy = AgentUUID.Text;
                    employee.CreatedAt = DateTimeOffset.UtcNow;
                    employee.UpdatedAt = DateTimeOffset.UtcNow;

                    employee.SetSex(Sex.Female);
                    employee.Name = FIO;
                    employee.Position = pos;
                    employee.Pharmacy = pharmacyUUID;

                    string path = "Employee?access_token=" + AccessToken.Text;
                    request = new RestRequest(path, Method.POST);
                    request.AddJsonBody(employee);
                    response = client.Execute(request);
                    switch (response.StatusCode)
                    {
                        case HttpStatusCode.OK:
                        case HttpStatusCode.Created:
                            workSheet.Cells[i, "B"].Value = employee.UUID;
                            break;
                        default:
                            string message = string.Format(
                                "StatusCode: {0}, StatusDescription: {1}, Content: {2}",
                                response.StatusCode, response.StatusDescription, response.Content
                                );
                            Console.WriteLine(message);
                            break;
                    }

                }
            }

            // Garbage collecting
            GC.Collect();
            GC.WaitForPendingFinalizers();

            ExcelApp.Workbooks.Close();

            ExcelApp.Quit();


            // Clean up references to all COM objects
            // As per above, you're just using a Workbook and Excel Application instance, so release them:
            //Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(ExcelApp);
        }
    }
}
