//using ClosedXML.Excel;

using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace DemoApp
{
    public partial class ApiCall : Form
    {
        public ApiCall()
        {
            InitializeComponent();
            getDetails();
        }

        public void getDetails()
        {
            var client = new RestClient("https://jsonplaceholder.typicode.com/posts");            
            var request = new RestRequest(Method.GET);
            //request.AddHeader("X-Token-Key", "dsds-sdsdsds-swrwerfd-dfdfd");
            IRestResponse response = client.Execute(request);

            var content = response.Content; // raw content as string
            dynamic json = JsonConvert.DeserializeObject(content);
            
            var client1 = new RestClient("https://jsonplaceholder.typicode.com/comments");            
            IRestResponse response1 = client1.Execute(request);

            var content1 = response1.Content; // raw content as string
            dynamic json1 = JsonConvert.DeserializeObject(content1);

            DataSet ds = new DataSet();
            DataTable dt = new DataTable("posts");
            DataTable dt1 = new DataTable("comments");
            dt1.Columns.Add(new DataColumn("name", typeof(string)));
            dt1.Columns.Add(new DataColumn("email", typeof(string)));
            dt1.Columns.Add(new DataColumn("body", typeof(string)));

            for (int i = 0; i < json.Count; i++) {                

                DataRow dr = dt1.NewRow();
                dr["name"] = json1[i].name;
                dr["email"] = json1[i].email;
                dr["body"]= json1[i].body;
                dt1.Rows.Add(dr);
                
            }
            dt.Columns.Add(new DataColumn("id", typeof(int)));
            dt.Columns.Add(new DataColumn("title", typeof(string)));
            dt.Columns.Add(new DataColumn("body", typeof(string)));

            for (int i = 0; i < json.Count; i++)
            {

                DataRow dr = dt.NewRow();
                dr["id"] = json[i].id;
                dr["title"] = json[i].title;
                dr["body"] = json[i].body;
                dt.Rows.Add(dr);

            }
            ds.Tables.Add(dt);
            ds.Tables.Add(dt1);
            ExportDataSetToExcel(ds, @"C:\Users\Hxtreme\OneDrive\Desktop\exceldownload");
          
        }

        private void ExportDataSetToExcel(DataSet ds, string strPath)
        {
            int inHeaderLength = 0, inColumn = 0, inRow = 0;
            System.Reflection.Missing Default = System.Reflection.Missing.Value;
            //Create Excel File
            strPath += @"\Excel" + DateTime.Now.ToString().Replace(':', '-') + ".xlsx";
            OfficeExcel.Application excelApp = new OfficeExcel.Application();
            OfficeExcel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);
            foreach (DataTable dtbl in ds.Tables)
            {
                //Create Excel WorkSheet
                OfficeExcel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);
                excelWorkSheet.Name = dtbl.TableName;//Name worksheet

                //Write Column Name
                for (int i = 0; i < dtbl.Columns.Count; i++)
                    excelWorkSheet.Cells[inHeaderLength + 1, i + 1] = dtbl.Columns[i].ColumnName.ToUpper();

                //Write Rows
                for (int m = 0; m < dtbl.Rows.Count; m++)
                {
                    for (int n = 0; n < dtbl.Columns.Count; n++)
                    {
                        inColumn = n + 1;
                        inRow = inHeaderLength + 2 + m;
                        excelWorkSheet.Cells[inRow, inColumn] = dtbl.Rows[m].ItemArray[n].ToString();
                        //if (m % 2 == 0)
                        //    excelWorkSheet.get_Range("A" + inRow.ToString(), "G" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#FCE4D6");
                    }
                }

                //Excel Header
                //OfficeExcel.Range cellRang = excelWorkSheet.get_Range("A1", "G3");
                //cellRang.Merge(false);
                //cellRang.Interior.Color = System.Drawing.Color.White;   
                //cellRang.Font.Color = System.Drawing.Color.Gray;
                //cellRang.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignCenter;
                //cellRang.VerticalAlignment = OfficeExcel.XlVAlign.xlVAlignCenter;
                //cellRang.Font.Size = 26;
                //excelWorkSheet.Cells[1, 1] = "Greate Novels Of All Time";

                //Style table column names
                OfficeExcel.Range cellRang = excelWorkSheet.get_Range("A1", "C1");
                OfficeExcel.Range cellRang1 = excelWorkSheet.get_Range("B1", "C1");
                cellRang.Font.Bold = true;
                //cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                //cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ED7D31");
                
                cellRang1.ColumnWidth = 50;
                excelWorkSheet.Columns.AutoFit();
            }
                
            //Delete First Page
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet lastWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[1];
            lastWorkSheet.Delete();
            excelApp.DisplayAlerts = true;

            //Set Defualt Page
            (excelWorkBook.Sheets[1] as OfficeExcel._Worksheet).Activate();

            excelWorkBook.SaveAs(strPath, Default, Default, Default, false, Default, OfficeExcel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
            excelWorkBook.Close();
            excelApp.Quit();

            MessageBox.Show("Excel generated successfully \n As " + strPath);
        }
      
    
    }
}
