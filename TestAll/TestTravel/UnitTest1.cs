using NUnit.Framework;
using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using TestTravel.Models;
using System.Linq;

namespace TestTravel
{
    public class Tests
    {

        [Test]
        public void TestExcel()
        {
            var isSuccess = ExcelCheckLogin();
            Assert.True(isSuccess);
        }

        public bool ExcelCheckLogin()
        {
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Users\\Windows\\OneDrive\\Desktop\\DATN\\API\\TestTravel\\test.xlsx");

                Excel.Range xlTestRange;
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                xlTestRange = xlWorksheet.UsedRange;

                //xlWorksheet.Range["F2"].Value = "test";

                // tổng dòng có dữ liệu
                int row_used = xlTestRange.Cells[1, 1].End(Excel.XlDirection.xlDown).Row;
                var lsCorrect = new List<TestLoginModel>();
                var lsNotCorrect = new List<TestLoginFailedModel>();

                for (int i = 2; i <= row_used; i++)
                {
                    string resultTK = xlWorksheet.Range[$"A{i}"].Value.ToString();
                    string resultPW = xlWorksheet.Range[$"B{i}"].Value.ToString();

                    string resultTKCorrect = xlWorksheet.Range[$"C{i}"].Value.ToString();
                    string resultPWCorrect = xlWorksheet.Range[$"D{i}"].Value.ToString();

                    string KetQua = xlWorksheet.Range[$"E{i}"].Value.ToString();


                    // so sánh các trường hợp theo như yêu cầu mong muốn 
                    //(resultTK == resultTKCorrect && resultPW == resultPWCorrect) => true
                    //(resultTK == resultTKCorrect && resultPW == resultPWCorrect) => false
                    if ((resultTK == resultTKCorrect && resultPW == resultPWCorrect) == bool.Parse(KetQua)) // trường hợp đúng
                    {
                        TestLoginModel correctObj = new TestLoginModel();
                        correctObj.TaiKhoanDung = resultTK;
                        correctObj.MatKhauDung = resultPW;
                        lsCorrect.Add(correctObj);

                        xlWorksheet.Range[$"F{i}"].Value = "PASS";

                    }
                    else // trường hợp sai
                    {

                        TestLoginFailedModel notObject = new TestLoginFailedModel();
                        notObject.TaiKhoanSai = resultTK;
                        notObject.MatKhauSai = resultPW;
                        lsNotCorrect.Add(notObject);

                        xlWorksheet.Range[$"F{i}"].Value = "FAIL";

                    }
                }
                xlWorkbook.Save();
                xlWorkbook.Close();
                if (lsNotCorrect.Count() > 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
                //return lsNotCorrect.Count() == 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}