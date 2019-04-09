using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Standalone_UploadExcelToDatabaseLibrary
{

    public class UploadExcelToDatabaseMethods
    {


        /// <summary>
        /// This method will try to upload an Excel file to a certain table in a database.
        /// In order to succeed the Excel Column names and number MUST match with the ones on the table, if that's not the case
        /// It will return the reason / cause of the failure
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns>string with the result of the process, either Successful or Unsuccessful</returns>
        /// <summary>
        /// This method will try to upload an Excel file to a certain table in a database.
        /// In order to succeed the Excel Column names and number MUST match with the ones on the table, if that's not the case
        /// It will return the reason / cause of the failure
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns>string with the result of the process, either Successful or Unsuccessful</returns>
        public  string VerifiedDataAndUploadToDatabase(string pathToExcelFile, string databaseConnectionString, string tableName, int columnsToIgnore)
        {
            int excelRows = GetTotalAmountOfRowsInExcel(pathToExcelFile)-1;
            int dataBasePreviousRows = countDataBaseTableColumns(tableName, databaseConnectionString);

            string excelSheetName = getExcelActiveSheetName(pathToExcelFile);

            List<string> excelColumNames = getExcelColumnNames(pathToExcelFile);
            List<string> databaseColumnNames = getDataBaseTableColumns(databaseConnectionString, tableName);

            string comparationTest = CompareColumnsBetweenSystems(excelColumNames, databaseColumnNames, columnsToIgnore);

            if (comparationTest!="Approved")
            {
                return comparationTest;
            }

            comparationTest=uploadFileToDatabase(pathToExcelFile, excelSheetName, databaseConnectionString, tableName);

            if (comparationTest!="Success")
            {
                return comparationTest;
            }

            int dataBaseNewRows = countDataBaseTableColumns(tableName, databaseConnectionString);

            if (dataBaseNewRows-dataBasePreviousRows==excelRows)
            {
                return "Upload Successful";
            }
            else
            {
                return "Upload Unsuccessful, missing: "+(dataBaseNewRows-dataBasePreviousRows+excelRows);
            }

            //return "";
        }


        /// <summary>
        /// This method performs a select Count(*) operation on the designed table and returns the register count
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="databaseConnectionString"></param>
        /// <returns>integer with the register count found</returns>
        private  int countDataBaseTableColumns(string tableName, string databaseConnectionString)
        {
            string query = "SELECT COUNT(*) FROM ["+tableName+"]";
            int count = 0;
            using (SqlConnection con = GetConnectionDevDataBase(databaseConnectionString))
            {
                using (SqlCommand commandCount = new SqlCommand(query, con))
                {
                    con.Open();
                    count=(int)commandCount.ExecuteScalar();
                }
            }

            return count;
        }


        /// <summary>
        /// This method compares both the column number, the column names and order for both the Excel File and the 
        /// database table to Match, if one case fails, it returns the failure cause. Otherwise it will return approved
        /// </summary>
        /// <param name="excelColumns"></param>
        /// <param name="dataBaseColumns"></param>
        /// <returns>string with comparation result</returns>
        private  string CompareColumnsBetweenSystems(List<string> excelColumns, List<string> dataBaseColumns, int columnsToIgnore)
        {


            if (excelColumns.Count==0)
            {
                return "Excel File either had 1 column blank at the start or was empty.";
            }

            if (dataBaseColumns.Count==0)
            {
                return "The Data Base column Count was zero, please check if your table name was correct.";
            }

            if ((dataBaseColumns.Count-columnsToIgnore)!=excelColumns.Count)
            {
                return "Found: "+dataBaseColumns.Count+" Columns in the database and: "+excelColumns.Count+" columns in the excel file, they don't match, process aborted.";
            }

            for (int i = 0; i<(dataBaseColumns.Count-columnsToIgnore); i++)
            {
                if (dataBaseColumns[i]!=excelColumns[i])
                {

                    return "There was a mismatch in the order of columns from the Excel File towards the database. Column in DB: "+dataBaseColumns[i].ToString()+" Column at the Excel File: "+excelColumns[i].ToString();
                }
            }

            return "Approved";
        }


        /// <summary>
        /// This method will return a list with the Excel Column Names found in the ACTIVE SHEET
        /// </summary>
        /// <param name="pathToFile"></param>
        /// <returns></returns>
        private  List<string> getExcelColumnNames(string pathToFile)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(pathToFile);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;
            int columnCount = worksheet.UsedRange.Columns.Count;

            List<string> columnNames = new List<string>();

            for (int i = 1; i<=columnCount; i++)
            {
                if (worksheet.Cells[1, i].Value2!=null&&worksheet.Cells[1, i].Value2!="")
                {
                    string columnName = worksheet.Columns[i].Address;

                    columnNames.Add(worksheet.Cells[1, i].Value2);
                }
            }

            workbook.Close();

            //Con esto de Marshal se libera de manera completa el objeto desde Interop Services, si no haces esto
            //El objeto sigue en memoria, no lo libera C#
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return columnNames;

        }


        /// <summary>
        /// This method will perform a select column name from the information schema table and return
        /// a list with the column names from the designed table 
        /// </summary>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private  List<string> getDataBaseTableColumns(string databaseConnectionString, string tableName)
        {
            List<string> columnNames = new List<string>();
            //Console.WriteLine("Getting db connection...");
            SqlConnection conexion1 = GetConnectionDevDataBase(databaseConnectionString);
            //OleDbConnection conexion1 = GetOleDbConnection();

            try
            {
                //Console.WriteLine("Opening connection...");
                conexion1.Open();
                //Console.WriteLine("Connection succesful");

                try
                {

                    using (SqlCommand command = new SqlCommand("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '"+tableName+"'", conexion1))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            for (int i = 0; i<reader.FieldCount; i++)
                            {
                                columnNames.Add(reader.GetValue(i).ToString());

                            }

                        }
                    }



                }
                catch (Exception e)
                {


                }

            }
            catch (Exception e)
            {

            }
            finally
            {
                conexion1.Close();
            }

            return columnNames;

        }


        /// <summary>
        /// THIS METHOD CAN WORK ON ITS OWN
        /// The method will perform a query to the excel file selecting everything and upload it to the designed table in the database
        /// using a bulk upload operation
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="excelSheetName"></param>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns>String with "Success" or the error message</returns>
        private  string uploadFileToDatabase(string filePath, string excelSheetName, string databaseConnectionString, string tableName)
        {


            string strConnection = databaseConnectionString;

            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\"", filePath);
            //Create Connection to Excel work book 


            try
            {
                using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                {
                    //Create OleDbCommand to fetch data from Excel 

                    using (OleDbCommand cmd = new OleDbCommand("Select * from ["+excelSheetName+"$]", excelConnection))
                    {

                        excelConnection.Open();


                        using (OleDbDataReader dReader = cmd.ExecuteReader())
                        {
                            //Console.WriteLine(strConnection);
                            using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                            {
                                sqlBulk.BulkCopyTimeout=0;

                                //Give your Destination table name 
                                tableName="["+tableName+"]";
                                sqlBulk.DestinationTableName=tableName;
                                sqlBulk.WriteToServer(dReader);

                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {

                return e.Message;
            }

            return "Success";
        }



        /// <summary>
        /// This method will return the total amount of rows of a given excel that are not null
        /// </summary>
        /// <param name="pathToExcel"></param>
        /// <returns>Integer with the total amount of rows</returns>
        private  int GetTotalAmountOfRowsInExcel(string pathToExcel)
        {
            //Se crea una instancia de una aplicación de Excel
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //False para que no abra la aplicación, sino que lo haga "por atrás"
            myExcel.Visible=false;
            //Aquí usando la instancia de Aplicación de excel, abro el libro mandando como parámetro la ruta a mi archivo
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcel);
            //Después uso una instancia de Worksheet (clase de Interop) para obtener la Hoja actual del archivo Excel
            Microsoft.Office.Interop.Excel.Worksheet worksheet = myExcel.ActiveSheet;
            //En ese worksheet, en la propiedad de Name, tenemos el nombre de la hoja actual, que mando en el query 1 como parámetro
            //Console.WriteLine("WorkSheet.Name: " + worksheet.Name);


            bool exceptionDetected = false;
            int initialRow = 1;
            int totalRows = 0;

            while (!exceptionDetected)
            {
                try
                {
                    if (worksheet.Cells[initialRow, 1].Value2!=null)
                    {
                        //initialRow++;
                        totalRows++;
                        initialRow++;
                        //Console.WriteLine(worksheet.Cells[initialRow, 1].Value2 + " " + initialRow);
                    }
                    else
                    {

                        break;
                    }
                }
                catch (Exception e)
                {
                    exceptionDetected=true;
                }

            }



            //Al finalizar tu proceso debes cerrar tu workbook

            workbook.Close();
            myExcel.Quit();

            //Con esto de Marshal se libera de manera completa el objeto desde Interop Services, si no haces esto
            //El objeto sigue en memoria, no lo libera C#
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();


            return totalRows;

        }



        /// <summary>
        /// Conexión a la base de datos
        /// </summary>
        /// <returns></returns>
        private  SqlConnection GetConnectionDevDataBase(string databaseConnectionString)
        {
            return new SqlConnection(databaseConnectionString);
        }

        /// <summary>
        /// This method returns the name of the active sheet in the excel Document provided
        /// </summary>
        /// <param name="pathToExcelFile"></param>
        /// <returns>String with the name of the Active Sheet</returns>
        private  string getExcelActiveSheetName(string pathToExcelFile)
        {


            //I create an instance of a Microsoft Excel Application
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            //Set visible to false to let the app work "behind".. a.k.a not OPEN the excel application
            myExcel.Visible=false;
            //Using the excel application instance, I open the book with the path to my excel file as parameter
            Microsoft.Office.Interop.Excel.Workbook workbook = myExcel.Workbooks.Open(pathToExcelFile);
            //Later I use a Worksheet Instance to get the actual sheet in the Excel File
            Worksheet worksheet = myExcel.ActiveSheet;
            //In this worksheet, in the Name property we can find the name of the Excel active Worksheet

            string excelSheet = worksheet.Name;

            //We close our workbook at the end of our process
            workbook.Close();

            //As a sideNote, Excel InteropServices don't release the object in memory, there's an instance of Excel still running
            //We use Marchal Final Release to release the object and close the Excel Process
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(myExcel);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return excelSheet;

        }



        /// <summary>
        /// STANDALONE VERSION OF THE UPLOAD METHOD
        /// The method will perform a query to the excel file selecting everything and upload it to the designed table in the database
        /// using a bulk upload operation
        /// It also calls the "Get excel active sheet name" so the user doesn't needs to obtain it later 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="excelSheetName"></param>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns>String with "Success" or the error message</returns>
        public string standaloneUploadFileToDatabase(string filePath, string databaseConnectionString, string tableName)
        {
            string excelSheetName = getExcelActiveSheetName(filePath);

            string strConnection = databaseConnectionString;

            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\"", filePath);
            //Create Connection to Excel work book 


            try
            {
                using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                {
                    //Create OleDbCommand to fetch data from Excel 

                    using (OleDbCommand cmd = new OleDbCommand("Select * from ["+excelSheetName+"$]", excelConnection))
                    {

                        excelConnection.Open();


                        using (OleDbDataReader dReader = cmd.ExecuteReader())
                        {
                            //Console.WriteLine(strConnection);
                            using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                            {
                                sqlBulk.BulkCopyTimeout=0;

                                //Give your Destination table name 
                                tableName="["+tableName+"]";
                                sqlBulk.DestinationTableName=tableName;
                                sqlBulk.WriteToServer(dReader);
                                    
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {

                return e.Message;
            }

            return "Success";
        }



        /// <summary>
        /// STANDALONE VERSION OF THE UPLOAD METHOD
        /// The method will perform a query to the excel file selecting everything and upload it to the designed table in the database
        /// using a bulk upload operation
        /// It takes the sheetName as parameter
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="excelSheetName"></param>
        /// <param name="databaseConnectionString"></param>
        /// <param name="tableName"></param>
        /// <returns>String with "Success" or the error message</returns>
        public string standaloneUploadFileToDatabase(string filePath, string databaseConnectionString, string tableName, string excelSheetName)
        {

            string strConnection = databaseConnectionString;

            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\"", filePath);
            //Create Connection to Excel work book 


            try
            {
                using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                {
                    //Create OleDbCommand to fetch data from Excel 

                    using (OleDbCommand cmd = new OleDbCommand("Select * from ["+excelSheetName+"$]", excelConnection))
                    {

                        excelConnection.Open();


                        using (OleDbDataReader dReader = cmd.ExecuteReader())
                        {
                            //Console.WriteLine(strConnection);
                            using (SqlBulkCopy sqlBulk = new SqlBulkCopy(strConnection))
                            {
                                sqlBulk.BulkCopyTimeout=0;

                                //Give your Destination table name 
                                tableName="["+tableName+"]";
                                sqlBulk.DestinationTableName=tableName;
                                sqlBulk.WriteToServer(dReader);

                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {

                return e.Message;
            }

            return "Success";
        }
               

    }

}




