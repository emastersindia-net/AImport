using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System;
using System.Configuration;
using AImport.Models;
using ADOX;
using System.Xml.Linq;

namespace AImport
{
    internal class FileImport
    {
        private SqlConnection dbSP = new SqlConnection(ConfigurationManager.ConnectionStrings["TradeIntelligenceDataSearch"].ConnectionString);

        public void ProcessFile()
        {
            try
            {
                var processingData = MIS_List();
                //file process code//
                processfiles(processingData);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
            }

        }

        public bool GetMISDataSearch(SearchModal sm)
        {
            bool DBError = false;
            DataTable dt = new DataTable();
            try
            {
                using (var cmd1 = new SqlCommand("ImportAccessDB", dbSP))
                {
                    if (dbSP.State == ConnectionState.Closed)
                    {
                        dbSP.Open();
                    }

                    cmd1.CommandType = CommandType.StoredProcedure;
                    cmd1.Parameters.AddWithValue("@DataType", sm.Data_Type ?? "");
                    cmd1.Parameters.AddWithValue("@WhereQuery", sm.WhereQuery);
                    cmd1.Parameters.AddWithValue("@Month", sm.Month);
                    cmd1.Parameters.AddWithValue("@Year", sm.Year);
                    cmd1.Parameters.AddWithValue("@filepath", sm.filepath);
                    cmd1.Parameters.AddWithValue("@fileName", sm.finalTblName);
                    cmd1.CommandTimeout = 10000;
                    if (cmd1.ExecuteNonQuery() > 0)
                    {
                        dbSP.Close();
                        DBError = true;
                    }

                    return DBError;
                }

            }
            catch (Exception ex)
            {
                var fileLog = WriteToFile("\r\n Log Created At : " + DateTime.Now + "\r\n" + ex.Message + "processfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                var queryS1 = "UPDATE tbl_MIS_File_Download SET ErrorOccured = @ErrorOccured , FileLog = @FileLog WHERE Id = @Id";
                using (SqlCommand cmd = new SqlCommand(queryS1, dbSP))
                {
                    if (dbSP.State == ConnectionState.Closed)
                    {
                        dbSP.Open();
                    }
                    // Assuming objs.Id is an integer
                    cmd.Parameters.AddWithValue("@Id", sm.Id);
                    cmd.Parameters.AddWithValue("@ErrorOccured", ex.Message + "processfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                    cmd.Parameters.AddWithValue("@FileLog", fileLog);
                    cmd.ExecuteNonQuery();
                    dbSP.Close();
                }

                Console.WriteLine(ex.Message + "\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return false;
            }
            finally
            {
                dbSP.Close();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        private void processfiles(List<tbl_MIS_File_Download> model)
        {
            foreach (var objs1 in model)
            {
                SearchModal objs = new SearchModal();

                objs.Period_From_Month = objs1.FromMonth;
                objs.Period_To_Month = objs1.ToMonth;
                objs.Period_From_Year = objs1.FromYear.ToString();
                objs.Period_To_Year = objs1.ToYear.ToString();
                objs.Data_Type = objs1.IEType == "I" ? "Import" : "Export";
                objs.Exporter = objs1.IEType == "E" ? objs1.IE : "";
                objs.Buyer = objs1.IEType == "E" ? objs1.BS : "";
                objs.Supplier_Name = objs1.IEType == "E" ? objs1.BS : "";
                objs.Importer = objs1.IEType == "E" ? objs1.IE : "";
                objs.Country = objs1.Country;
                objs.BE_No = objs1.IEType == "I" ? objs1.SB_BE : "";
                objs.SB_No = objs1.IEType == "E" ? objs1.SB_BE : "";
                objs.BE_Type = objs1.BEType;
                objs.Port = objs1.Port;
                objs.Mode = objs1.Mode;
                objs.IEC = objs1.IEC;
                objs.HS_Code = objs1.HSCode;
                objs.Product = objs1.Product;
                objs.fileName = objs1.FileName;
                objs.Id = objs1.Id;
                objs.dummyfilename = objs1.DummyFileName;
                var dbID = objs.Id;
                var queryS = "UPDATE tbl_MIS_File_Download SET ProcessStarted = 1 , ProcessStartDate = @date WHERE Id = @Id";
                using (SqlCommand cmd = new SqlCommand(queryS, dbSP))
                {
                    if (dbSP.State == ConnectionState.Closed)
                    {
                        dbSP.Open();
                    }
                    // Assuming objs.Id is an integer
                    cmd.Parameters.AddWithValue("@Id", objs.Id);
                    cmd.Parameters.AddWithValue("@date", DateTime.Now);
                    cmd.ExecuteNonQuery();
                    dbSP.Close();
                }


                if (DTTOACCDB(objs, objs.fileName,dbID.ToString()))
                {
                    if (GenerateZipFile(objs.fileName.Replace(".zip", "").Replace("--", "-")+"-"+dbID.ToString(), objs.filepath))
                    {
                        var storeFilePath = ConfigurationManager.AppSettings["ZipPath"] + "/" + objs.fileName.Replace(".zip", "").Replace("--", "-") + "-" + dbID.ToString() + ".zip";
                        var query = "UPDATE tbl_MIS_File_Download SET Status = 1 ,ProcessStarted = 0 , FilePath = @FilePath , ProcessDate = @DateTime WHERE Id = @Id";
                        using (SqlCommand cmd = new SqlCommand(query, dbSP))
                        {
                            if (dbSP.State == ConnectionState.Closed)
                            {
                                dbSP.Open();
                            }
                            // Assuming objs.Id is an integer
                            cmd.Parameters.AddWithValue("@Id", objs.Id);
                            cmd.Parameters.AddWithValue("@FilePath", storeFilePath);
                            cmd.Parameters.AddWithValue("@DateTime", DateTime.Now);
                            cmd.ExecuteNonQuery();
                            dbSP.Close();
                        }
                    }
                }
                else
                {
                    var queryS1 = "UPDATE tbl_MIS_File_Download SET Status = 0 , ProcessStarted = 0 WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(queryS1, dbSP))
                    {
                        if (dbSP.State == ConnectionState.Closed)
                        {
                            dbSP.Open();
                        }
                        // Assuming objs.Id is an integer
                        cmd.Parameters.AddWithValue("@Id", objs.Id);
                        cmd.ExecuteNonQuery();
                        dbSP.Close();
                    }


                }
            }
        }

        public string WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            return filepath;
        }

        public bool GenerateZipFile(string dbName, string mainZipPath)
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            // Construct the path to the App_Data folder and your Access database file
            string databasePath = System.IO.Path.Combine(baseDirectory, "App_Data", dbName);
            string zipFilePath = databasePath;
            if (System.IO.File.Exists(zipFilePath))
            {
                System.IO.File.Delete(zipFilePath);
            }
            List<string> listoffiles = new List<string>();
            foreach (var filePath in Directory.GetFiles(databasePath))
            {
                listoffiles.Add(filePath);
            }
            if (CreateZipFile(listoffiles, databasePath))
            {
                return true;
            }
            return false;
        }

        public static bool CreateZipFile(List<string> filePath, string zipFilePath)
        {
            try
            {
                zipFilePath = zipFilePath.Split('\\').LastOrDefault();
                var storeFilePath = ConfigurationManager.AppSettings["ZipPath"] + "/" + zipFilePath + ".zip";

                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string databasePath = System.IO.Path.Combine(baseDirectory, "App_Data", zipFilePath);

                using (FileStream zipToCreate = new FileStream(storeFilePath, FileMode.Create))
                {
                    using (ZipArchive archive = new ZipArchive(zipToCreate, ZipArchiveMode.Create))
                    {
                        foreach (var item in filePath)
                        {
                            if (System.IO.File.Exists(item))
                            {
                                string entryName = Path.GetFileName(item);
                                archive.CreateEntryFromFile(item, entryName, CompressionLevel.Optimal);
                                if (System.IO.File.Exists(item))
                                {
                                    System.IO.File.Delete(item);
                                }
                            }
                        }

                        if (System.IO.Directory.Exists(databasePath.Replace(".zip", "")))
                        {
                            System.IO.Directory.Delete(databasePath.Replace(".zip", ""));
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "createaccessfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return false;
            }
        }

        public static bool CreateAccessDatabase(string databasePath)
        {
            try
            {
                // Create a Catalog object
                Catalog catalog = new Catalog();

                // Set the attributes for the new database
                string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Jet OLEDB:Engine Type=5";
                catalog.Create(connectionString);

                // Dispose of the Catalog object
                System.Runtime.InteropServices.Marshal.ReleaseComObject(catalog);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "DataCreate\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return false;
            }
        }

        private SearchModal CreateMISWhereQuery(SearchModal objs)
        {
            var orderByQuery = "";
            string WQ = null;
            objs.WhereQuery = null;
            if (!String.IsNullOrEmpty(objs.HS_Code))
            {
                var List = objs.HS_Code.Split(',');
                foreach (var item in List)
                {
                    WQ = String.IsNullOrEmpty(WQ) ? "(([ITC HS] LIKE '" + item.Trim() + "%')" : WQ + " OR ([ITC HS] LIKE '" + item.Trim() + "%')";
                    //var Values = '"' + item.Trim() + '*' + '"';

                    if ((objs.Data_Type == "Export" && String.IsNullOrEmpty(objs.Product) && String.IsNullOrEmpty(objs.Exporter) && String.IsNullOrEmpty(objs.Buyer) && String.IsNullOrEmpty(objs.SB_No) && String.IsNullOrEmpty(objs.Country) && String.IsNullOrEmpty(objs.IEC) && String.IsNullOrEmpty(objs.Port) && String.IsNullOrEmpty(objs.Mode)) ||
                        (objs.Data_Type == "Import" && String.IsNullOrEmpty(objs.Product) && String.IsNullOrEmpty(objs.Importer) && String.IsNullOrEmpty(objs.Supplier_Name) && String.IsNullOrEmpty(objs.BE_No) && String.IsNullOrEmpty(objs.BE_Type) && String.IsNullOrEmpty(objs.Country) && String.IsNullOrEmpty(objs.IEC) && String.IsNullOrEmpty(objs.Port) && String.IsNullOrEmpty(objs.Mode)))
                    {
                        if (item.Trim().Length == 2 || item.Trim().Length == 4)
                        {
                            objs.OrderByQuery = " Order by [HS Code] , Date ASC";
                        }
                    }
                }
                WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                if (WQ != null)
                {
                    objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                }
                if (String.IsNullOrEmpty(objs.OrderByQuery))
                {
                    objs.OrderByQuery = " Order by [HS Code] , Date ASC";
                }

            }
            WQ = null;
            if (!String.IsNullOrEmpty(objs.Product))
            {
                var List = objs.Product.Split(',');
                foreach (var item in List)
                {
                    if (item.ToLower().Trim() != "books" && item.ToLower().Trim() != "book")
                    {
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS([Item Desc],'" + item.Trim() + "')" : WQ + " OR CONTAINS([Item Desc],'" + item.Trim() + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS([Item Desc],'" + "\"" + item.Trim() + "\"" + "')" : WQ + " OR CONTAINS([Item Desc],'" + "\"" + item.Trim() + "\"" + "')";
                    }
                    //var Value = '"' + "*" + item.Trim() + "*" + '"';
                    //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(ItemDes,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(ItemDes,'" + Value.Trim() + "')";
                }
                WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                if (WQ != null)
                {
                    objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                }
            }



            if (objs.Data_Type == "Export")
            {
                WQ = null;
                if (!String.IsNullOrEmpty(objs.Exporter))
                {
                    var List = objs.Exporter.Split(',');
                    foreach (var item in List)
                    {
                        //var Value = '"' + "*" + item.Trim() + "*" + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(Exporter,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(Exporter,'" + Value.Trim() + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "([Exporter Name] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Exporter Name] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }

                WQ = null;
                if (!String.IsNullOrEmpty(objs.Buyer))
                {
                    var List = objs.Buyer.Split(',');
                    foreach (var item in List)
                    {
                        //var Value = '"' + "*" + item.Trim() + "*" + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(IMPORTER,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(IMPORTER,'" + Value.Trim() + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "([Consignee Name] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Consignee Name] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }


                WQ = null;
                if (!String.IsNullOrEmpty(objs.SB_No))
                {
                    var List = objs.SB_No.Split(',');
                    foreach (var item in List)
                    {
                        WQ = String.IsNullOrEmpty(WQ) ? "([SB No] LIKE '%" + item.Trim() + "%'" : WQ + " OR [SB No] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }

                WQ = null;
                if (!String.IsNullOrEmpty(objs.Country))
                {
                    var List = objs.Country.Split(',');
                    foreach (var item in List)
                    {
                        WQ = String.IsNullOrEmpty(WQ) ? "([Country of Destination] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Country of Destination] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }
            }
            else if (objs.Data_Type == "Import")
            {
                WQ = null;
                if (!String.IsNullOrEmpty(objs.Importer))
                {
                    var List = objs.Importer.Split(',');
                    foreach (var item in List)
                    {
                        //var Value = '"' + "*" + item.Trim() + "*" + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(IMPORTER,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(IMPORTER,'" + Value.Trim() + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "([Importer] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Importer] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }

                WQ = null;
                if (!String.IsNullOrEmpty(objs.Supplier_Name))
                {
                    var List = objs.Supplier_Name.Split(',');
                    foreach (var item in List)
                    {
                        //var Value = '"' + "*" + item.Trim() + "*" + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(IMPORTER,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(IMPORTER,'" + Value.Trim() + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "([Sup Name] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Sup Name] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }


                WQ = null;
                if (!String.IsNullOrEmpty(objs.BE_No))
                {
                    var List = objs.BE_No.Split(',');
                    foreach (var item in List)
                    {
                        WQ = String.IsNullOrEmpty(WQ) ? "([BE No] LIKE '%" + item.Trim() + "%'" : WQ + " OR [BE No] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }

                WQ = null;
                if (!String.IsNullOrEmpty(objs.BE_Type))
                {
                    var List = objs.BE_Type.Split(',');
                    foreach (var item in List)
                    {
                        WQ = String.IsNullOrEmpty(WQ) ? "([Type] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Type] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }

                WQ = null;
                if (!String.IsNullOrEmpty(objs.Country))
                {
                    var List = objs.Country.Split(',');
                    foreach (var item in List)
                    {
                        WQ = String.IsNullOrEmpty(WQ) ? "([Country] LIKE '%" + item.Trim() + "%'" : WQ + " OR [Country] LIKE '%" + item.Trim() + "%'";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }
            }
            WQ = null;
            if (!String.IsNullOrEmpty(objs.IEC))
            {
                var List = objs.IEC.Split(',');
                foreach (var item in List)
                {
                    //var Value = '"' + item.Trim() + '"';
                    //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(IEC,'" + Value.Trim() + "')" : WQ + " OR CONTAINS(IEC,'" + Value.Trim() + "')";
                    WQ = String.IsNullOrEmpty(WQ) ? "(IEC LIKE '%" + item.Trim() + "%'" : WQ + " OR IEC LIKE '%" + item.Trim() + "%'";
                }
                WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                if (WQ != null)
                {
                    objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                }
            }
            WQ = null;
            if (!String.IsNullOrEmpty(objs.Port))
            {
                if (objs.Port.Trim().ToLower() != "all ports" && objs.Port.Trim().ToLower() != "all port" && objs.Port.Trim().ToLower() != "all")
                {
                    var Port_list = objs.Port.Split(',');
                    foreach (var item in Port_list)
                    {
                        var list = item.Split(' ');
                        var aggregatedString = GetPortList(item);

                        //var Value = '"' + item.Trim() + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(Port,'" + Value + "')" : WQ + " OR CONTAINS(Port,'" + Value + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "(([Port] in (" + aggregatedString.Trim() + "))" : WQ + " OR ([Port] in (" + aggregatedString.Trim() + "))";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }
            }
            WQ = null;
            if (!String.IsNullOrEmpty(objs.Mode))
            {
                if (objs.Mode.Trim().ToLower() != "all mode" && objs.Mode.Trim().ToLower() != "all mode" && objs.Mode.Trim().ToLower() != "all")
                {
                    var Mode_List = objs.Mode.Split(',');
                    foreach (var item in Mode_List)
                    {
                        var list = item.Split(' ');
                        //var Value = '"' + item.Trim() + '"';
                        //WQ = String.IsNullOrEmpty(WQ) ? "(CONTAINS(Port,'" + Value + "')" : WQ + " OR CONTAINS(Port,'" + Value + "')";
                        WQ = String.IsNullOrEmpty(WQ) ? "(([Mode] LIKE '%" + item.Trim() + "%')" : WQ + " OR ([Mode] LIKE '%" + item.Trim() + "%')";
                    }
                    WQ = String.IsNullOrEmpty(WQ) ? null : WQ + ")";
                    if (WQ != null)
                    {
                        objs.WhereQuery = String.IsNullOrEmpty(objs.WhereQuery) ? WQ : objs.WhereQuery + " AND " + WQ;
                    }
                }
            }
            objs.WhereQuery = objs.WhereQuery + " " + orderByQuery;
            return objs;
        }

        public string GetPortList(string item)
        {

            dbSP.Open();
            List<string> portNames = new List<string>();

            using (SqlCommand command = new SqlCommand("SELECT PORT_NAME FROM PortLocationCodes WHERE PORT IS NOT NULL AND LOWER(RTRIM(LTRIM(PORT))) = LOWER(@Item)", dbSP))
            {
                command.Parameters.AddWithValue("@Item", item.Trim());

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string portName = reader["PORT_NAME"].ToString();
                        portNames.Add("'" + portName + "'");
                    }
                }
            }
            dbSP.Close();
            return string.Join(",", portNames);
        }

        static string[] GetMonthNames()
        {
            // You can customize the list of months based on your needs
            return new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        }

        public bool DTTOACCDB(SearchModal objs, string zipfilename,string DBID)
        {
            var connectionString1 = "";
            try
            {
                zipfilename = zipfilename.Replace(".zip", "");
                string databasePath2 = "";
                string baseDirectory2 = "";
                // Assuming you have a DataTable named "yourDataTable" with data
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                // Construct the path to the App_Data folder and your Access database file
                string databasePath = System.IO.Path.Combine(baseDirectory, "App_Data", "misTemplate.accdb");

                var createFolderPath = System.IO.Path.Combine(baseDirectory, "App_Data/" + objs.fileName.Replace("--","").Replace(".zip", "")+"-"+ DBID);
                if (!System.IO.Directory.Exists(createFolderPath))
                {
                    System.IO.Directory.CreateDirectory(createFolderPath);
                }

                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databasePath + ";Persist Security Info=False;";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    var WhereQuery = CreateMISWhereQuery(objs).WhereQuery;
                    int fromMonthNumber = DateTime.ParseExact(objs.Period_From_Month, "MMMM", CultureInfo.CurrentCulture).Month;
                    int toMonthNumber = DateTime.ParseExact(objs.Period_To_Month, "MMMM", CultureInfo.CurrentCulture).Month;
                    for (int year = Convert.ToInt32(objs.Period_From_Year); year <= Convert.ToInt32(objs.Period_To_Year); year++)
                    {
                        for (int month = fromMonthNumber; month <= toMonthNumber; month++)
                        {
                            DateTime currentDate = new DateTime(year, month, 1);
                            var smonth = currentDate.ToString("MMMM");
                            // Your logic here, e.g., print the current date
                            DateTime date = new DateTime(year, DateTime.ParseExact(smonth, "MMMM", System.Globalization.CultureInfo.CurrentCulture).Month, 1);
                            string strSQL = "SELECT * FROM [Export] ";
                            OleDbCommand command = new OleDbCommand(strSQL, connection);
                            OleDbDataAdapter da = new OleDbDataAdapter(command);
                            DataTable dt1 = new DataTable();
                            da.Fill(dt1);

                            string strSQL1 = "SELECT * FROM [Import]";
                            OleDbCommand command1 = new OleDbCommand(strSQL1, connection);
                            OleDbDataAdapter da1 = new OleDbDataAdapter(command1);
                            DataTable dt2 = new DataTable();
                            da1.Fill(dt2);
                            var ExporttableName = "";
                            var ImporttableName = "";
                            var finalTblName = "";
                            if (objs.Data_Type == "Export")
                            {
                                ExporttableName = "Export" + smonth+ "" + objs.Period_To_Year;
                                finalTblName = ExporttableName;
                            }
                            else
                            {
                                ExporttableName = "Export";
                            }
                            
                            if (objs.Data_Type == "Import")
                            {
                                ImporttableName = "Import" + smonth + "" + objs.Period_To_Year;
                                finalTblName = ImporttableName;
                            }
                            else
                            {
                                ImporttableName = "Import";
                            }

                            var dbname1 = smonth.Substring(0, 3) + "for" + year;
                            if (objs.dummyfilename == "Export-" || objs.dummyfilename == "Import-")
                            {
                                dbname1 = smonth.Substring(0, 3) + "for" + year;
                            }
                            else
                            {
                                dbname1 = smonth.Substring(0, 3) + "for" + year;
                            }
                            //if (objs.dummyfilename == "Export-" || objs.dummyfilename == "Import-")
                            //{
                            //    dbname1 = objs.Data_Type +""+ smonth+ "" + objs.Period_To_Year;
                            //}
                            objs.finalTblName = finalTblName;
                            var tableName = objs.Data_Type + "" + smonth + "" + objs.Period_To_Year; ;
                            if (createAccessTable(ExporttableName, ImporttableName, dbname1, dt1.Columns, dt2.Columns, zipfilename.Replace("--", "").Replace(".zip", "") + "-" + DBID))
                            {
                                connection.Close();
                                baseDirectory2 = AppDomain.CurrentDomain.BaseDirectory;
                                // Construct the path to the App_Data folder and your Access database file
                                objs.filepath = System.IO.Path.Combine(createFolderPath, dbname1 + ".accdb");
                                objs.Month = smonth;
                                objs.Year = Convert.ToString(year);
                                GC.Collect();
                                if (!GetMISDataSearch(objs))
                                {
                                    string fileDeletePath = AppDomain.CurrentDomain.BaseDirectory;
                                    string databaseDeletePath = System.IO.Path.Combine(baseDirectory, "App_Data", zipfilename, dbname1 + ".accdb");
                                    if (System.IO.File.Exists(databaseDeletePath))
                                    {
                                        System.IO.File.Delete(databaseDeletePath);
                                        if (connection.State != ConnectionState.Closed)
                                        {
                                            connection.Close();
                                        }
                                        GC.Collect();
                                        return false;
                                    }
                                }
                            }
                        }
                    }
                    if(connection.State != ConnectionState.Closed)
                    {
                        connection.Close();
                    }
                    GC.Collect();
                }
                return true;
            }
            catch (Exception ex)
            {
                var fileLog = WriteToFile("\r\n Log Created At : " + DateTime.Now + "\r\n" + ex.Message + "processfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                var queryS1 = "UPDATE tbl_MIS_File_Download SET ErrorOccured = @ErrorOccured , FileLog = @FileLog WHERE Id = @Id";
                using (SqlCommand cmd = new SqlCommand(queryS1, dbSP))
                {
                    if (dbSP.State == ConnectionState.Closed)
                    {
                        dbSP.Open();
                    }
                    // Assuming objs.Id is an integer
                    cmd.Parameters.AddWithValue("@Id", DBID);
                    cmd.Parameters.AddWithValue("@ErrorOccured", ex.Message + "processfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                    cmd.Parameters.AddWithValue("@FileLog", fileLog);
                    cmd.ExecuteNonQuery();
                    dbSP.Close();
                }
                Console.WriteLine(ex.Message + "processfile\r\n" + (ex.InnerException != null ? ex.InnerException.Message : ""));
                return false;
            }

        }
        public bool createAccessTable(string TableName1,string TableName2, string dbname, DataColumnCollection dt, DataColumnCollection d2, string zipfilename)
        {
            var connectionpath = "";
            try
            {
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                // Construct the path to the App_Data folder and your Access database file
                string databasePath = System.IO.Path.Combine(baseDirectory, "App_Data", zipfilename, dbname + ".accdb");
                if (CreateAccessDatabase(databasePath))
                {
                    string createTableStatement = $"CREATE TABLE {TableName1} (";

                    foreach (DataColumn row in dt)
                    {
                        string columnName = "[" + row.ColumnName + "]";
                        string dataType = "varchar(255)";

                        // Append column definition to the CREATE TABLE statement
                        createTableStatement += $"{columnName} {dataType}, ";
                    }

                    // Remove the trailing comma and space
                    createTableStatement = createTableStatement.TrimEnd(',', ' ');
                    createTableStatement += ")";
                    string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databasePath + ";Persist Security Info=False;";
                    connectionpath = connectionString;
                    using (OleDbConnection DynamicSQLConnection = new OleDbConnection(connectionString))
                    {
                        if (DynamicSQLConnection.State == ConnectionState.Closed)
                        {
                            DynamicSQLConnection.Open();
                        }
                        using (OleDbCommand command = new OleDbCommand(createTableStatement, DynamicSQLConnection))
                        {
                            command.ExecuteNonQuery();
                        }
                        DynamicSQLConnection.Close();
                    }

                    createTableStatement = $"CREATE TABLE {TableName2} (";

                    foreach (DataColumn row in d2)
                    {
                        string columnName = "[" + row.ColumnName + "]";
                        string dataType = "varchar(255)";
                        if (row.ColumnName == "Date")
                        {
                            dataType = "datetime"; 
                        }else if (row.ColumnName == "Qty")
                        {
                            dataType = "decimal(18,2)";
                        }
                        else if (row.ColumnName == "Rate(INR)")
                        {
                            dataType = "decimal(18,2)";
                        }
                       

                        // Append column definition to the CREATE TABLE statement
                        createTableStatement += $"{columnName} {dataType}, ";
                    }

                    // Remove the trailing comma and space
                    createTableStatement = createTableStatement.TrimEnd(',', ' ');
                    createTableStatement += ")";
                    connectionpath = connectionString;
                    using (OleDbConnection DynamicSQLConnection = new OleDbConnection(connectionString))
                    {
                        if (DynamicSQLConnection.State == ConnectionState.Closed)
                        {
                            DynamicSQLConnection.Open();
                        }
                        using (OleDbCommand command = new OleDbCommand(createTableStatement, DynamicSQLConnection))
                        {
                            command.ExecuteNonQuery();
                        }
                        DynamicSQLConnection.Close();
                    }
                    return true;
                }
                return false;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "createaccesstable\r\n" +(ex.InnerException!=null?ex.InnerException.Message:""));
                return false;
            }


        }

        public List<tbl_MIS_File_Download> MIS_List()
        {
            List<tbl_MIS_File_Download> obj = new List<tbl_MIS_File_Download>();
            if (dbSP.State == ConnectionState.Closed)
            {
                dbSP.Open();
            }
            var query = "Select * from tbl_MIS_File_Download where Status = 0 OR Status is null or Status = ''";
            using (SqlCommand cmd = new SqlCommand(query, dbSP))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        tbl_MIS_File_Download obj1 = new tbl_MIS_File_Download()
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            IEType = reader["IEType"].ToString(),
                            FromMonth = reader["FromMonth"].ToString(),
                            FromYear = reader["FromYear"] == DBNull.Value ? (int?)null : Convert.ToInt32(reader["FromYear"]),
                            ToMonth = reader["ToMonth"].ToString(),
                            ToYear = reader["ToYear"] == DBNull.Value ? (int?)null : Convert.ToInt32(reader["ToYear"]),
                            HSCode = reader["HSCode"].ToString(),
                            Product = reader["Product"].ToString(),
                            IE = reader["IE"].ToString(),
                            BS = reader["BS"].ToString(),
                            Country = reader["Country"].ToString(),
                            Mode = reader["Mode"].ToString(),
                            IEC = reader["IEC"].ToString(),
                            SB_BE = reader["SB/BE"].ToString(),
                            BEType = reader["BEType"].ToString(),
                            Port = reader["Port"].ToString(),
                            FileName = reader["FileName"].ToString(),
                            FilePath = reader["FilePath"].ToString(),
                            DummyFileName = reader["DummyFileName"].ToString(),
                            Status = reader["Status"] == DBNull.Value ? (bool?)null : Convert.ToBoolean(reader["Status"])
                        };
                        obj.Add(obj1);
                    }
                }
            }
            return obj;
        }

    }
}