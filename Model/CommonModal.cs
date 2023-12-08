using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AImport.Models
{
    public partial class SearchModal
    {
        public string HS_Code { get; set; }
        public string Exporter { get; set; }
        public string Importer { get; set; }
        public string Data_Type { get; set; }
        public string Port { get; set; }
        public string Mode { get; set; }
        public string Country { get; set; }
        public string Product { get; set; }
        public string IEC { get; set; }
        public string Period_From_Month { get; set; }
        public string Period_To_Month { get; set; }
        public string Period_From_Year { get; set; }
        public string Period_To_Year { get; set; }
        public string WhereQuery { get; set; }
        public string OrderByQuery { get; set; }
        public string Database_Name { get; set; }
        public string DBName { get; set; }
        public int Id { get; set; }
        public string TBLName { get; set; }
        public string SB_No { get; set; }
        public string BE_No { get; set; }
        public string ASearch { get; set; }
        public string BE_Type { get; set; }
        public string Product_Description { get; set; }
        public string Supplier_Name { get; set; }
        public string Buyer { get; set; }
        public string Supplier_Country { get; set; }
        public string Column { get; set; }
        public string SearchKeyword { get; set; }
        public List<string> adFilter { get; set; }
        public int page { get; set; }
        public int skip { get; set; }
        public int RowCount { get; set; }
        public bool btnClick { get; set; }
        public string filepath { get; set; }
        public string Month { get; set; }
        public string Year { get; set; }
        public string fileName { get; set; }
        public string dummyfilename { get; set; }
    }

    public partial class tbl_MIS_File_Download
    {
        public int Id { get; set; }
        public Nullable<int> UserId { get; set; }
        public string IEType { get; set; }
        public string FromMonth { get; set; }
        public Nullable<int> FromYear { get; set; }
        public string ToMonth { get; set; }
        public Nullable<int> ToYear { get; set; }
        public string HSCode { get; set; }
        public string Product { get; set; }
        public string IE { get; set; }
        public string BS { get; set; }
        public string Country { get; set; }
        public string Mode { get; set; }
        public string IEC { get; set; }
        public string SB_BE { get; set; }
        public string BEType { get; set; }
        public Nullable<System.DateTime> DateTime { get; set; }
        public string Port { get; set; }
        public Nullable<int> Record_Found { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string FileSize { get; set; }
        public string DummyFileName { get; set; }
        public Nullable<bool> Status { get; set; }
    }

}
