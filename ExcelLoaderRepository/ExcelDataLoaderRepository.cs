using CommonModule.Domain.Entity;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoaderRepository
{
    public class ExcelDataLoaderRepository : IExcelDataLoaderRepository
    {
        private string connectionString = "";
        public ExcelDataLoaderRepository()
        {
            connectionString = ConfigurationManager.AppSettings["SQLiteDbPath"]; //@"Data Source=DESKTOP-6CS8HG2\SQLEXPRESS;Initial Catalog=Test;User ID=admin;Password=admin";
        }

        public IList<ExcelDataLoader> GetComoditityData()
        {
            IList<ExcelDataLoader> commoditityDetails = new List<ExcelDataLoader>();
            try
            {
                using (SQLiteConnection con = new SQLiteConnection(@"Data Source=" + connectionString + ";"))
                {
                    con.Open();
                    SQLiteDataAdapter reader;
                    DataTable dt = new DataTable();
                    string sqlQuery = @"Select * from CommodityDetails";

                    SQLiteCommand command = new SQLiteCommand(sqlQuery, con);
                    reader = new SQLiteDataAdapter(command);
                    reader.Fill(dt);
                    commoditityDetails = ConvertDTAToDomain(dt);

                    con.Close();
                }
                return commoditityDetails;
            }
            catch (SQLiteException ex)
            {
                throw ex;
            }
        }

        private IList<ExcelDataLoader> ConvertDTAToDomain(DataTable dtable)
        {
            IList<ExcelDataLoader> commoditityDetails = new List<ExcelDataLoader>();
            foreach(DataRow row in dtable.Rows)
            {
                commoditityDetails.Add(new ExcelDataLoader()
                {
                    AllMonthLimit = Convert.ToDouble(row["AllMonthLimit"].ToString()),
                    AnyOneMonthLimit = Convert.ToDouble(row["AnyOneMonthLimit"].ToString()),
                    CommodityCode = row["CommodityCode"].ToString(),
                    DiminishingBalanceContract = row["DiminishingBalanceContract"].ToString(),
                    ExpiryMonthLimit = Convert.ToDouble(row["ExpiryMonthLimit"].ToString()),
                    ValidFrom = Convert.ToDateTime(row["ValidFrom"].ToString()),
                });
            }

            return commoditityDetails;
        }

        public void SaveExcelToSQL(IList<ExcelDataLoader> excelData)
        {
            try
            {
                using (SQLiteConnection con = new SQLiteConnection(@"Data Source=" + connectionString + ";"))
                {
                    con.Open();
                    string query = @"INSERT INTO CommodityDetails (CommodityCode,DiminishingBalanceContract,ExpiryMonthLimit,AllMonthLimit,
                            AnyOneMonthLimit,ValidFrom) VALUES (@CommodityCode,@DiminishingBalanceContract,@ExpiryMonthLimit,@AllMonthLimit,
                            @AnyOneMonthLimit,@ValidFrom)";
                    foreach (ExcelDataLoader data in excelData)
                    {
                        SQLiteCommand cmd = new SQLiteCommand(query, con);

                        cmd.Parameters.AddWithValue("@CommodityCode", data.CommodityCode);
                        cmd.Parameters.AddWithValue("@DiminishingBalanceContract", data.DiminishingBalanceContract);
                        cmd.Parameters.AddWithValue("@ExpiryMonthLimit", data.ExpiryMonthLimit);
                        cmd.Parameters.AddWithValue("@AllMonthLimit", data.AllMonthLimit);
                        cmd.Parameters.AddWithValue("@AnyOneMonthLimit", data.AnyOneMonthLimit);
                        cmd.Parameters.AddWithValue("@ValidFrom", data.ValidFrom);

                        cmd.ExecuteNonQuery();

                    }


                    //sqlQuery = @"select PolicyId FROM FMEA_DocumentGroup Where (GroupId = @documentGroupId)";
                    //SqlDataReader reader = cmd.ExecuteReader();

                    //lockParams.IsLocked = Convert.ToBoolean(GetByte(reader, "IsLocked"));
                    //lockParams.LockedByUser = reader.GetInt32("LockedBy");
                    //lockParams.LockAliveTimeStamp = reader.GetDateTime("LockAliveTimeStamp");
                    //while (reader.Read())
                    //{
                    //    //do something
                    //}
                    con.Close();
                }
            }
            catch (SQLiteException ex)
            {
                throw ex;
            }
        }
    }
}
