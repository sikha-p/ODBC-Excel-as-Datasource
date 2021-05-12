using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ODBC_Excel_As_Datasource
{
    public class Class1
    {
        public string ConnectAndQuery(string connectionString,string queryType, string query,string csvFilePath)
        {
            OdbcConnection oConn = new OdbcConnection();
            oConn.ConnectionString = connectionString;
            OdbcCommand oComm = new OdbcCommand();
            oComm.Connection = oConn;
            oComm.CommandText = query;
            try
            {
                DataSet ds = new DataSet();
                OdbcDataAdapter oAdapter = new OdbcDataAdapter(oComm);
                oConn.Open();
                oAdapter.Fill(ds);
                Console.WriteLine(ds);
                if(queryType == "SELECT")
                {
                    DataTable dataTable = new DataTable();
                    dataTable = ds.Tables[0];
                    ToCSV(dataTable, csvFilePath);
                }
                return "Sucess";
                //string xml = ds.GetXml();
                //return xml;
            }
            catch (IOException caught) { 
                Console.WriteLine(caught.Message);
                return caught.Message;
            }
            catch (OdbcException caught) {
                Console.WriteLine(caught.Message);
                 return caught.Message;
            }
            finally
            {
                oConn.Close();
            }
        }

        private void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

    }
}
