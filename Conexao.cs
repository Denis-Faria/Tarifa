using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Collections;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text.RegularExpressions;
using System.Net.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using System.Runtime.Serialization;
using MetroFramework.Forms;

namespace tarifa
{
    class Conexao
    {
        public DataSet importarExcel(string caminhoPlanilhaSaldo,string caminhoPlanilhaTarifa)
        {
               
                string strConexao = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"", caminhoPlanilhaSaldo);
                OleDbConnection conn = new OleDbConnection(strConexao);
                conn.Open();

                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataSet ds = new DataSet();

                foreach (DataRow row in dt.Rows)
                {
                    string sheet = row["TABLE_NAME"].ToString();
                
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                    cmd.CommandType = CommandType.Text;

                    DataTable tabelaPlanilha = new DataTable(sheet);
                    ds.Tables.Add(tabelaPlanilha);
                    new OleDbDataAdapter(cmd).Fill(tabelaPlanilha);

                }


            string strConexaoTarifa = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"", caminhoPlanilhaTarifa);
            OleDbConnection connTarifa = new OleDbConnection(strConexaoTarifa);
            connTarifa.Open();
            

            DataTable dtTarifa = connTarifa.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            

            foreach (DataRow row in dtTarifa.Rows)
            {
                string sheet = row["TABLE_NAME"].ToString();
                
                OleDbCommand cmdTarifa = new OleDbCommand("SELECT * FROM [" + sheet + "]", connTarifa);
                cmdTarifa.CommandType = CommandType.Text;

                DataTable tabelaPlanilhaTarifa = new DataTable(sheet);
                ds.Tables.Add(tabelaPlanilhaTarifa);

                new OleDbDataAdapter(cmdTarifa).Fill(tabelaPlanilhaTarifa);

            }



            return ds;
        }

    }
}
