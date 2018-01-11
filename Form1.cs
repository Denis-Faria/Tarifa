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
using System.Security.Permissions;


namespace tarifa
{
    public partial class Form1 : Form
    {
        string caminhoPlanilhaSaldo;
        string caminhoPlanilhaTarifa;
        private MySqlConnection mConn;
        private DataSet mDataSet;
        bool existeDiretorio;
        int inicioPlanilhaSaldo;
        int inicioPlanilhaTarifa;
        double valorTotalTarifa;
        int qtdTarifa = 0;
        int gerenteInvalido = 0;
        double restoResult;
        int cabecalhoInativa = 0;
        int contaBloqueada = 0;
        string caminhoLog;

        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                caminhoPlanilhaSaldo = textBox1.Text;
            }
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog2.FileName;
                caminhoPlanilhaTarifa = textBox2.Text;
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists("C:/Tarifa/"))
                {
                    ArrayList apagadosArquivos = new ArrayList();
                    string[] arquivosApagados = System.IO.Directory.GetFiles("C:/Tarifa/", "*.*");
                    int qtdArquivos = arquivosApagados.Length;
                    for (int u = 0; u < qtdArquivos; u++)
                    {
                        string nomeArquivo = System.IO.Path.GetFileName(arquivosApagados[u]);
                        FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.Read, "C:/Tarifa/" + nomeArquivo);
                        f2.AddPathList(FileIOPermissionAccess.Write | FileIOPermissionAccess.Read, "C:/Tarifa/" + nomeArquivo);
                        System.IO.File.Delete("C:/Tarifa/" + nomeArquivo);
                    }
                }

                if (Directory.Exists("C:/Tarifa/Saldos_Bloqueados/"))
                {
                    ArrayList apagadosArquivos = new ArrayList();
                    string[] arquivosApagados = System.IO.Directory.GetFiles("C:/Tarifa/Saldos_Bloqueados/", "*.*");
                    int qtdArquivos = arquivosApagados.Length;
                    for (int u = 0; u < qtdArquivos; u++)
                    {
                        string nomeArquivo = System.IO.Path.GetFileName(arquivosApagados[u]);
                        FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.Read, "C:/Tarifa/Saldos_Bloqueados/" + nomeArquivo);
                        f2.AddPathList(FileIOPermissionAccess.Write | FileIOPermissionAccess.Read, "C:/Tarifa/Saldos_Bloqueados/" + nomeArquivo);
                        System.IO.File.Delete("C:/Tarifa/Saldos_Bloqueados/" + nomeArquivo);
                    }
                }

                if (Directory.Exists("C:/Tarifa/inativa/"))
                {
                    ArrayList apagadosArquivos = new ArrayList();
                    string[] arquivosApagados = System.IO.Directory.GetFiles("C:/Tarifa/inativa/", "*.*");
                    int qtdArquivos = arquivosApagados.Length;
                    for (int u = 0; u < qtdArquivos; u++)
                    {
                        string nomeArquivo = System.IO.Path.GetFileName(arquivosApagados[u]);
                        FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.Read, "C:/Tarifa/inativa/" + nomeArquivo);
                        f2.AddPathList(FileIOPermissionAccess.Write | FileIOPermissionAccess.Read, "C:/Tarifa/inativa/" + nomeArquivo);
                        System.IO.File.Delete("C:/Tarifa/inativa/" + nomeArquivo);
                    }
                }

                if (Directory.Exists("C:/Tarifa/inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd")))
                {
                    ArrayList apagadosArquivos = new ArrayList();
                    string[] arquivosApagados = System.IO.Directory.GetFiles("C:/Tarifa/inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd"), "*.*");
                    int qtdArquivos = arquivosApagados.Length;
                    for (int u = 0; u < qtdArquivos; u++)
                    {
                        string nomeArquivo = System.IO.Path.GetFileName(arquivosApagados[u]);
                        FileIOPermission f2 = new FileIOPermission(FileIOPermissionAccess.AllAccess, "C:/Tarifa/inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/" + nomeArquivo);
                        f2.AddPathList(FileIOPermissionAccess.Write | FileIOPermissionAccess.Read, "C:/Tarifa/inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/" + nomeArquivo);
                        System.IO.File.Delete("C:/Tarifa/inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/" + nomeArquivo);
                    }
                }



                //Conexão com banco
                mDataSet = new DataSet();
                mConn = new MySqlConnection("Server=10.11.17.30;Database=4030;Uid=root;Pwd=chinchila@acida12244819;");
                //mConn = new MySqlConnection("Server=192.168.0.107;Database=tarifa;Uid=Denis;Pwd=Dtf@4030;");
                mConn.Open();

                if (caminhoPlanilhaSaldo == null || caminhoPlanilhaTarifa == null)
                    MessageBox.Show("Selecione os caminhos dos arquivos");

                if (existeDiretorio == false)
                    System.IO.Directory.CreateDirectory("c:/Tarifa");


                StreamWriter writePrincipal = new StreamWriter(@"C:\Tarifa\Tarifas-" + DateTime.Now.Date.ToString("yyyyMMdd") + ".txt");
                StreamWriter writeLog = new StreamWriter(@"C:\Tarifa\Log-" + DateTime.Now.Date.ToString("yyyyMMdd") + ".txt");

                writePrincipal.WriteLine("0175640300001562SICOOBDIVI5803" + DateTime.Now.Date.ToString("yyyyMMdd") + "                                                                                                                                                                  ");

                writeLog.WriteLine("                                    -------------          ");
                writeLog.WriteLine("                                    |Log Tarifas|          ");
                writeLog.WriteLine("                                    -------------          ");


                Conexao conectaPlanilha = new Conexao();
                DataSet output = conectaPlanilha.importarExcel(caminhoPlanilhaSaldo, caminhoPlanilhaTarifa);

                dataGridView1.DataSource = output.Tables[0];
                dataGridView2.DataSource = output.Tables[1];

                for (int i = 0; i <= output.Tables[1].Rows.Count; i++)
                {
                    string nomeColuna=dataGridView2.Columns[0].HeaderText;
                    
                    if (output.Tables[1].Rows[i][nomeColuna].ToString().Length > 0)
                    {
                        if (!output.Tables[1].Rows[i][nomeColuna].ToString().Contains("Conta") && !output.Tables[1].Rows[i][nomeColuna].ToString().Contains("4030 -") && !output.Tables[1].Rows[i][nomeColuna].ToString().Contains("00 - CCLA"))
                        {
                            inicioPlanilhaTarifa = i;
                            break;
                        }
                    }
                }
                
                for (int j = 0; j <= output.Tables[0].Rows.Count; j++)
                {
                    if (output.Tables[0].Rows[j]["F1"].ToString().Length > 0)
                    {
                        if (!output.Tables[0].Rows[j]["F1"].ToString().Contains("Conta"))
                        {
                            inicioPlanilhaSaldo = j;
                            break;
                        }
                    }
                }

                int sinal = 0;
                double saldoRestante = 0;
                double saldoLimite = 0;
                double saldoBloqueado = 0;
                string contaTarifa = " ";
                string idGerente = "";
                int contaAtiva = 0;

                // /*
                if (checkBox1.Checked.Equals(true))
                {
                    string nomeColuna = dataGridView2.Columns[0].HeaderText;
                    string nomeColunaDescricao = dataGridView2.Columns[7].HeaderText;
                    string nomeColunaValor = dataGridView2.Columns[10].HeaderText;

                    for (int i = inicioPlanilhaTarifa; i < output.Tables[1].Rows.Count; i++)
                    {


                        int linhabranco = output.Tables[1].Rows[i][nomeColuna].ToString().Length +
                                          output.Tables[1].Rows[i][nomeColunaDescricao].ToString().Length +
                                          output.Tables[1].Rows[i][nomeColunaValor].ToString().Length;



                        if (linhabranco > 0)
                        {

                            if (output.Tables[1].Rows[i][nomeColuna].ToString().Contains("TOTAL") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("PA") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("Conta") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("CCO"))
                            {
                                sinal = 0;
                            }
                            else
                            {

                                if (sinal == 0)
                                {
                                    contaTarifa = output.Tables[1].Rows[i][nomeColuna].ToString();

                                    MySqlCommand consultaBloqueada = new MySqlCommand("select numcontacorrente from contasbloqueadas where " +
                                                                    " numcontacorrente='" + contaTarifa.Replace("-", "").Replace(".", "") + "'", mConn);

                                    MySqlDataAdapter conBloqueada = new MySqlDataAdapter(consultaBloqueada);
                                    DataTable tabelaBloqueada = new DataTable();
                                    conBloqueada.Fill(tabelaBloqueada);

                                    if (tabelaBloqueada.Rows.Count > 0)
                                    {

                                        contaBloqueada = 1;
                                    }
                                    else
                                    {
                                        contaBloqueada = 0;
                                    }

                                    if (contaTarifa.Length > 0)
                                    {
                                        MySqlCommand consultaGerente = new MySqlCommand("select a.idgerente from pessoas as a inner join" +
                                                                                                            " contasclientes as b on a.id = b.idcliente where b.numcontacorrente='" + contaTarifa.Replace("-", "").Replace(".", "") + "'" +
                                                                                                            " and a.idgerente!='2743454' ", mConn);
                                        MySqlDataAdapter conGerente = new MySqlDataAdapter(consultaGerente);
                                        DataTable tabelaGerente = new DataTable();
                                        conGerente.Fill(tabelaGerente);

                                        MySqlCommand consultaSituacao = new MySqlCommand("select situacao from contascorrentes where numcontacorrente='" + contaTarifa.ToString().Replace("-", "").Replace(".", "") + "'", mConn);
                                        MySqlDataAdapter conSituacao = new MySqlDataAdapter(consultaSituacao);
                                        DataTable tabelaSituacao = new DataTable();
                                        conSituacao.Fill(tabelaSituacao);

                                        contaAtiva = Convert.ToInt32(tabelaSituacao.Rows[0][0]);

                                        if (tabelaGerente.Rows.Count > 0)
                                        {
                                            idGerente = tabelaGerente.Rows[0][0].ToString();
                                        }
                                    }

                                }

                                double valorTarifa = Convert.ToDouble(output.Tables[1].Rows[i][nomeColunaValor]);
                                string descTarifa = output.Tables[1].Rows[i][nomeColunaDescricao].ToString();

                                if (contaAtiva == 1 && contaBloqueada == 0)
                                {
                                    if (output.Tables[1].Rows[i][nomeColuna].ToString().Length > 0)
                                    {
                                        if (sinal == 0)
                                        {
                                            for (int j = inicioPlanilhaSaldo; j < output.Tables[0].Rows.Count; j++)
                                            {
                                                if (contaTarifa == output.Tables[0].Rows[j]["F1"].ToString())
                                                {

                                                    saldoLimite = Math.Round((Convert.ToDouble(output.Tables[0].Rows[j]["F26"])) + (Convert.ToDouble(output.Tables[0].Rows[j]["F35"])), 2);
                                                    saldoBloqueado = Convert.ToDouble(output.Tables[0].Rows[j]["F29"]);
                                                    saldoRestante = Math.Round(saldoLimite + saldoBloqueado, 2);
                                                    sinal = 1;
                                                    break;
                                                }
                                            }
                                        }

                                        if (valorTarifa <= saldoLimite && sinal == 1)
                                        {

                                            saldoLimite = Math.Round(saldoLimite - valorTarifa, 2);
                                            saldoRestante = Math.Round(saldoRestante - valorTarifa);
                                            writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");

                                            qtdTarifa = qtdTarifa + 1;
                                            valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);
                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                    contaBloqueada = 0;
                                                }
                                            }

                                        }
                                        else if (valorTarifa <= saldoRestante && sinal == 1)
                                        {
                                            if (saldoLimite > 0)
                                            {

                                                double valorSaldoBloqueado = Math.Round(valorTarifa - saldoLimite, 2);

                                                writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");

                                                qtdTarifa = qtdTarifa + 1;
                                                valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);
                                                if (idGerente.Length > 0)
                                                {
                                                    MySqlCommand comm = new MySqlCommand("INSERT INTO saldobloqueado (id_gerente,conta_cliente,saldo_bloqueado,enviada) values('" + idGerente + "','" + contaTarifa + "','" + valorSaldoBloqueado.ToString().Replace(",", ".") + "',0)", mConn);
                                                    comm.ExecuteNonQuery();
                                                }

                                                saldoRestante = Math.Round(saldoRestante - saldoLimite - valorTarifa, 2);
                                                saldoLimite = 0;

                                            }
                                            else
                                            {
                                                double valorSaldoBloqueado = valorTarifa;
                                                if (idGerente.Length > 0)
                                                {
                                                    MySqlCommand comm = new MySqlCommand("INSERT INTO saldobloqueado (id_gerente,conta_cliente,saldo_bloqueado,enviada) values('" + idGerente + "','" + contaTarifa + "','" + valorSaldoBloqueado.ToString().Replace(",", ".") + "',0)", mConn);
                                                    comm.ExecuteNonQuery();
                                                }
                                                writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");

                                                qtdTarifa = qtdTarifa + 1;
                                                valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                                saldoRestante = saldoRestante - valorTarifa;
                                            }



                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                    contaBloqueada = 0;
                                                }
                                            }
                                        }
                                        else
                                        {

                                            MySqlCommand comm = new MySqlCommand("INSERT INTO logtarifa (conta,valor,descricao) values('" + contaTarifa + "','" + valorTarifa.ToString().Replace(",", ".") + "','" + descTarifa.ToString() + "')", mConn);
                                            comm.ExecuteNonQuery();

                                            string msg = ("Tarifa não debitada por falta de saldo");
                                            writeLog.WriteLine("Conta:" + contaTarifa.ToString().PadLeft(10, ' ') + " | Tarifa: " + valorTarifa.ToString("N2").PadLeft(10, ' ') + " | " + msg.ToString().PadLeft(40, ' ') + "|");
                                            writeLog.WriteLine("---------------- | ------------------ | --------------------------------------- |");

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                    contaBloqueada = 0;
                                                }
                                            }
                                        }

                                    }
                                    else if (sinal == 1)
                                    {

                                        if (valorTarifa <= saldoLimite)
                                        {

                                            saldoLimite = Math.Round(saldoLimite - valorTarifa, 2);

                                            writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");

                                            qtdTarifa = qtdTarifa + 1;
                                            valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                            saldoRestante = Math.Round(saldoRestante - valorTarifa, 2);
                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }


                                        }
                                        else if (valorTarifa <= saldoRestante)
                                        {
                                            MySqlCommand consultasaldobloqueado = new MySqlCommand("select saldo_bloqueado from saldobloqueado where conta_cliente ='" + contaTarifa + "' ", mConn);
                                            MySqlDataAdapter conSB = new MySqlDataAdapter(consultasaldobloqueado);
                                            DataTable tabelaSB = new DataTable();
                                            conSB.Fill(tabelaSB);

                                            if (tabelaSB.Rows.Count <= 0)
                                            {

                                                if (saldoLimite > 0)
                                                {
                                                    double valorSaldoBloqueado = Math.Round(valorTarifa - saldoLimite, 2);
                                                    if (idGerente.Length > 0)
                                                    {
                                                        MySqlCommand comm = new MySqlCommand("INSERT INTO saldobloqueado (id_gerente,conta_cliente,saldo_bloqueado,enviada) values('" + idGerente + "','" + contaTarifa + "','" + valorSaldoBloqueado.ToString().Replace(",", ".") + "',0)", mConn);
                                                        comm.ExecuteNonQuery();
                                                    }
                                                    writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                            "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");
                                                    saldoRestante = Math.Round(saldoRestante - valorTarifa);
                                                    saldoLimite = 0;

                                                    qtdTarifa = qtdTarifa + 1;
                                                    valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                                }
                                                else
                                                {
                                                    double valorSaldoBloqueado = valorTarifa;
                                                    if (idGerente.Length > 0)
                                                    {
                                                        MySqlCommand comm = new MySqlCommand("INSERT INTO saldobloqueado (id_gerente,conta_cliente,saldo_bloqueado,enviada) values('" + idGerente + "','" + contaTarifa + "','" + valorSaldoBloqueado.ToString().Replace(",", ".") + "',0)", mConn);
                                                        comm.ExecuteNonQuery();
                                                    }
                                                    writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                            "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");
                                                    saldoRestante = Math.Round(saldoRestante - valorTarifa);

                                                    qtdTarifa = qtdTarifa + 1;
                                                    valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                                }
                                            }

                                            if (tabelaSB.Rows.Count > 0)
                                            {

                                                if (saldoLimite > 0)
                                                {
                                                    double valorSaldoBloqueadoExistente = Convert.ToDouble(tabelaSB.Rows[0][0]);
                                                    double valorSaldoBloqueado = Math.Round((valorTarifa - saldoLimite) + valorSaldoBloqueadoExistente, 2);
                                                    if (idGerente.Length > 0)
                                                    {
                                                        MySqlCommand comm1 = new MySqlCommand("UPDATE saldobloqueado set saldo_bloqueado='" + valorSaldoBloqueado.ToString().Replace(",", ".") + "'where conta_cliente='" + contaTarifa + "'", mConn);
                                                        comm1.ExecuteNonQuery();
                                                    }
                                                    writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                            "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");
                                                    saldoRestante = Math.Round(saldoRestante - valorTarifa);
                                                    saldoLimite = 0;

                                                    qtdTarifa = qtdTarifa + 1;
                                                    valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                                }
                                                else
                                                {
                                                    double valorSaldoBloqueadoExistente = Convert.ToDouble(tabelaSB.Rows[0][0]);
                                                    double valorSaldoBloqueado = Math.Round(valorTarifa + valorSaldoBloqueadoExistente, 2);
                                                    if (idGerente.Length > 0)
                                                    {
                                                        MySqlCommand comm1 = new MySqlCommand("UPDATE saldobloqueado set saldo_bloqueado='" + valorSaldoBloqueado.ToString().Replace(",", ".") + "'where conta_cliente='" + contaTarifa + "'", mConn);
                                                        comm1.ExecuteNonQuery();
                                                    }

                                                    writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                            "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");
                                                    saldoRestante = Math.Round(saldoRestante - valorTarifa);

                                                    qtdTarifa = qtdTarifa + 1;
                                                    valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                                }
                                            }

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            MySqlCommand comm = new MySqlCommand("INSERT INTO logtarifa (conta,valor,descricao) values('" + contaTarifa + "','" + valorTarifa.ToString().Replace(",", ".") + "','" + descTarifa.ToString() + "')", mConn);
                                            comm.ExecuteNonQuery();
                                            string msg = ("Tarifa não debitada por falta de saldo");
                                            writeLog.WriteLine("Conta:" + contaTarifa.ToString().PadLeft(10, ' ') + " | Tarifa: " + valorTarifa.ToString("N2").PadLeft(10, ' ') + " | " + msg.ToString().PadLeft(40, ' ') + "|");
                                            writeLog.WriteLine("---------------- | ------------------ | --------------------------------------- |");

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }
                                        }

                                    }
                                }
                                else if (contaAtiva == 2 && contaTarifa.Length > 0 && contaBloqueada == 0)
                                {

                                    if (idGerente.Length > 0)
                                    {
                                        if (existeDiretorio == false)
                                            System.IO.Directory.CreateDirectory("c:/Tarifa/inativa");

                                        string pathString = "C:/Tarifa/inativa/#" + idGerente + " - CONTAS INATIVAS.txt";
                                        FileStream inat = new FileStream(pathString, FileMode.Append, FileAccess.Write, FileShare.Write);
                                        inat.Close();
                                        StreamWriter inativo = new StreamWriter("C:/Tarifa/inativa/#" + idGerente + " - CONTAS INATIVAS.txt", true, Encoding.ASCII);
                                        string Line = "|Conta: " + contaTarifa.PadLeft(10, ' ') + "|";
                                        inativo.WriteLine(Line);
                                        string Line2 = "------------------- ";
                                        inativo.WriteLine(Line2);
                                        inativo.Close();
                                    }


                                }


                            }
                        }
                        else
                        {
                            sinal = 0;
                        }

                    }


                    //Cria o arquivo de Saldo Bloqueado - INICIO

                    MySqlCommand conIDGerente = new MySqlCommand("select id_gerente from saldobloqueado", mConn);
                    MySqlDataAdapter con4 = new MySqlDataAdapter(conIDGerente);
                    DataTable tabelaIDGerente = new DataTable();
                    con4.Fill(tabelaIDGerente);

                    int count = 1;
                    int m = 1;


                    while (true)
                    {

                        if (count > tabelaIDGerente.Rows.Count) break;

                        MySqlCommand consultaEnviada = new MySqlCommand("select id_gerente,enviada from saldobloqueado where id='" + count + "'", mConn);
                        MySqlDataAdapter con5 = new MySqlDataAdapter(consultaEnviada);
                        DataTable tabelaEnviada = new DataTable();
                        con5.Fill(tabelaEnviada);

                        string gSB = tabelaEnviada.Rows[0][0].ToString();
                        string env = tabelaEnviada.Rows[0][1].ToString();

                        if (env == "0" && gSB != "2743454")//alterar para o id da Louise
                        {

                            double totalSaldo = 0;
                            if (existeDiretorio == false)
                                System.IO.Directory.CreateDirectory("c:/Tarifa/Saldos_Bloqueados");

                            StreamWriter writeSaldobloqueado = new StreamWriter(@"C:\Tarifa\Saldos_Bloqueados\#" + gSB + " - SALDO BLOQUEADO.txt");
                            writeSaldobloqueado.WriteLine("          ------------------------------------          ");
                            writeSaldobloqueado.WriteLine("          | Arquivo de Saldo Bloqueado Usado |          ");
                            writeSaldobloqueado.WriteLine("          ------------------------------------          ");

                            for (m = 1; m <= tabelaIDGerente.Rows.Count; m++)
                            {

                                MySqlCommand consultaEnviada1 = new MySqlCommand("select enviada,id_gerente,saldo_bloqueado,conta_cliente from saldobloqueado where id='" + m + "'", mConn);

                                MySqlDataAdapter con1 = new MySqlDataAdapter(consultaEnviada1);
                                DataTable tabelaEnviada1 = new DataTable();
                                con1.Fill(tabelaEnviada1);

                                int enviada = Convert.ToInt32(tabelaEnviada1.Rows[0][0]);
                                int gerenteSaldoB = Convert.ToInt32(tabelaEnviada1.Rows[0][1]);
                                double saldo_bloqueado = Convert.ToDouble(tabelaEnviada1.Rows[0][2]);
                                string contaCliente = tabelaEnviada1.Rows[0][3].ToString();

                                if (gerenteSaldoB.ToString() == gSB.ToString() && enviada == 0)
                                {
                                    //writeSaldobloqueado.WriteLine("Gerente" + gerenteSaldoB);
                                    totalSaldo = totalSaldo + saldo_bloqueado;
                                    MySqlCommand comm1 = new MySqlCommand("UPDATE saldobloqueado set enviada='1' where id='" + m + "'", mConn);
                                    comm1.ExecuteNonQuery();
                                    writeSaldobloqueado.WriteLine("-----------------------------------------------------");
                                    writeSaldobloqueado.WriteLine("|Conta:" + contaCliente.ToString().PadLeft(10, ' ') + " | Saldo Descontado:" + saldo_bloqueado.ToString("N2").PadLeft(15, ' ') + "|");
                                }

                            }
                            //tabelaSaldoBloqueado.Clear();


                            writeSaldobloqueado.Dispose();
                            writeSaldobloqueado.Close();
                        }
                        count++;
                    }

                    MySqlCommand delete = new MySqlCommand("delete from saldobloqueado", mConn);
                    delete.ExecuteNonQuery();

                    MySqlCommand increment = new MySqlCommand("alter table saldobloqueado auto_increment=1", mConn);
                    increment.ExecuteNonQuery();
                }
                else
                {
                    string nomeColuna = dataGridView2.Columns[0].HeaderText;
                    string nomeColunaDescricao = dataGridView2.Columns[7].HeaderText;
                    string nomeColunaValor= dataGridView2.Columns[10].HeaderText;
                    for (int i = inicioPlanilhaTarifa; i < output.Tables[1].Rows.Count; i++)
                    {
                        
                       
                            int linhabranco = output.Tables[1].Rows[i][nomeColuna].ToString().Length +
                                              output.Tables[1].Rows[i][nomeColunaDescricao].ToString().Length +
                                              output.Tables[1].Rows[i][nomeColunaValor].ToString().Length;
                        
                       


                        if (linhabranco > 0)
                        {

                            if (output.Tables[1].Rows[i][nomeColuna].ToString().Contains("TOTAL") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("PA") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("Conta") ||
                                output.Tables[1].Rows[i][nomeColuna].ToString().Contains("CCO"))
                            {
                                sinal = 0;
                            }
                            else
                            {

                                if (sinal == 0)
                                {
                                    contaTarifa = output.Tables[1].Rows[i][nomeColuna].ToString();
                                    if (contaTarifa.Length > 0)
                                    {

                                        MySqlCommand consultaBloqueada = new MySqlCommand("select numcontacorrente from contasbloqueadas where " +
                                                                    " numcontacorrente='" + contaTarifa.Replace("-", "").Replace(".", "") + "'", mConn);

                                        MySqlDataAdapter conBloqueada = new MySqlDataAdapter(consultaBloqueada);
                                        DataTable tabelaBloqueada = new DataTable();
                                        conBloqueada.Fill(tabelaBloqueada);

                                        if (tabelaBloqueada.Rows.Count > 0)
                                        {
                                            contaBloqueada = 1;
                                        }
                                        else
                                        {
                                            contaBloqueada = 0;
                                        }

                                        MySqlCommand consultaGerente = new MySqlCommand("select a.idgerente from pessoas as a inner join" +
                                                                                                            " contasclientes as b on a.id = b.idcliente where b.numcontacorrente='" + contaTarifa.Replace("-", "").Replace(".", "") + "'" +
                                                                                                            " and a.idgerente!='2743454' ", mConn);
                                        MySqlDataAdapter conGerente = new MySqlDataAdapter(consultaGerente);
                                        DataTable tabelaGerente = new DataTable();
                                        conGerente.Fill(tabelaGerente);

                                        //MySqlCommand consultaSituacao = new MySqlCommand("select situacao from contasclientes where numcontacorrente='" + conta.ToString().Replace("-", "").Replace(".", "") + "'", mConn);
                                        //MySqlDataAdapter conISS = new MySqlDataAdapter(consultaSituacao);

                                        MySqlCommand consultaSituacao = new MySqlCommand("select situacao from contascorrentes where  numcontacorrente='" + contaTarifa.ToString().Replace("-", "").Replace(".", "") + "'", mConn);
                                        MySqlDataAdapter conSituacao = new MySqlDataAdapter(consultaSituacao);
                                        DataTable tabelaSituacao = new DataTable();
                                        conSituacao.Fill(tabelaSituacao);

                                        if (tabelaGerente.Rows.Count > 0)
                                        {
                                            idGerente = tabelaGerente.Rows[0][0].ToString();
                                        }

                                        contaAtiva = Convert.ToInt32(tabelaSituacao.Rows[0][0]);

                                    }

                                }

                                double valorTarifa = Convert.ToDouble(output.Tables[1].Rows[i][nomeColunaValor]);
                                string descTarifa = output.Tables[1].Rows[i][nomeColunaDescricao].ToString();

                                //string nomeColunaValor = dataGridView1.Columns[11].HeaderText;

                                string nomeColunaConta = dataGridView1.Columns[0].HeaderText;
                                string nomeColunaAux26 = dataGridView1.Columns[25].HeaderText;
                                string nomeColunaAux35 = dataGridView1.Columns[34].HeaderText;

                                if (contaAtiva == 1 && contaBloqueada == 0)
                                {
                                    if (output.Tables[1].Rows[i][nomeColuna].ToString().Length > 0)
                                    {
                                        if (sinal == 0)
                                        {
                                            for (int j = inicioPlanilhaSaldo; j < output.Tables[0].Rows.Count; j++)
                                            {
                                                if (contaTarifa == output.Tables[0].Rows[j][nomeColunaConta].ToString())
                                                {

                                                    saldoLimite = Math.Round((Convert.ToDouble(output.Tables[0].Rows[j][nomeColunaAux26])) + (Convert.ToDouble(output.Tables[0].Rows[j][nomeColunaAux35])), 2);

                                                    saldoRestante = Math.Round(saldoLimite + saldoBloqueado, 2);
                                                    sinal = 1;
                                                    break;
                                                }
                                            }
                                        }

                                        if (valorTarifa <= saldoLimite && sinal == 1)
                                        {

                                            saldoLimite = Math.Round(saldoLimite - valorTarifa, 2);
                                            saldoRestante = Math.Round(saldoRestante - valorTarifa);
                                            writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E") + "                  ");

                                            qtdTarifa = qtdTarifa + 1;
                                            valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);
                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            MySqlCommand comm = new MySqlCommand("INSERT INTO logtarifa (conta,valor,descricao) values('" + contaTarifa + "','" + valorTarifa.ToString().Replace(",", ".") + "','" + descTarifa.ToString() + "')", mConn);
                                            comm.ExecuteNonQuery();
                                            string msg = ("Tarifa não debitada por falta de saldo");
                                            writeLog.WriteLine("Conta:" + contaTarifa.ToString().PadLeft(10, ' ') + " | Tarifa: " + valorTarifa.ToString("N2").PadLeft(10, ' ') + " | " + msg.ToString().PadLeft(40, ' ') + "|");
                                            writeLog.WriteLine("---------------- | ------------------ | --------------------------------------- |");

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }
                                        }

                                    }
                                    else if (sinal == 1)
                                    {

                                        if (valorTarifa <= saldoLimite)
                                        {

                                            saldoLimite = Math.Round(saldoLimite - valorTarifa, 2);

                                            writePrincipal.WriteLine("1D" + contaTarifa.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorTarifa.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                        "          000N" + descTarifa.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "Ó").Replace("Ê", "E") + "                  ");

                                            qtdTarifa = qtdTarifa + 1;
                                            valorTotalTarifa = Math.Round(valorTotalTarifa + valorTarifa, 2);

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }


                                        }

                                        else
                                        {
                                            MySqlCommand comm = new MySqlCommand("INSERT INTO logtarifa (conta,valor,descricao) values('" + contaTarifa + "','" + valorTarifa.ToString().Replace(",", ".") + "','" + descTarifa.ToString() + "')", mConn);
                                            comm.ExecuteNonQuery();
                                            string msg = ("Tarifa não debitada por falta de saldo");
                                            writeLog.WriteLine("Conta:" + contaTarifa.ToString().PadLeft(10, ' ') + " | Tarifa: " + valorTarifa.ToString("N2").PadLeft(10, ' ') + " | " + msg.ToString().PadLeft(40, ' ') + "|");
                                            writeLog.WriteLine("---------------- | ------------------ | --------------------------------------- |");

                                            int aux = i + 1;
                                            if (i <= (output.Tables[1].Rows.Count) - 2)
                                            {
                                                if (output.Tables[1].Rows[aux][nomeColuna].ToString().Length > 0)
                                                {
                                                    sinal = 0;
                                                }
                                            }
                                        }

                                    }
                                }
                                else if (contaAtiva == 2 && contaTarifa.Length > 0 && idGerente.Length > 0 && contaBloqueada == 0)
                                {

                                    if (idGerente.Length > 0)
                                    {
                                        if (existeDiretorio == false)
                                            System.IO.Directory.CreateDirectory("c:/Tarifa/Inativa/" + DateTime.Now.Date.ToString("yyyyMMdd"));

                                        string pathString = "C:/Tarifa/inativa/#" + idGerente + " - CONTAS INATIVAS.txt";
                                        FileStream inat = new FileStream(pathString, FileMode.Append, FileAccess.Write, FileShare.Write);
                                        inat.Close();
                                        StreamWriter inativo = new StreamWriter("C:/Tarifa/inativa/#" + idGerente + " - CONTAS INATIVAS.txt", true, Encoding.ASCII);
                                        string Line = "|Conta: " + contaTarifa.PadLeft(10, ' ') + "|";
                                        inativo.WriteLine(Line);
                                        string Line2 = "------------------- ";
                                        inativo.WriteLine(Line2);
                                        inativo.Close();
                                    }


                                }


                            }
                        }
                        else
                        {
                            sinal = 0;
                        }

                    }


                    //Cria o arquivo de Saldo Bloqueado - INICIO

                    MySqlCommand conIDGerente = new MySqlCommand("select id_gerente from saldobloqueado", mConn);
                    MySqlDataAdapter con4 = new MySqlDataAdapter(conIDGerente);
                    DataTable tabelaIDGerente = new DataTable();
                    con4.Fill(tabelaIDGerente);

                    int count = 1;
                    int m = 1;


                    while (true)
                    {

                        if (count > tabelaIDGerente.Rows.Count) break;

                        MySqlCommand consultaEnviada = new MySqlCommand("select id_gerente,enviada from saldobloqueado where id='" + count + "'", mConn);
                        MySqlDataAdapter con5 = new MySqlDataAdapter(consultaEnviada);
                        DataTable tabelaEnviada = new DataTable();
                        con5.Fill(tabelaEnviada);

                        string gSB = tabelaEnviada.Rows[0][0].ToString();
                        string env = tabelaEnviada.Rows[0][1].ToString();

                        if (env == "0" && gSB != "2743454")//alterar para o id da Louise
                        {

                            double totalSaldo = 0;
                            if (existeDiretorio == false)
                                System.IO.Directory.CreateDirectory("c:/Tarifa/Saldos_Bloqueados");

                            StreamWriter writeSaldobloqueado = new StreamWriter(@"C:\Tarifa\Saldos_Bloqueados\#" + gSB + " - SALDO BLOQUEADO.txt");
                            writeSaldobloqueado.WriteLine("          ------------------------------------          ");
                            writeSaldobloqueado.WriteLine("          | Arquivo de Saldo Bloqueado Usado |          ");
                            writeSaldobloqueado.WriteLine("          ------------------------------------          ");

                            for (m = 1; m <= tabelaIDGerente.Rows.Count; m++)
                            {

                                MySqlCommand consultaEnviada1 = new MySqlCommand("select enviada,id_gerente,saldo_bloqueado,conta_cliente from saldobloqueado where id='" + m + "'", mConn);

                                MySqlDataAdapter con1 = new MySqlDataAdapter(consultaEnviada1);
                                DataTable tabelaEnviada1 = new DataTable();
                                con1.Fill(tabelaEnviada1);

                                int enviada = Convert.ToInt32(tabelaEnviada1.Rows[0][0]);
                                int gerenteSaldoB = Convert.ToInt32(tabelaEnviada1.Rows[0][1]);
                                double saldo_bloqueado = Convert.ToDouble(tabelaEnviada1.Rows[0][2]);
                                string contaCliente = tabelaEnviada1.Rows[0][3].ToString();

                                if (gerenteSaldoB.ToString() == gSB.ToString() && enviada == 0)
                                {
                                    //writeSaldobloqueado.WriteLine("Gerente" + gerenteSaldoB);
                                    totalSaldo = totalSaldo + saldo_bloqueado;
                                    MySqlCommand comm1 = new MySqlCommand("UPDATE saldobloqueado set enviada='1' where id='" + m + "'", mConn);
                                    comm1.ExecuteNonQuery();
                                    writeSaldobloqueado.WriteLine("-----------------------------------------------------");
                                    writeSaldobloqueado.WriteLine("|Conta:" + contaCliente.ToString().PadLeft(10, ' ') + " | Saldo Descontado:" + saldo_bloqueado.ToString("N2").PadLeft(15, ' ') + "|");
                                }

                            }



                            writeSaldobloqueado.Dispose();
                            writeSaldobloqueado.Close();
                        }
                        count++;
                    }

                }

                writePrincipal.WriteLine("9" + qtdTarifa.ToString().PadLeft(5, '0') + valorTotalTarifa.ToString("N2").Replace(",", "").Replace(".", "").PadLeft(17, '0') + "0000000000000000000000                                                                                                                                                           ");

                writeLog.Dispose();
                writeLog.Close();
                writePrincipal.Dispose();
                writePrincipal.Close();



            }
            catch (ApplicationException msg)
            {
                MessageBox.Show("Erro no processamento!" + msg);
            }

            //Cria o arquivo de Saldo Bloqueado - FIM

            //Arquivo oficial segunda execução - INICIO

            var wb = new XLWorkbook();//Varíavel para a planilha(Pasta de trabalho)
            var ws = wb.Worksheets.Add("TARIFAS");//Variavel para a planilha dentro do workbook
            //var range = ws.Range("A1:O1");//Seleção do intervalo

            ws.Cell("A1").Value = "Conta";
            ws.Cell("H1").Value = "Tarifa";
            ws.Cell("K1").Value = "Vlr. Passiv. Cob.";
            
            MySqlCommand consulta2 = new MySqlCommand("SELECT * from logtarifa ", mConn);
            MySqlDataAdapter con2 = new MySqlDataAdapter(consulta2);

            DataTable tabela2 = new DataTable();
            con2.Fill(tabela2);
            int countLogTarifa = tabela2.Rows.Count;
            int linha3 = 2;
            int auxConta = 0;


            for (int i = 0; i < countLogTarifa; i++)
            {
                if (auxConta == 0)
                {
                    ws.Cell("A" + linha3.ToString()).Value = tabela2.Rows[i][0].ToString();
                }
                else
                {
                    ws.Cell("A" + linha3.ToString()).Value = "";
                }
                

                ws.Cell("K" + linha3.ToString()).Value = tabela2.Rows[i][1];
                ws.Cell("H" + linha3.ToString()).Value = Convert.ToString(tabela2.Rows[i][2]);

                if (i < countLogTarifa - 1)
                {
                    if (tabela2.Rows[i + 1][0].ToString() == tabela2.Rows[i][0].ToString())
                    {
                        auxConta = 1;
                    }
                    else
                    {
                        auxConta = 0;
                    }
                }
                    linha3++;
         
            }


            //Salvar o arquivo no disco
            wb.SaveAs(@"C:\Tarifa\"+ DateTime.Now.Hour.ToString() +"-"+ DateTime.Now.Minute.ToString()+".xlsx");

            //Liberar objetos da memoria
            ws.Dispose();
            wb.Dispose();
            //Arquivo oficial segunda execução - FIM


            //Cria Email - INICIO

            try
            {
                if (checkBox1.Checked.Equals(true))
                {
                    ArrayList anexos = new ArrayList();
                    if (Directory.Exists("C:/Tarifa/Saldos_Bloqueados/"))
                    {
                        string[] arquivos = System.IO.Directory.GetFiles("C:/Tarifa/Saldos_Bloqueados/", "*.*");
                        int qtdArquivos = arquivos.Length;

                        if (qtdArquivos > 0)
                        {
                            for (int a = 0; a < qtdArquivos; a++)
                            {
                                int posicaohash = arquivos[a].IndexOf("#");
                                int posicaoEspaco = arquivos[a].IndexOf(" ");
                                String codigoGerente = arquivos[a].Substring((posicaohash), (posicaoEspaco - posicaohash));
                                String codigoGerenteSQL = arquivos[a].Substring((posicaohash + 1), (posicaoEspaco - (posicaohash)));//Verificar o calculo


                                MySqlCommand consultaNomeGerente = new MySqlCommand("select nome from pessoas where id ='" + codigoGerenteSQL + "'", mConn);
                                MySqlDataAdapter conNomeGerente = new MySqlDataAdapter(consultaNomeGerente);
                                DataTable tabelaNomeGerente = new DataTable();
                                conNomeGerente.Fill(tabelaNomeGerente);
                                string nomeGerente = tabelaNomeGerente.Rows[0][0].ToString();

                                MySqlCommand consultaEmail = new MySqlCommand("select email from usuarios where nome LIKE '%" + nomeGerente + "%'", mConn);
                                MySqlDataAdapter conEmail = new MySqlDataAdapter(consultaEmail);
                                DataTable tabelaEmail = new DataTable();
                                conEmail.Fill(tabelaEmail);



                                string email = tabelaEmail.Rows[0][0].ToString();
                                // string remetente = "denis.tavares@divicred.com.br";
                                string destinatario = email;
                                string assunto = "Tarifas Pendentes";
                                string enviaMensagem = "Prezados Gerentes," +
                                    "\nSegue para acompanhamento gerencial, relação de Cooperados dos quais tiveram tarifas \ndebitadas mediante saldo bloqueado nesta data." +
                                    "\nEm tempo, com o intuito de preservar a nossa base de cooperados, enviamos também relação\ndos cooperados com conta INATIVA e que possuem" +
                                    " tarifas pendentes de débito \npara providências junto aos mesmos." +
                                    "\n\nCertos da colaboração de todos,\nAtenciosamente,";

                                SmtpClient smtp = new SmtpClient();
                                MailMessage mensagemEmail = new MailMessage();



                                if (File.Exists("C:/Tarifa/Saldos_Bloqueados/" + codigoGerente + " - SALDO BLOQUEADO.txt"))
                                {

                                    anexos.Add("C:/Tarifa/Saldos_Bloqueados/" + codigoGerente + " - SALDO BLOQUEADO.txt");
                                }

                                if (File.Exists("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt"))
                                {
                                    if (existeDiretorio == false)
                                        System.IO.Directory.CreateDirectory("c:/Tarifa/Inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd"));

                                    anexos.Add("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt");

                                }

                                foreach (string item in anexos)
                                {

                                    int sinalCaminhoInativa = 0;

                                    if (item.Substring(0, 18) == "C:/Tarifa/Inativa/")
                                    {
                                        sinalCaminhoInativa = 1;
                                    }
                                    else
                                    {

                                        Attachment anexado = new Attachment(item);
                                        mensagemEmail.Attachments.Add(anexado);
                                    }

                                    if (File.Exists("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt") && sinalCaminhoInativa == 1)
                                    {

                                        System.IO.File.Move("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt", "c:/Tarifa/Inativa/Enviados" +
                                                    DateTime.Now.Date.ToString("yyyyMMdd") + "/" + codigoGerente + "  - CONTAS INATIVAS.txt");

                                        Attachment anexado = new Attachment("C:/Tarifa/Inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/" + codigoGerente + "  - CONTAS INATIVAS.txt");

                                        mensagemEmail.Attachments.Add(anexado);

                                    }
                                    
                                }

                                anexos.Clear();
                               
                                smtp.Host = "smtp.gmail.com";
                                smtp.Port = 587;

                                smtp.EnableSsl = true;
                                smtp.UseDefaultCredentials = false;

                                smtp.Credentials = new System.Net.NetworkCredential("sicoob4030@gmail.com", "drive365");//Login e senha
                                mensagemEmail.From = new MailAddress("sicoob4030@gmail.com");

                                mensagemEmail.Subject = assunto;
                                mensagemEmail.Body = enviaMensagem;

                                mensagemEmail.To.Add(new MailAddress(email));
                                //mensagemEmail.To.Add(new MailAddress("denis.tavares@divicred.com.br"));
                                smtp.Send(message: mensagemEmail);
                                
                            }
                            
                        }


                    }

                    if (Directory.Exists("C:/Tarifa/Inativa/"))
                    {
                        int qtdArquivosInativaEnviados = 0;
                        if (Directory.Exists("C:/Tarifa/Inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/"))
                        {
                            string[] arquivosInativaEnviados = System.IO.Directory.GetFiles("C:/Tarifa/Inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd") + "/", "*.txt");
                            qtdArquivosInativaEnviados = arquivosInativaEnviados.Length;
                        }

                        int qtdArquivosInativa = 0;
                        if (Directory.Exists("C:/Tarifa/Inativa/"))
                        {
                            string[] arquivosInativa = System.IO.Directory.GetFiles("C:/Tarifa/Inativa/", "*.txt");
                            qtdArquivosInativa = arquivosInativa.Length;
                        }



                        int auxInativa = qtdArquivosInativa - qtdArquivosInativaEnviados;

                        if (auxInativa > 0)
                        {
                            string[] arquivosInativa = System.IO.Directory.GetFiles("C:/Tarifa/Inativa/", "*.*");


                            for (int a = 0; a < qtdArquivosInativa; a++)
                            {
                                int posicaohash = arquivosInativa[a].IndexOf("#");
                                int posicaoEspaco = arquivosInativa[a].IndexOf(" ");
                                String codigoGerente = arquivosInativa[a].Substring((posicaohash), (posicaoEspaco - posicaohash));
                                String codigoGerenteSQL = arquivosInativa[a].Substring((posicaohash + 1), (posicaoEspaco - (posicaohash)));//Verificar o calculo


                                MySqlCommand consultaNomeGerente = new MySqlCommand("select nome from pessoas where id ='" + codigoGerenteSQL + "'", mConn);
                                MySqlDataAdapter conNomeGerente = new MySqlDataAdapter(consultaNomeGerente);
                                DataTable tabelaNomeGerente = new DataTable();
                                conNomeGerente.Fill(tabelaNomeGerente);
                                string nomeGerente = tabelaNomeGerente.Rows[0][0].ToString();

                                MySqlCommand consultaEmail = new MySqlCommand("select email from usuarios where nome LIKE '%" + nomeGerente + "%'", mConn);
                                MySqlDataAdapter conEmail = new MySqlDataAdapter(consultaEmail);
                                DataTable tabelaEmail = new DataTable();
                                conEmail.Fill(tabelaEmail);



                                string email = tabelaEmail.Rows[0][0].ToString();
                                string remetente = "denis.tavares@divicred.com.br";
                                string destinatario = email;
                                string assunto = "Tarifas Pendentes";
                                string enviaMensagem = "Prezados Gerentes," +
                                    "\nSegue para acompanhamento gerencial, relação de Cooperados dos quais tiveram tarifas \ndebitadas mediante saldo bloqueado nesta data." +
                                    "\nEm tempo, com o intuito de preservar a nossa base de cooperados, enviamos também relação\ndos cooperados com conta INATIVA e que possuem" +
                                    " tarifas pendentes de débito \npara providências junto aos mesmos." +
                                    "\n\nCertos da colaboração de todos,\nAtenciosamente,";


                                SmtpClient smtp = new SmtpClient();
                                MailMessage mensagemEmail = new MailMessage();

                                if (File.Exists("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt"))
                                {
                                    if (existeDiretorio == false)
                                        System.IO.Directory.CreateDirectory("c:/Tarifa/Inativa/Enviados" + DateTime.Now.Date.ToString("yyyyMMdd"));

                                    anexos.Add("C:/Tarifa/Inativa/" + codigoGerente + " - CONTAS INATIVAS.txt");

                                }

                                foreach (string item in anexos)
                                {
                                    Attachment anexado = new Attachment(item);
                                    mensagemEmail.Attachments.Add(anexado);

                                }
                                anexos.Clear();
                                smtp.Host = "smtp.gmail.com";
                                smtp.Port = 587;

                                smtp.EnableSsl = true;
                                smtp.UseDefaultCredentials = false;

                                //
                                smtp.Credentials = new System.Net.NetworkCredential("sicoob4030@gmail.com", "drive365");//Login e senha
                                mensagemEmail.From = new MailAddress("sicoob4030@gmail.com");

                                mensagemEmail.Subject = assunto;
                                mensagemEmail.Body = enviaMensagem;

                                mensagemEmail.To.Add(new MailAddress(email));
                                //mensagemEmail.To.Add(new MailAddress("denis.tavares@divicred.com.br"));
                                smtp.Send(message: mensagemEmail);


                            }
                        }
                    }
                }
                else
                {

                    ArrayList anexos = new ArrayList();

                    string[] arquivos = System.IO.Directory.GetFiles("C:/Tarifa/Inativa/", "*.*");
                    int qtdArquivos = arquivos.Length;

                    if (qtdArquivos > 0)
                    {
                        for (int a = 0; a < qtdArquivos; a++)
                        {
                            int posicaohash = arquivos[a].IndexOf("#");
                            int posicaoEspaco = arquivos[a].IndexOf(" ");
                            String codigoGerente = arquivos[a].Substring((posicaohash), (posicaoEspaco - posicaohash));
                            String codigoGerenteSQL = arquivos[a].Substring((posicaohash + 1), (posicaoEspaco - (posicaohash)));//Verificar o calculo


                            MySqlCommand consultaNomeGerente = new MySqlCommand("select nome from pessoas where id ='" + codigoGerenteSQL + "'", mConn);
                            MySqlDataAdapter conNomeGerente = new MySqlDataAdapter(consultaNomeGerente);
                            DataTable tabelaNomeGerente = new DataTable();
                            conNomeGerente.Fill(tabelaNomeGerente);
                            string nomeGerente = tabelaNomeGerente.Rows[0][0].ToString();

                            MySqlCommand consultaEmail = new MySqlCommand("select email from usuarios where nome LIKE '%" + nomeGerente + "%'", mConn);
                            MySqlDataAdapter conEmail = new MySqlDataAdapter(consultaEmail);
                            DataTable tabelaEmail = new DataTable();
                            conEmail.Fill(tabelaEmail);



                            string email = tabelaEmail.Rows[0][0].ToString();
                            //string remetente = "denis.tavares@divicred.com.br";
                            string destinatario = email;
                            string assunto = "Tarifas Pendentes";
                            string enviaMensagem = "Prezados Gerentes," +
                                "\nSegue para acompanhamento gerencial, relação de Cooperados dos quais tiveram tarifas \ndebitadas mediante saldo bloqueado nesta data." +
                                "\nEm tempo, com o intuito de preservar a nossa base de cooperados, enviamos também relação\ndos cooperados com conta INATIVA e que possuem" +
                                " tarifas pendentes de débito \npara providências junto aos mesmos." +
                                "\n\nCertos da colaboração de todos,\nAtenciosamente,";

                            SmtpClient smtp = new SmtpClient();
                            MailMessage mensagemEmail = new MailMessage();



                            if (File.Exists("C:/Tarifa/inativa/" + codigoGerente + " - CONTAS INATIVAS.txt"))
                            {

                                anexos.Add("C:/Tarifa/inativa/" + codigoGerente + " - CONTAS INATIVAS.txt");

                            }



                            foreach (string item in anexos)
                            {
                                Attachment anexado = new Attachment(item);
                                mensagemEmail.Attachments.Add(anexado);

                            }
                            
                            anexos.Clear();
                            smtp.Host = "smtp.gmail.com";
                            smtp.Port = 587;

                            smtp.EnableSsl = true;
                            smtp.UseDefaultCredentials = false;

                            smtp.Credentials = new System.Net.NetworkCredential("sicoob4030@gmail.com", "drive365");//Login e senha
                            mensagemEmail.From = new MailAddress("sicoob4030@gmail.com");

                            mensagemEmail.Subject = assunto;
                            mensagemEmail.Body = enviaMensagem;

                            mensagemEmail.To.Add(new MailAddress(email));
                            //mensagemEmail.To.Add(new MailAddress("denis.tavares@divicred.com.br"));
                            smtp.Send(message: mensagemEmail);

                        }
                    }


                }


            }
            
            catch (ApplicationException msg)
            {
                MessageBox.Show("Erro ao enviar email\n" + msg);
            }
            //Cria Email - FIM
            MySqlCommand delLogTarifa = new MySqlCommand("delete from logtarifa", mConn);
            delLogTarifa.ExecuteNonQuery();
            mConn.Close();

            MessageBox.Show("Arquivo Gerado com sucesso!!");
            
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        
    }
}
