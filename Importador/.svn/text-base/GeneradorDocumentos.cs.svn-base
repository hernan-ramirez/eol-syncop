using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Data.Common;
using System.Data;
using System.Data.OleDb;
using SincronizadorConsultasProfesionales.Log;

namespace SincronizadorConsultasProfesionales.Importador
{
    public class GeneradorDocumentos
    {
        public DbConnection dbConnection = null;
        public DbCommand dbCommand = null;
        public string path = string.Empty;
        public string db = string.Empty;
        private string _connectionString = string.Empty;
        private bool tipoConsulta;
        private string pathEstilos;

        public GeneradorDocumentos(string db, string connectionString, string pathDocs, bool delDia, string pathEstilos)
        {
            tipoConsulta = delDia;
            path = pathDocs;
            _connectionString = connectionString;
            this.db = db;
            this.pathEstilos = pathEstilos;
            switch (db)
            {
                case "csAccess":
                    dbConnection = new OleDbConnection(connectionString);
                    dbCommand = dbConnection.CreateCommand();
                    dbCommand.CommandText = "select * from public_exportdb";
                    break;
                case "csSql":
                    dbConnection = new SqlConnection(connectionString);
                    string spName = delDia ? "sp_Consultas_Profesionales_Para_Documentos_Del_Dia" : "sp_Consultas_Profesionales_Para_Documentos";
                    dbCommand = new SqlCommand(spName, (SqlConnection)dbConnection);
                    dbCommand.CommandType = CommandType.StoredProcedure;
                    break;
            }
        }

        public void GenerarDocumentosDesdeDB()
        {
            dbConnection.Open();
            RecorrerReader(dbCommand.ExecuteReader());
            dbConnection.Close();
        }

        private void RecorrerReader(DbDataReader dbReader)
        {
            while (dbReader.Read())
            {
                CrearDocumento(dbReader["consulta"].ToString(), dbReader["titulo"].ToString(), dbReader["consultor"].ToString(), dbReader["pregunta"].ToString(), dbReader["respuesta"].ToString(), RefactorFecha((DateTime)dbReader["fecha"]), ObtenerCarpeta(dbReader), this.pathEstilos);
            }
        }

        private string RefactorFecha(DateTime dateTime)
        {
            return string.Concat(dateTime.Day, "/", dateTime.Month, "/", dateTime.Year);
        }

        private string ObtenerCarpeta(DbDataReader dbReader)
        {
            string carpeta = string.Empty;
            switch (db)
            {
                case "csAccess":
                    carpeta = RefactorCarpeta(dbReader["carpeta"].ToString());
                    break;
                case "csSql":
                    carpeta = ObtenerCarpetaPorSql(Convert.ToInt16(dbReader["idArbol"]));
                    break;
                default:
                    break;
            }
            return carpeta;
        }

        private string ObtenerCarpetaPorSql(Int16 idArbol)
        {
            string carpeta = string.Empty;
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                using (SqlCommand categoriasCommand = new SqlCommand("sp_consultar_categoria_arbol", conn))
                {
                    conn.Open();
                    categoriasCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    SqlParameter categoriaParameter = categoriasCommand.CreateParameter();
                    categoriaParameter.ParameterName = "@idArbol";
                    categoriaParameter.Value = idArbol;
                    categoriasCommand.Parameters.Add(categoriaParameter);
                    categoriasCommand.Parameters.Add(new SqlParameter() { ParameterName = "@carpeta", Value = string.Empty });
                    SqlDataReader categoriaReader = categoriasCommand.ExecuteReader();
                    categoriaReader.Read();
                    carpeta = RefactorCarpeta(categoriaReader[0].ToString());
                    carpeta = carpeta.Remove(carpeta.LastIndexOf('/'));
                    conn.Close();
                }
            }
            return carpeta;
        }

        private string RefactorCarpeta(string carpeta)
        {
            switch (db)
            {
                case "csAccess":
                    if (carpeta.Split('/').Length > 2)
                    {
                        string[] vIndice = carpeta.Split('/');
                        string indiceRefactor = vIndice[vIndice.Length - 2].Trim();
                        indiceRefactor += string.Concat("/", vIndice[vIndice.Length - 1].Trim());
                        string[] vIndice2 = vIndice.Take(vIndice.Length - 2).ToArray();
                        vIndice2 = vIndice2.Reverse().ToArray();
                        string indiceRefactor2 = String.Join("/", vIndice2).Trim();
                        indiceRefactor += "/" + indiceRefactor2;
                        carpeta = indiceRefactor;
                    }
                    break;
                case "csSql":
                    carpeta = String.Join("/", carpeta.Split('/').Reverse().ToArray());
                    break;
                default:
                    carpeta = string.Empty;
                    break;
            }
            return CleanSPSpecialCharacters(carpeta);
        }

        public static string CleanSPSpecialCharacters(string inputString)
        {
            if (inputString.Length > 240)
            {
                inputString = inputString.Substring(0, 240);
            }

            Regex replace_Quote = new Regex("[\"]", RegexOptions.Compiled);
            Regex replace_LT = new Regex("[<]", RegexOptions.Compiled);
            Regex replace_GT = new Regex("[>]", RegexOptions.Compiled);
            Regex replace_SC = new Regex("[;]", RegexOptions.Compiled);
            Regex replace_Pipe = new Regex("[|]", RegexOptions.Compiled);
            Regex replace_Tab = new Regex("[\t]", RegexOptions.Compiled);
            Regex replace_Amp = new Regex("[&]", RegexOptions.Compiled);
            Regex replace_Space = new Regex("[ ]", RegexOptions.Compiled);
            RegexOptions options = RegexOptions.None;
            Regex replacesSpaces = new Regex(@"[ ]{2,}", options);
            inputString = replacesSpaces.Replace(inputString, @" ");
            inputString = replace_Quote.Replace(inputString, "'");
            inputString = replace_LT.Replace(inputString, "(");
            inputString = replace_GT.Replace(inputString, ")");
            inputString = replace_SC.Replace(inputString, "");
            inputString = replace_Pipe.Replace(inputString, "");
            inputString = replace_Tab.Replace(inputString, "");
            inputString = replace_Amp.Replace(inputString, "＆");
            inputString = replace_Space.Replace(inputString, " ");

            return inputString.Trim();
        }

        private void CrearDocumento(string consulta, string titulo, string consultor, string pregunta, string respuesta, string fecha, string carpeta, string pathEstilos)
        {
            try
            {
                Documento documento = new Documento(consulta, titulo, consultor, pregunta, respuesta, fecha, carpeta, pathEstilos);
                documento.CrearDocumento(path);
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Importador COP", string.Concat("Error al tratar de crear la siguiente consulta: ", consulta, " Mensaje: ", ex.Message));
            }
        }
    }
}
