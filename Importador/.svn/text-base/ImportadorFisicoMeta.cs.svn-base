using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.IO;
using System.Data.SqlClient;
using System.Data.Common;
using SincronizadorConsultasProfesionales.Log;
using System.Data;
using System.Data.OleDb;

namespace SincronizadorConsultasProfesionales.Importador
{
    public class ImportadorFisicoMeta
    {
        private string _path;
        private string _url;
        private string _db;
        private string _connectionString;
        private DbConnection _dbConnection;
        private CommandType _cType;
        private List<Campo> campos = new List<Campo>();
        private string _query;
        private string _fecha;

        public ImportadorFisicoMeta(string path, string siteUrl, string db, string connectionString, string fecha)
        {
            _path = path;
            _url = siteUrl;
            _db = db;
            _connectionString = connectionString;
            _fecha = fecha;
            BuilDBConectivity();

        }

        private void BuilDBConectivity()
        {
            switch (_db)
            {
                case "csAccess":
                    _dbConnection = new OleDbConnection(_connectionString);
                    _query = "select * from public_exportdb where consulta = @consulta";
                    _cType = CommandType.Text;
                    break;
                case "csSql":
                    _dbConnection = new SqlConnection(_connectionString);
                    _query = "sp_Consultas_Profesionales__Para_Documentos_Por_IdConsulta";
                    _cType = CommandType.StoredProcedure;
                    break;
            }
        }

        public void ImportarDocx()
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                using (SPSite site = new SPSite(_url))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        try
                        {
                            string[] archivos = Directory.GetFiles(_path, "*.docx", SearchOption.AllDirectories);
                            StreamWriter sw = new StreamWriter(_path + "\\" + _fecha + ".txt", true);
                            foreach (string archivo in archivos)
                            {
                                string consulta = archivo.Split('\\').Last().Split('.').First();
                                string hora = string.Concat(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString("00"), DateTime.Now.Day.ToString("00"), DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00"), DateTime.Now.Second.ToString("00"), (DateTime.Now.Millisecond).ToString("000"), ".docx");
                                string urlNueva = string.Concat(_url, "/", hora);
                                string urlConsulta = string.Concat(_url, "/", consulta, ".docx");

                                if (web.GetFile(urlConsulta).Item == null)
                                {
                                    string exception;
                                    if (SharepointUtility.SubirArchivo(archivo, urlNueva, web, out exception))
                                    {
                                        File.Move(archivo, archivo.Remove(archivo.LastIndexOf('\\')) + '\\' + hora);
                                        //LogProcesados.Crear(LogForErrepar.LogErrepar.TipoDeLog.INFO, hora + " $ " + consulta);
                                        sw.WriteLine(hora + "$" + consulta);
                                        ImportarMeta(site, web, urlNueva, consulta);
                                    }
                                    else
                                        LoggingService.LogError("Importador COP", exception);
                                }
                                else
                                {
                                    sw.Close();
                                    StreamReader sr = new StreamReader(_path + "\\" + _fecha + ".txt");
                                    string reader = sr.ReadToEnd();
                                    string consultaId = reader.Split('\n').First(c => c.Contains(consulta)).Split('$')[1].Trim();
                                    this.ImportarMeta(site, web, urlConsulta, consultaId);
                                    sr.Close();
                                    sw = new StreamWriter(_path + "\\" + _fecha + ".txt", true);
                                }
                            }
                            sw.Close();
                        }
                        catch (Exception ex)
                        {
                            LoggingService.LogError("Importador COP", ex.Message);
                        }
                    }
                }
            });
        }

        private void ImportarMeta(SPSite site, SPWeb web, string url, string consulta)
        {
            string nombre = string.Empty;
            try
            {
                SPListItem listItem = web.GetFile(url).Item;
                nombre = listItem.Name;
                SharepointUtility.CambiarContentType(listItem, "Consulta");
                if (RealizarConsulta(consulta))
                {
                    foreach (Campo campo in this.campos)
                    {
                        bool aplicaValor = _db.Equals("csSql") && campo.Nombre.Equals("copVoces") ? false : true;
                        switch (campo.Tipo)
                        {
                            case "Metadato":
                                if (aplicaValor)
                                    SharepointUtility.SetMetadata(ref listItem, site, campo.Nombre, (string)campo.Valor, 3082, false);
                                break;
                            default:
                                SharepointUtility.SetColumnValue(listItem, listItem.ParentList.Fields.GetFieldByInternalName(campo.Nombre), campo.Valor);
                                break;

                        }
                    }
                    listItem.SystemUpdate();
                    if (_db.Equals("csSql"))
                    {
                        try
                        {
                            if (listItem.File.CheckOutType == SPFile.SPCheckOutType.None || listItem.File.CheckOutType == SPFile.SPCheckOutType.Online)
                            {
                                listItem.File.CheckIn(string.Empty, SPCheckinType.MinorCheckIn);
                            }
                        }
                        catch (Exception ex)
                        {
                            LoggingService.LogError("Importador COP", "Error al procesar el archivo: " + nombre + " Mensaje: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Importador COP", "Error al procesar el archivo: " + nombre + " Mensaje: " + ex.Message);
            }

        }

        private bool RealizarConsulta(string consulta)
        {
            try
            {
                this._dbConnection.Open();
                DbCommand dbCommand = this._dbConnection.CreateCommand();
                dbCommand.CommandText = _query;
                dbCommand.CommandType = _cType;
                DbParameter dbParameter = dbCommand.CreateParameter();
                dbParameter.ParameterName = "@consulta";
                dbParameter.Value = consulta;
                dbCommand.Parameters.Add(dbParameter);
                DbDataReader dbDataReader = dbCommand.ExecuteReader();
                this.campos = new List<Campo>();
                while (dbDataReader.Read())
                {
                    this.campos.Add(new Campo() { Nombre = "copVoces", Tipo = "Metadato", Valor = ObtenerCarpeta(dbDataReader) });
                    this.campos.Add(new Campo() { Nombre = "copFecha", Valor = ((DateTime)dbDataReader["fecha"]), Tipo = "Date" });
                    this.campos.Add(new Campo() { Nombre = "Title", Valor = dbDataReader["titulo"].ToString(), Tipo = "String" });
                    this.campos.Add(new Campo() { Nombre = "copConsultor", Valor = RefactorAutor(dbDataReader["consultor"].ToString()), Tipo = "Metadato" });
                    this.campos.Add(GetObra(campos.First(c => c.Nombre.Equals("copVoces"))));
                }
                this._dbConnection.Close();
                return true;
            }
            catch (Exception ex)
            {
                LoggingService.LogError("Importador COP", "Error en la consulta: " + consulta + " Mensaje: " + ex.Message);
                this._dbConnection.Close();
                return false;
            }
        }

        private Campo GetObra(Campo campo)
        {
            return new Campo() { Nombre = "copObra", Valor = campo.Valor.ToString().Split('/').First(), Tipo = "Metadato" };
        }

        private string ObtenerCarpeta(DbDataReader dbReader)
        {
            string carpeta = string.Empty;
            switch (_db)
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
            switch (_db)
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
            return SharepointUtility.CleanSPSpecialCharacters(carpeta);
        }


        private static string RefactorAutor(string autor)
        {
            return SharepointUtility.CleanSpecialCharacters(autor.Trim()[0].ToString()).Trim() + "/" + autor.Trim();
        }

        public void BorrarDocumentos()
        {
            string[] archivos = Directory.GetFiles(_path, "*.docx", SearchOption.AllDirectories);
            using (SPSite site = new SPSite(this._url))
            {
                using (SPWeb Web = site.OpenWeb())
                {
                    foreach (string archivo in archivos)
                    {
                        string consulta = archivo.Split('\\').Last().Split('.').First();
                        string urlConsulta = string.Concat(_url, "/", consulta, ".docx");
                        SPFile file = Web.GetFile(urlConsulta);
                        if (file.Exists && file.InDocumentLibrary)
                        {
                            SPListItem listItem = file.Item;
                            listItem.Delete();
                        }
                    }
                }
            }
        }
    }
}
