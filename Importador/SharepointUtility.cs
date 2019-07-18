using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Taxonomy;
using Microsoft.SharePoint;
using System.IO;

namespace SincronizadorConsultasProfesionales.Importador
{
    public class SharepointUtility
    {
        /// <summary>
        /// Crea los términos que no existan.
        /// El parámetro "arbol" representa el path del término, siendo el último elemento del vector el término a devolver.
        /// </summary>
        /// <param name="termStore"></param>
        /// <param name="termSet"></param>
        /// <param name="arbol"></param>
        /// <returns></returns>
        public static Term CrearNuevosTerminos(ref TermStore termStore, ref TermSet termSet, string[] arbol, int lcid)
        {
            Term termResult = null;
            string cadena = String.Empty;
            string cadenaAnterior = String.Empty;
            for (int j = 0; j <= arbol.Length - 1; j++)
            {
                cadena += arbol[j];
                termResult = new List<Term>(termSet.GetTerms(arbol[j], false, StringMatchOption.ExactMatch, 100000, false)).Find(c => c.GetPath().ToLower().Equals(cadena.ToLower()));
                if (termResult == null)
                    CrearTermino(ref termStore, ref termSet, arbol, ref termResult, cadenaAnterior, j, lcid);

                cadena += ";";
                cadenaAnterior += arbol[j] + ";";
            }
            return termResult;
        }

        private static void CrearTermino(ref TermStore termStore, ref TermSet termSet, string[] arbol, ref Term termResult, string cadenaAnterior, int j, int lcid)
        {
            if (j == 0)
            {
                termResult = termSet.CreateTerm(arbol[j], lcid);
                termStore.CommitAll();
            }
            else
            {
                Term termPadre = new List<Term>(termSet.GetTerms(arbol[j - 1], false, StringMatchOption.ExactMatch, 100000, false)).Find(t => t.GetPath().ToLower().Equals(cadenaAnterior.Remove(cadenaAnterior.Length - 1, 1).ToLower()));
                termResult = termPadre.CreateTerm(arbol[j], lcid);
                termStore.CommitAll();
            }
        }

        /// <summary>
        /// Sube un archivo cualquiera al destino que se indique
        /// dentro de un portal sharepoint
        /// </summary>
        /// <param name="archivo"></param>
        /// <param name="destinoUrl"></param>
        public static bool SubirArchivo(string archivo, string destinoUrl, SPWeb Web, out string exception)
        {
            try
            {
                exception = string.Empty;
                FileStream fStream = File.Open(archivo, FileMode.Open);
                byte[] contents = new byte[fStream.Length];

                fStream.Read(contents, 0, (int)fStream.Length);
                fStream.Close();

                EnsureParentFolder(Web, destinoUrl);
                Web.Files.Add(destinoUrl, contents);
                return true;
            }
            catch (Exception ex)
            {
                exception = string.Concat("Error al subir el siguiente archivo: ", archivo, " Destino: ", destinoUrl, " Mensaje: ", ex.Message, " Stack Trace: ", ex.StackTrace);
                return false;
            }

        }

        /// <summary>
        /// Se asegura de que las carpetas padres existan.
        /// O sea que toda la ruta o ubicacion exista dentro del site indicado
        /// </summary>
        /// <param name="parentSite"></param>
        /// <param name="destinUrl"></param>
        /// <returns></returns>
        public static string EnsureParentFolder(SPWeb parentSite, string destinUrl)
        {
            destinUrl = parentSite.GetFile(destinUrl).Url;

            int index = destinUrl.LastIndexOf("/");
            string parentFolderUrl = string.Empty;

            if (index > -1)
            {
                parentFolderUrl = destinUrl.Substring(0, index);

                SPFolder parentFolder = parentSite.GetFolder(parentFolderUrl);

                if (!parentFolder.Exists)
                {
                    SPFolder currentFolder = parentSite.RootFolder;

                    foreach (string folder in parentFolderUrl.Split('/'))
                    {
                        currentFolder = currentFolder.SubFolders.Add(folder);
                        try
                        {
                            currentFolder.Item["Title"] = folder;
                            currentFolder.Item.Update();
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            return parentFolderUrl;
        }

        /// <summary>
        /// Devuelve la cantidad de carpetas de una biblioteca
        /// </summary>
        public static int CantidadCarpetas(string web, string lista)
        {
            SPWeb Web = new SPSite(web).OpenWeb();
            return Web.Lists[lista].Folders.Count;
        }

        /// <summary>
        /// Devuelve la cantidad de Archivos de una biblioteca
        /// </summary>
        public static int CantidadArchivos(string web, string lista)
        {
            SPWeb Web = new SPSite(web).OpenWeb();
            return Web.Lists[lista].ItemCount;
        }

        /// <summary>
        /// Setea un el Content Type al documento especificado
        /// </summary>
        /// <param name="listItem"></param>
        /// <param name="contentTypeName"></param>
        public static void CambiarContentType(SPListItem listItem, string contentTypeName)
        {
            SPContentType contentType = listItem.ParentList.ContentTypes[contentTypeName];
            listItem["ContentType"] = contentType.Name;
            listItem["ContentTypeId"] = contentType.Id.ToString();
        }

        /// <summary>
        /// Setea el valor de una columna que no sean del tipo Metadato
        /// (Este método sólo fue probado con el tipo String y el tipo Date)
        /// </summary>
        /// <param name="listItem"></param>
        /// <param name="columnName"></param>
        /// <param name="value"></param>
        public static void SetColumnValue(SPListItem listItem, SPField field, object value)
        {
            listItem[field.Id] = value;
        }

        /// <summary>
        /// Aplica el Metadato al listItem especificado.        
        /// </summary>
        /// <param name="listItem"></param>
        /// <param name="site"></param>
        /// <param name="campo">Field Internal Name</param>
        /// <param name="valor">Term1/Term1Son/Term1GrandSon</param>
        /// <param name="lcid"></param>        
        /// <returns></returns>
        public static void SetMetadata(ref SPListItem listItem, SPSite site, string campo, string valor, int lcid, bool multi)
        {
            TaxonomySession session = new TaxonomySession(site);
            TaxonomyField tagsField = (TaxonomyField)listItem.Fields.GetField(campo);
            TermStore termStore = session.TermStores[tagsField.SspId];
            TermSet termSet = termStore.GetTermSet(tagsField.TermSetId);
            Term termResult;
            if (multi)
            {
                ICollection<Term> termCol = new List<Term>();
                foreach (string term in valor.Split(';'))
                {
                    termResult = CrearNuevosTerminos(ref termStore, ref termSet, valor.Split('/'), lcid);
                    termCol.Add(termResult);
                }
                tagsField.SetFieldValue(listItem, termCol);
            }
            else
            {
                termResult = CrearNuevosTerminos(ref termStore, ref termSet, valor.Split('/'), lcid);
                tagsField.SetFieldValue(listItem, termResult);
            }
            listItem.SystemUpdate();
        }

        /// <summary>
        /// Elimina los acentos del string enviado.
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
        public static string CleanSpecialCharacters(string inputString)
        {
            Regex replace_a_Accents = new Regex("[á|à|ä|â]", RegexOptions.Compiled);
            Regex replace_e_Accents = new Regex("[é|è|ë|ê]", RegexOptions.Compiled);
            Regex replace_i_Accents = new Regex("[í|ì|ï|î]", RegexOptions.Compiled);
            Regex replace_o_Accents = new Regex("[ó|ò|ö|ô]", RegexOptions.Compiled);
            Regex replace_u_Accents = new Regex("[ú|ù|ü|û]", RegexOptions.Compiled);
            Regex replace_A_Accents = new Regex("[Á|À|Ä|Â]", RegexOptions.Compiled);
            Regex replace_E_Accents = new Regex("[É|È|Ë|Ê]", RegexOptions.Compiled);
            Regex replace_I_Accents = new Regex("[Í|Ì|Ï|Î]", RegexOptions.Compiled);
            Regex replace_O_Accents = new Regex("[Ó|Ò|Ö|Ô]", RegexOptions.Compiled);
            Regex replace_U_Accents = new Regex("[Ú|Ù|Ü|Û]", RegexOptions.Compiled);
            inputString = replace_a_Accents.Replace(inputString, "a");
            inputString = replace_e_Accents.Replace(inputString, "e");
            inputString = replace_i_Accents.Replace(inputString, "i");
            inputString = replace_o_Accents.Replace(inputString, "o");
            inputString = replace_u_Accents.Replace(inputString, "u");
            inputString = replace_A_Accents.Replace(inputString, "A");
            inputString = replace_E_Accents.Replace(inputString, "E");
            inputString = replace_I_Accents.Replace(inputString, "I");
            inputString = replace_O_Accents.Replace(inputString, "O");
            inputString = replace_U_Accents.Replace(inputString, "U");
            return inputString;
        }

        /// <summary>
        /// Limpia el string de los caracteres no aceptados por Sharepoint.
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
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
    }
}
