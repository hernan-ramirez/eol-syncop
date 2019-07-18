using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using System.Configuration;
using System.IO;
using SincronizadorConsultasProfesionales.Importador;
using SincronizadorConsultasProfesionales.Log;

namespace SincronizadorConsultasProfesionales
{
    class SincronizadorTimerJob : SPJobDefinition
    {
        public SincronizadorTimerJob()

            : base()
        {

        }

        public SincronizadorTimerJob(string jobName, SPService service, SPServer server, SPJobLockType targetType)

            : base(jobName, service, server, targetType)
        {

        }

        public SincronizadorTimerJob(string jobName, SPWebApplication webApplication)

            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {

            this.Title = "Sincronizador Consultas Profesionales";

        }

        public override void Execute(Guid contentDbId)
        {
            ConnectionStringSettings connectionStringCodigos = System.Configuration.ConfigurationManager.ConnectionStrings["csSql"];
            AppSettingsReader app = new AppSettingsReader();
            string fecha = string.Concat(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString("00"), DateTime.Now.Day.ToString("00"), DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00"), DateTime.Now.Second.ToString("00"), (DateTime.Now.Millisecond).ToString("000"));
            string path = Path.Combine(app.GetValue("PathDocumentos", typeof(string)).ToString(), fecha + "\\");
            string pathEstilos = app.GetValue("PathEstiloXml", typeof(string)).ToString();
            Directory.CreateDirectory(path);
            GeneradorDocumentos generador = new GeneradorDocumentos(
                "csSql",
                connectionStringCodigos.ConnectionString,
                path, true, pathEstilos);
            generador.GenerarDocumentosDesdeDB();
            ImportadorFisicoMeta importador = new ImportadorFisicoMeta(path, app.GetValue("SiteUrl", typeof(string)).ToString(), "csSql", connectionStringCodigos.ConnectionString, fecha);
            importador.ImportarDocx();
            string[] archivos = Directory.GetFiles(path, "*.docx", SearchOption.AllDirectories);
            foreach (string archivo in archivos) 
            {                
                File.Delete(archivo);
            }
        }
    }
}
