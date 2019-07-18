using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System.IO;
using HtmlAgilityPack;

namespace SincronizadorConsultasProfesionales.Importador
{
    public class Documento
    {
        private string _consulta;
        private string _titulo;
        private string _consultor;
        private string _pregunta;
        private string _respuesta;
        private string _fecha;
        private string _carpeta;
        private string _pathEstilos;

        public Documento(string consulta, string titulo, string consultor, string pregunta, string respuesta, string fecha, string carpeta, string pathEstilos)
        {
            this._consulta = consulta;
            this._titulo = titulo;
            this._consultor = consultor;
            this._fecha = fecha;
            this._carpeta = carpeta;
            this._pathEstilos = pathEstilos;
            HtmlDocument preguntaHtml = new HtmlDocument();
            preguntaHtml.LoadHtml(pregunta);
            this._pregunta = StripHTML(preguntaHtml.DocumentNode.OuterHtml);
            HtmlDocument respuestaHtml = new HtmlDocument();
            respuestaHtml.LoadHtml(respuesta);
            this._respuesta = StripHTML(respuestaHtml.DocumentNode.OuterHtml);
        }

        public Documento() { }

        private static bool EsParseable(HtmlDocument hdoc)
        {
            try
            {
                XDocument x = ToXMLDocument(hdoc);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static XDocument ToXMLDocument(HtmlDocument doc)
        {
            using (StringWriter strWriter = new StringWriter())
            {
                doc.OptionOutputAsXml = true;
                doc.Save(strWriter);
                return XDocument.Parse(strWriter.GetStringBuilder().ToString());
            }
        }

        public void CrearDocumento(string pathDocumentos)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(pathDocumentos + this._consulta + ".docx", WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // Create the document structure and add some text.
                mainPart.Document = new Document();

                //Carga de estilos
                StyleDefinitionsPart styleDefinitionsPart =
                    mainPart.AddNewPart<StyleDefinitionsPart>();
                //récupération du template de style
                FileStream stylesTemplate =
                    new FileStream(this._pathEstilos, FileMode.Open, FileAccess.Read);
                styleDefinitionsPart.FeedData(stylesTemplate);
                styleDefinitionsPart.Styles.Save();

                Body body = new Body();
                mainPart.Document.Append(body);
                this.GetBodyContent(mainPart.Document.Body);
                // Save changes to the main document part.                     

                //HeaderPart hp;
                SectionProperties sectPr;

                //mainPart.Document.Append(GetHeaderSectionProperties(mainPart, out hp, out sectPr));

                //mainPart.Document.Save();
                //hp.Header = GetHeader(this._titulo);
                //hp.Header.Save();

                FooterPart fp;
                //SectionProperties fSectPr;
                mainPart.Document.Append(GetFooterSectionProperties(mainPart, out fp, out sectPr));
                mainPart.Document.Save();
                fp.Footer = GetFooter();
                fp.Footer.Save();
                wordDocument.Close();
            }
        }

        private static SectionProperties GetHeaderSectionProperties(MainDocumentPart mainPart, out HeaderPart hp, out SectionProperties sectPr)
        {
            hp = mainPart.AddNewPart<HeaderPart>();
            string relId = mainPart.GetIdOfPart(hp);
            sectPr = new SectionProperties();
            HeaderReference headerReference = new HeaderReference();
            headerReference.Id = relId;
            sectPr.Append(headerReference);
            return sectPr;
        }

        private void GetBodyContent(Body body)
        {
            body.Append(GetTable(this._consultor, this._fecha, this._titulo, this._carpeta));
            body.Append(GetParagraph("", true, 4));//Línea separadora
            body.Append(GetParagraph("", false, 0));//Párrafo Vacío

            #region Appendeo Pregunta
            body.Append(TitleParagraphStyle("PREGUNTA"));
            //if (this._xPregunta == null && !string.IsNullOrEmpty(this._pregunta))
            body.Append(TextParagraphStyle(this._pregunta));
            //else
            //  body.Append(GetParagraphByXDoc(this._xPregunta));
            #endregion

            body.Append(GetParagraph("", false, 0));

            #region Appendeo Respuesta
            body.Append(TitleParagraphStyle("RESPUESTA"));
            //if (this._xRespuesta == null && !string.IsNullOrEmpty(this._respuesta))
            body.Append(TextParagraphStyle(this._respuesta));
            //else
            //  body.Append(GetParagraphByXDoc(this._xRespuesta));
            #endregion

        }

        private static ParagraphProperties GetParagraphProperties(UInt32 size)
        {
            //création de propriétés pour le paragraphe
            ParagraphProperties titleProperties = new ParagraphProperties();
            //on utilise le style Title de word 2007 pour ce paragraphe
            BottomBorder bb = new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size, Space = 1, Color = "auto" };
            ParagraphBorders pb = new ParagraphBorders { BottomBorder = bb };
            titleProperties.ParagraphBorders = pb;
            return titleProperties;
        }

        private static Table GetTable(string consultor, string fecha, string titulo, string carpeta)
        {
            Table table = new Table();
            TableProperties tblPr = new TableProperties();

            //**********************Armado de estilo de tabla según una doctrina********************************//
            TableWidth tw = new TableWidth() { Width = "5000", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Pct) };
            tblPr.Append(tw);
            TableCellSpacing tcs = new TableCellSpacing() { Width = "0", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) };
            tblPr.Append(tcs);
            TableCellMarginDefault tcm = new TableCellMarginDefault
            {
                TopMargin = new TopMargin() { Width = "60", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                //LeftMargin = new LeftMargin() { Width = "60", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                //RightMargin = new RightMargin() { Width = "60", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) },
                BottomMargin = new BottomMargin() { Width = "60", Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa) }
            };
            tblPr.Append(tcm);
            TableLook tl = new TableLook()
            {
                Val = new HexBinaryValue() { Value = "04A0" },
                FirstRow = new OnOffValue() { Value = true },
                LastRow = new OnOffValue() { Value = false },
                FirstColumn = new OnOffValue() { Value = true },
                LastColumn = new OnOffValue() { Value = false },
                NoHorizontalBand = new OnOffValue() { Value = false },
                NoVerticalBand = new OnOffValue() { Value = true }
            };
            tblPr.Append(tl);
            table.Append(tblPr);
            TableGrid tg = new TableGrid();
            GridColumn gc1 = new GridColumn() { Width = "2844" };
            GridColumn gc2 = new GridColumn() { Width = "6636" };
            tg.Append(gc1);
            tg.Append(gc2);
            table.Append(tg);
            //**********************Fin de Armado de estilo de tabla según una doctrina*************************//            

            List<OpenXmlElement> rows = new List<OpenXmlElement>();
            //Preparo lista de columnas para el consultor
            List<OpenXmlElement> cells = new List<OpenXmlElement>();

            //Preparo lista de columnas para el Título
            TableCell tcTitulo = new TableCell(TableParagraphStyle("TÍTULO:"));
            cells.Add(tcTitulo);
            TableCell tcTituloValor = new TableCell(TableParagraphStyle(titulo));
            cells.Add(tcTituloValor);

            rows.Add(new TableRow(cells));

            //Preparo lista de columnas para la fecha
            cells = new List<OpenXmlElement>();

            TableCell tcAutores = new TableCell(TableParagraphStyle("AUTOR/ES:"));
            cells.Add(tcAutores);
            TableCell tcConsultor = new TableCell(TableParagraphStyle(consultor));
            cells.Add(tcConsultor);
            rows.Add(new TableRow(cells));

            //Preparo lista de columnas para la fecha
            cells = new List<OpenXmlElement>();

            TableCell tcFecha = new TableCell(TableParagraphStyle("FECHA:"));
            cells.Add(tcFecha);
            TableCell tcFechaValor = new TableCell(TableParagraphStyle(fecha));
            cells.Add(tcFechaValor);
            rows.Add(new TableRow(cells));

            //Preparo lista de columnas para las Voces
            cells = new List<OpenXmlElement>();

            TableCell tcVoces = new TableCell(TableParagraphStyle("TEMA:"));
            cells.Add(tcVoces);
            TableCell tcCarpeta = new TableCell(TableParagraphStyle(carpeta));
            cells.Add(tcCarpeta);
            rows.Add(new TableRow(cells));


            table.Append(rows);

            return table;
        }

        private static Paragraph TableParagraphStyle(string texto)
        {
            /*
             <w:p w:rsidR="004F6930"
                     w:rsidRDefault="00373D15">
                  <w:pPr>
                    <w:pStyle w:val="rotulonovedades"/>
                    <w:rPr>
                      <w:rFonts w:cs="Arial"/>
                    </w:rPr>
                  </w:pPr>
                  <w:bookmarkStart w:id="0"
                                   w:name="_GoBack"/>
                  <w:bookmarkEnd w:id="0"/>
                  <w:r>
                    <w:rPr>
                      <w:rStyle w:val="rotulo1"/>
                      <w:rFonts w:cs="Arial"/>
                    </w:rPr>
                    <w:t>TÍTULO:</w:t>
                  </w:r>
                  <w:r>
                    <w:rPr>
                      <w:rFonts w:cs="Arial"/>
                    </w:rPr>
                    <w:t xml:space="preserve"> </w:t>
                  </w:r>
                </w:p>
             */
            Paragraph p = new Paragraph();
            ParagraphProperties pPr = new ParagraphProperties();
            ParagraphStyleId pStyle = new ParagraphStyleId() { Val = "rotulonovedades" };
            pPr.Append(pStyle);
            ParagraphMarkRunProperties rPrParagraphProperties = new ParagraphMarkRunProperties();
            RunFonts rFontsParagraph = new RunFonts() { ComplexScript = "Arial" };
            rPrParagraphProperties.Append(rFontsParagraph);
            pPr.Append(rPrParagraphProperties);
            p.Append(pPr);
            Run r = new Run();
            RunProperties rpr = new RunProperties();
            RunStyle rStyle = new RunStyle() { Val = "rotulo1" };
            RunFonts rPrRun = new RunFonts() { ComplexScript = "Arial" };
            rpr.Append(rStyle, rPrRun);
            r.Append(rpr);
            Text text = new Text() { Text = texto };
            r.Append(text);
            p.Append(r);
            Run runEmpty = new Run();
            RunProperties rpEmpty = new RunProperties();
            RunFonts rfEmpty = new RunFonts() { ComplexScript = "Arial" };
            rpEmpty.Append(rfEmpty);
            runEmpty.Append(rpEmpty);
            Text textEmpty = new Text() { Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve) };
            runEmpty.Append(textEmpty);
            p.Append(runEmpty);
            return p;
        }

        private static Paragraph TitleParagraphStyle(string texto)
        {
            //  <w:p w14:paraId="2F58E943"
            //     w14:textId="77777777"
            //     w:rsidR="00F653E2"
            //     w:rsidRPr="00CE2347"
            //     w:rsidRDefault="006B06AF">
            //  <w:pPr>
            //    <w:pStyle w:val="titulocp11craya"/>
            //    <w:divId w:val="1202479076"/>
            //    <w:rPr>
            //      <w:rFonts w:cs="Arial"/>
            //      <w:lang w:val="es-ES"/>
            //    </w:rPr>
            //  </w:pPr>
            //  <w:r w:rsidRPr="00CE2347">
            //    <w:rPr>
            //      <w:rFonts w:cs="Arial"/>
            //      <w:lang w:val="es-ES"/>
            //    </w:rPr>
            //    <w:t>INSCRIPCIÓN</w:t>
            //  </w:r>
            //</w:p>
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            ParagraphStyleId ps = new ParagraphStyleId() { Val = "titulocp11craya" };
            pp.Append(ps);
            ParagraphMarkRunProperties prpr = new ParagraphMarkRunProperties();
            RunFonts prf = new RunFonts() { ComplexScript = "Arial" };
            Languages pLang = new Languages() { Val = "es-ES" };
            prpr.Append(prf);
            prpr.Append(pLang);
            pp.Append(prpr);
            p.Append(pp);
            Run r = new Run();
            RunProperties rpr = new RunProperties();
            RunFonts rFonts = new RunFonts() { ComplexScript = "Arial" };
            Languages lang = new Languages() { Val = "es-ES" };
            rpr.Append(rFonts);
            rpr.Append(lang);
            r.Append(rpr);
            Text text = new Text() { Text = texto };
            r.Append(text);
            p.Append(r);
            return p;
        }

        private static List<OpenXmlElement> TextParagraphStyle(string texto)
        {
            //  <w:p w14:paraId="2F58E944"
            //     w14:textId="77777777"
            //     w:rsidR="00F653E2"
            //     w:rsidRPr="00CE2347"
            //     w:rsidRDefault="006B06AF">
            //  <w:pPr>
            //    <w:pStyle w:val="sangrianovedades"/>
            //    <w:divId w:val="1202479076"/>
            //    <w:rPr>
            //      <w:rFonts w:cs="Arial"/>
            //      <w:lang w:val="es-ES"/>
            //    </w:rPr>
            //  </w:pPr>
            //  <w:r w:rsidRPr="00CE2347">
            //    <w:rPr>
            //      <w:rFonts w:cs="Arial"/>
            //      <w:lang w:val="es-ES"/>
            //    </w:rPr>
            //    <w:t>La administración de consorcios no puede ejercerse a título oneroso ni gratuito sin la previa inscripción en el Registro Público de Administradores de Consorcios de Propiedad Horizontal.</w:t>
            //  </w:r>
            //</w:p>
            List<OpenXmlElement> parrafos = new List<OpenXmlElement>();
            foreach (string linea in texto.Split('\r'))
            {
                Paragraph p = new Paragraph();
                ParagraphProperties pp = new ParagraphProperties();
                ParagraphStyleId ps = new ParagraphStyleId() { Val = "sangrianovedades" };
                pp.Append(ps);
                ParagraphMarkRunProperties prpr = new ParagraphMarkRunProperties();
                RunFonts prf = new RunFonts() { ComplexScript = "Arial" };
                Languages pLang = new Languages() { Val = "es-ES" };
                prpr.Append(prf);
                prpr.Append(pLang);
                pp.Append(prpr);
                p.Append(pp);
                Run r = new Run();
                RunProperties rpr = new RunProperties();
                RunFonts rFonts = new RunFonts() { ComplexScript = "Arial" };
                Languages lang = new Languages() { Val = "es-ES" };
                rpr.Append(rFonts);
                rpr.Append(lang);
                r.Append(rpr);
                Text text = new Text() { Text = linea };
                r.Append(text);
                p.Append(r);
                parrafos.Add(p);
            }
            return parrafos;
        }

        private static Paragraph GetParagraph(string texto, bool paragraphProperties, uint size)
        {
            Paragraph paragraph = new Paragraph();
            Run run_paragraph = new Run();
            // we want to put that text into the output document 
            Text text_paragraph = new Text(texto);
            //Append elements appropriately. 
            run_paragraph.Append(text_paragraph);
            if (paragraphProperties)
                paragraph.Append(GetParagraphProperties(size));
            paragraph.Append(run_paragraph);
            return paragraph;
        }

        private static List<Paragraph> GetParagraphByXDoc(XDocument xText)
        {
            List<Paragraph> parrafos = new List<Paragraph>();
            foreach (XElement element in xText.Descendants("span").Elements())
            {
                //parrafos.Add(TextParagraphStyle(element.Value));
            }
            return parrafos;
        }

        private static Header GetHeader(string titulo)
        {
            Header h = new Header();
            Paragraph p = new Paragraph();
            Run r = new Run();
            Text t = new Text();
            //t.Text = titulo;
            t.Text = string.Empty;
            r.Append(t);
            p.Append(r);
            h.Append(p);
            return h;
        }

        private static SectionProperties GetFooterSectionProperties(MainDocumentPart mainPart, out FooterPart fp, out SectionProperties sectPr)
        {
            fp = mainPart.AddNewPart<FooterPart>();
            string relId = mainPart.GetIdOfPart(fp);
            sectPr = new SectionProperties();
            FooterReference footerReference = new FooterReference();
            footerReference.Id = relId;
            sectPr.Append(footerReference);
            return sectPr;
        }

        private static Footer GetFooter()
        {
            //      <w:ftr mc:Ignorable="w14 wp14"
            //  <w:p w14:paraId="2F58EA34"
            //       w14:textId="77777777"
            //       w:rsidR="006B06AF"
            //       w:rsidRPr="006B06AF"
            //       w:rsidRDefault="006B06AF"
            //       w:rsidP="006B06AF">
            //    <w:pPr>
            //      <w:pStyle w:val="Piedepgina"/>
            //      <w:pBdr>
            //        <w:top w:val="single"
            //               w:sz="4"
            //               w:space="1"
            //               w:color="auto"/>
            //      </w:pBdr>
            //      <w:jc w:val="right"/>
            //      <w:rPr>
            //        <w:rFonts w:ascii="Verdana"
            //                  w:hAnsi="Verdana"/>
            //        <w:color w:val="000000"/>
            //        <w:sz w:val="18"/>
            //      </w:rPr>
            //    </w:pPr>
            //    <w:r>
            //      <w:rPr>
            //        <w:rFonts w:ascii="Verdana"
            //                  w:hAnsi="Verdana"/>
            //        <w:color w:val="000000"/>
            //        <w:sz w:val="18"/>
            //      </w:rPr>
            //      <w:t>Editorial Errepar</w:t>
            //    </w:r>
            //  </w:p>
            //</w:ftr>
            Footer f = new Footer();
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            ParagraphStyleId pStyle = new ParagraphStyleId() { Val = "Piedepagina" };
            TopBorder tb = new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4, Space = 1, Color = "auto" };
            ParagraphBorders pb = new ParagraphBorders { TopBorder = tb };
            pp.Append(pStyle, pb);
            Justification jc = new Justification() { Val = new EnumValue<JustificationValues>(JustificationValues.Right) };
            pp.Append(jc);
            ParagraphMarkRunProperties prpr = new ParagraphMarkRunProperties();
            RunFonts prf = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Color pc = new Color() { Val = "000000" };
            FontSize psz = new FontSize() { Val = "18" };
            prpr.Append(prf, pc, psz);
            pp.Append(prpr);
            Run r = new Run();
            RunProperties rp = new RunProperties();
            RunFonts rf = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana" };
            Color c = new Color() { Val = "000000" };
            FontSize sz = new FontSize() { Val = "18" };
            rp.Append(rf, c, sz);
            r.Append(rp);
            Text text = new Text() { Text = "Editorial Errepar" };
            r.Append(text);
            p.Append(pp, r);
            f.Append(p);
            return f;
        }

        private static string StripHTML(string source)
        {
            try
            {
                string result;

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = source.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating spaces because browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                                                                      @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*head([^>])*>", "<head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*head( )*>)", "</head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<head>).*(</head>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*script([^>])*>", "<script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*script( )*>)", "</script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result,
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty,
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<script>).*(</script>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*style([^>])*>", "<style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*style( )*>)", "</style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<style>).*(</style>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*td([^>])*>", "\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*br( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*li( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*div([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*tr([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*p([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything that's enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<[^>]*>", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @" ", " ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&bull;", " * ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lsaquo;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&rsaquo;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&trade;", "(tm)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&frasl;", "/",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lt;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&gt;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&copy;", "(c)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&reg;", "(r)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&(.{2,6});", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testing
                //System.Text.RegularExpressions.Regex.Replace(result,
                //       this.txtRegex.Text,string.Empty,
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4.
                // Prepare first to remove any whitespaces in between
                // the escaped characters and remove redundant tabs in between line breaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\t)", "\t\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\r)", "\t\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\t)", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multiple tabs following a line break with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for line breaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // That's it.
                return result;
            }
            catch
            {
                return source;
            }
        }
    }
}
