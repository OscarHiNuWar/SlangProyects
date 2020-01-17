using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Umbraco.Core;
using Umbraco.Core.Services;
using Umbraco.Core.Services.Implement;
using Umbraco.Core.Models;
using Umbraco.Core.Composing;
using Umbraco.Core.Logging;
using Umbraco.Core.Events;
using System.IO;
using System.Net;

namespace PruebaGeneral.Funciones
{

    
    public class ExcelToUmbraco : IComponent
    {
        private readonly ILogger _logger;
        public IContentService _prueba;

        [RuntimeLevel(MinLevel = RuntimeLevel.Run)]
        public class LogWhenPublishedComposer : ComponentComposer<ExcelToUmbraco> { }

        public ExcelToUmbraco(ILogger logger, IContentService prueba)
        {
            _logger = logger;
            _prueba = prueba; 
        }

        // initialize: runs once when Umbraco starts
        public void Initialize()
        {
            // subscribe to content service published event
            ContentService.Published += ContentService_Published;
        }

        public void ReadExcelFileSax(string fileName, IContent e)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;

                var sheets = workbookPart.Workbook.Sheets.Cast<Sheet>().ToList();
                var worksheetParts = workbookPart.WorksheetParts;
                worksheetParts = worksheetParts.OrderBy(w => sheets.IndexOf(sheets.Where(x => x.Id.Value == spreadsheetDocument.WorkbookPart.GetIdOfPart(w)).First()));

                OpenXmlReader reader = OpenXmlReader.Create(workbookPart);

                string text = string.Empty;
                Worksheet sheet = worksheetPart.Worksheet;
                var rows = sheet.Descendants<Row>().Skip(1);

                foreach (Row row in rows)
                {
                    var celdas = row.Elements<Cell>().ToArray();
                    if (celdas[0].CellValue != null)
                    {
                        var Nombre = sst.ChildElements[int.Parse(celdas[1].CellValue.Text)].InnerText;
                        var Descripcion = sst.ChildElements[int.Parse(celdas[2].CellValue.Text)].InnerText; 
                        var Sinonimos = sst.ChildElements[int.Parse(celdas[1].CellValue.Text)].InnerText;
                        //var Fecha = DateTime.Parse(sst.ChildElements[int.Parse(celdas[1].CellValue.Text)].InnerText);
                        var content = _prueba.Create(Nombre, e.Id, "slangs");
                        content.SetValue("descripcion", Descripcion);
                        content.SetValue("sinonimos", Sinonimos);
                        //content.SetValue("fechaPublicacion", Fecha);
                        _prueba.SaveAndPublish(content, "*", -1, false);
                    }
                    else { break; }
                }

                /*while (reader.Read())
                {
                    text = reader.GetText();
                    var content = _prueba.Create("My First Blog Post 2", e.Id, "slangs");

                    content.SetValue("descripcion", text);
                    _prueba.SaveAndPublish(content, "*", -1, false);
                }*/

            }
        }

        private void ContentService_Published(IContentService sender, ContentPublishedEventArgs e)
        {
            // the custom code to fire everytime content is published goes here!
            var ide = e.PublishedEntities.First();
            if (ide.ContentType.Alias == "contenedorSlang")
            {
                var xls = ide.GetValue("Documento");
                //var excel = Content.Properties["Documento"].GetValue();
                if(xls.ToString() != String.Empty || xls != null)
                {
                    var docExcel = Path.Combine("https://localhost:44347" + xls.ToString());
                    var cliente = new WebClient();
                    var fullPath = Path.GetTempFileName();
                    cliente.DownloadFile(docExcel, fullPath);
                    ReadExcelFileSax(fullPath, ide);
                }
                
                //var content = _prueba.Create("My First Blog Post 2", ide, "slangs");
                //content.SetValue("descripcion", "<p>Test</p>");
                //_prueba.SaveAndPublish(content, "*", -1, false);
            }

        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }
    }
}