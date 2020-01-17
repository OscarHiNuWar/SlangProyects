using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using Umbraco.Core;
using Umbraco.Core.Composing;
using Umbraco.Core.Events;
using Umbraco.Core.Models;
using Umbraco.Core.Services.Implement;


namespace WebApplication14.EventsSubscribers
{
    // register our component with Umbraco using a Composer
    [RuntimeLevel(MinLevel = RuntimeLevel.Run)]
    public class ContentPublishedComposer : ComponentComposer<ContentPublishedComponent>
    {
        // nothing needed to be done here!
    }

    public class ContentPublishedComponent : IComponent
    {
        // initialize: runs once when Umbraco starts
        public void Initialize()
        {
            // subscribe to content service published event
            ContentService.Published += ContentService_Published;
        }

        private void ContentService_Published(Umbraco.Core.Services.IContentService sender, Umbraco.Core.Events.ContentPublishedEventArgs e)
        {
            var Content = e.PublishedEntities.First();
            if (Content.ContentType.Alias == "categoria")
            {
                var excel = Content.Properties["importarExcel"].GetValue();
                if (excel!=null)
                {
                    var Excel = Path.Combine("http://localhost:65295" + excel.ToString());
                    var client = new WebClient();
                    var fullPath = Path.GetTempFileName();
                    client.DownloadFile(Excel, fullPath);
                    ObtenerEstudiantesMeritorios(fullPath, Content,e);
                }
            }
        }


        private void ObtenerEstudiantesMeritorios(string fileName, IContent Content, ContentPublishedEventArgs e)
        {
            long records = 0;
            int ChildCount = Current.Services.ContentService.CountChildren(Content.Id);
            if (ChildCount > 0)
            {
                var Childrens = Current.Services.ContentService.GetPagedChildren(Content.Id, 0, ChildCount, out records);
                foreach (var child in Childrens)
                {
                    Current.Services.ContentService.Delete(child);
                }
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;

                var sheets = workbookPart.Workbook.Sheets.Cast<Sheet>().ToList();
                var worksheetParts = workbookPart.WorksheetParts;
                worksheetParts = worksheetParts.OrderBy(w => sheets.IndexOf(sheets.Where(x => x.Id.Value == doc.WorkbookPart.GetIdOfPart(w)).First()));
                foreach (var worksheetPart in worksheetParts)
                {
                    Worksheet sheet = worksheetPart.Worksheet;
                    var rows = sheet.Descendants<Row>().Skip(1);

                    foreach (Row row in rows)
                    {
                        var celdas = row.Elements<Cell>().ToArray();
                        if (celdas[0].CellValue != null)
                        {
                            var Numero = celdas[0].CellValue.Text;
                            var Matricula = celdas[1].CellValue.Text;
                            var NombreCompleto = sst.ChildElements[int.Parse(celdas[2].CellValue.Text)].InnerText;
                            var Decanato = sst.ChildElements[int.Parse(celdas[3].CellValue.Text)].InnerText;
                            var Carrera = sst.ChildElements[int.Parse(celdas[4].CellValue.Text)].InnerText;
                            var CodigoCarrera = sst.ChildElements[int.Parse(celdas[5].CellValue.Text)].InnerText;
                            var Estudiante = Current.Services.ContentService.Create($"{Matricula} - {NombreCompleto}", Content.Id, "estudianteMeritorio");
                            Estudiante.Properties["no"].SetValue(Numero);
                            Estudiante.Properties["matricula"].SetValue(Matricula);
                            Estudiante.Properties["nombreCompleto"].SetValue(NombreCompleto);
                            Estudiante.Properties["decanato"].SetValue(Decanato);
                            Estudiante.Properties["carrera"].SetValue(Carrera);
                            Estudiante.Properties["codigoCarrera"].SetValue(CodigoCarrera);
                            Current.Services.ContentService.SaveAndPublish(Estudiante,"*", -1,false);
                            
                        }
                        else
                            break;

                    }



                }
            }
        }

        // terminate: runs once when Umbraco stops
        public void Terminate()
        {
            // do something when Umbraco terminates
        }
    }
}