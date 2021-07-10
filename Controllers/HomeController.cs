using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Procs.Data;
using Procs.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using cm = System.ComponentModel;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.IO;
using OfficeOpenXml;

namespace Procs.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ApplicationDbContext _dbContext;

        public HomeController(ILogger<HomeController> logger, ApplicationDbContext dbContext)
        {
            this._logger = logger;
            this._dbContext = dbContext;
        }

        public IActionResult Index()
        {           
            return View();
        }


#pragma warning disable CA2211
        public static List<Producto> oProductosExcel;

        public FileResult ExportarProductosExcel(string[] nombrePropiedades)
        {
            byte[] buffer = ExportarExcel(nombrePropiedades, oProductosExcel);
            return File(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        public IActionResult Procedure()
        {
            var oProductos = this._dbContext.ProductosDb
                .FromSqlRaw("prueba.sp_Productos").ToList();


            oProductosExcel = oProductos;
            return View(oProductos);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public byte[] ExportarExcel<T>(string[] nombrePropiedades, List<T> oListaDataGeneric)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                try
                {
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        ep.Workbook.Worksheets.Add("DATA");

                        ExcelWorksheet ws = ep.Workbook.Worksheets[0];

                        Dictionary<string, string> diccionario = cm.TypeDescriptor.GetProperties(typeof(T))
                            .Cast<cm.PropertyDescriptor>().ToDictionary(p => p.Name, p => p.DisplayName);

                        Color colorNombreColumns = System.Drawing.ColorTranslator.FromHtml("#000000");
                        //Color colorFilas = System.Drawing.ColorTranslator.FromHtml("#CCCCFF");

                        if (nombrePropiedades != null && nombrePropiedades.Length > 0)//Validar checks nulos
                        {
                            if (oListaDataGeneric.Count < 1)
                            {
                                for (int i = 0; i < nombrePropiedades.Length; i++)
                                {
                                    ws.Cells[1, i + 1].Value = diccionario[nombrePropiedades[i]];//Para llenar las nombres de las columnas
                                    ws.Cells[ws.Dimension.Address].Style.Font.Bold = true;
                                    ws.Cells[ws.Dimension.Address].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    ws.Cells[ws.Dimension.Address].Style.Fill.BackgroundColor.SetColor(colorNombreColumns);
                                    ws.Cells[ws.Dimension.Address].Style.Font.Color.SetColor(Color.White);
                                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                                    ws.Cells[ws.Dimension.Address].AutoFilter = true;
                                }
                            }
                            else
                            {
                                //For para llenar nombres de columnas
                                for (int i = 0; i < nombrePropiedades.Length; i++)
                                {
                                    ws.View.FreezePanes(2, 1);
                                    ws.Cells[1, i + 1].Style.Font.Bold = true;
                                    ws.Cells[1, i + 1].Value = diccionario[nombrePropiedades[i]];//Para llenar las nombres de las columnas  
                                    ws.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                    ws.View.ShowGridLines = false;
                                }

                                //Para empezar a pintar la data
                                int fila = 2;
                                int col = 1;

                                //item se convierte en una fila, porque es un objeto, un objeto de la lista (ELEMENTO, se pinta en una fila)
                                foreach (object item in oListaDataGeneric)
                                {
                                    col = 1;

                                    foreach (string propiedad in nombrePropiedades)
                                    {
                                        ws.Cells[fila, col].Value =
                                          item.GetType().GetProperty(propiedad).GetValue(item).ToString();
                                        ws.Cells[fila, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                                        col++;
                                    }

                                    fila++;
                                }

                                int filaFin = fila - 1;
                                int colFin = col - 1;

                                ExcelRange rg = ws.Cells[1, 1, filaFin, colFin];

                                string tableName = "Tabla";

                                ExcelTable table = ws.Tables.Add(rg, tableName);

                                table.TableStyle = TableStyles.Dark8;

                                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            }
                        }

                        ep.SaveAs(ms);
                        byte[] buffer = ms.ToArray();
                        return buffer;
                    }
                }
                catch (Exception)
                {                 
                    throw;
                }
            }
        }
    }
}
