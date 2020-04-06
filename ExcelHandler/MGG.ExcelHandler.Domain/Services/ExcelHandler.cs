using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MGG.ExcelHandler.Domain.Interfaces;
using MGG.ExcelHandler.Domain.Models;
using MGG.ExcelHandler.Domain.Resources;
using OfficeOpenXml;

namespace MGG.ExcelHandler.Domain.Services
{
    public class ExcelHandler : IExcelHandler
    {
        public IEnumerable<T> Read<T>(string path)
        {
            var file = GetFile(path);
            
            var excelPackage = new ExcelPackage(file);
            var excelModelList = excelPackage.Workbook.Worksheets.GetRowColumn();



            return null;
        }

        public FileInfo GetFile(string path)
        {
            if (File.Exists(path)) return new FileInfo(path);

            throw new FileNotFoundException(Labels.FileNotFound);
        }
    }

    internal static class ExcelHelper
    {
        internal static IEnumerable<ExcelModel> GetRowColumn(this ExcelWorksheets workSheets) =>
            workSheets.Select(excel => new ExcelModel
            {
                TotalRow = excel.Dimension.End.Row,
                TotalColumn = excel.Dimension.End.Column
            });
    }
}