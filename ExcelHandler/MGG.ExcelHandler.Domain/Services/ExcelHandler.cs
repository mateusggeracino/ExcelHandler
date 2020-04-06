using System.Collections;
using System.Collections.Generic;
using System.IO;
using MGG.ExcelHandler.Domain.Interfaces;
using MGG.ExcelHandler.Domain.Resources;
using OfficeOpenXml;

namespace MGG.ExcelHandler.Domain.Services
{
    public class ExcelHandler : IExcelHandler
    {
        public IEnumerable<T> Read<T>(string path)
        {
            var file = GetFile(path);

            //Verify if is open
            //Get count rows and columns

            var excelPackage = new ExcelPackage(file);

            return null;
        }

        public FileInfo GetFile(string path)
        {
            if (File.Exists(path)) return new FileInfo(path);
            
            throw new FileNotFoundException(Labels.FileNotFound);
        }
    }
}