using ExcelTestApp.Business;
using ExcelTestApp.Helper;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace ExcelTestApp
{
    public class ImportTool
    {
        private readonly ExcelDataManager<Disease> _excelDataManager;
        private readonly MapperService _mapperService;
        private readonly Mapper _mapper;

        public ImportTool()
        {
            _excelDataManager = new ExcelDataManager<Disease>();
            _mapperService = new MapperService();
            _mapper = new Mapper();
        }

        public int ImportDiseases()
        {
            string[] files = Directory.GetFiles("C:\\Users\\i.tesanovic\\Desktop\\TranslatedDiseases\\", "*.xlsx");
            var count = 0;

            foreach(var filePath in files)
            {
                var data = _excelDataManager.LoadDataFromFile(filePath);

                for (var i = 0; i < data.Count; i++)
                {
                    _mapperService.InsertDisease(data[i]);
                    count++;
                }
            }

            return count;
        }
    }
}