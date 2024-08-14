using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;


namespace OutputPipeValues
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<Element> pipes = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_PipeCurves)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Текстовый файл (*.xlsx)|*.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {

                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист1");
                int indexRow = 0;

                foreach (Element pipe in pipes)
                {
                    string pipeName = pipe.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString();
                    string pipeOuterDiametr = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString();
                    string pipeInnerDiametr = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsValueString();
                    string pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString();

                    IRow row = sheet.CreateRow(indexRow++);
                    row.CreateCell(0).SetCellValue(pipeName);
                    row.CreateCell(1).SetCellValue(pipeOuterDiametr);
                    row.CreateCell(2).SetCellValue(pipeInnerDiametr);
                    row.CreateCell(3).SetCellValue(pipeLength);
                }

                using (FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream, false);
                }
                TaskDialog.Show("!", "Данные записаны!");
            }


            return Result.Succeeded;
        }
    }
}
