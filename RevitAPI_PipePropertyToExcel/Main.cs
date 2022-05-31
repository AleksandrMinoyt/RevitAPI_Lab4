using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Autodesk.Revit.DB.Plumbing;

namespace RevitAPI_PipePropertyToExcel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var pipes = new FilteredElementCollector(doc)
             .OfClass(typeof(Pipe))
             .Cast<Pipe>()
             .ToList();


            SaveFileDialog sfd = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel files (*.xlsx) | *.xlsx",
                FileName = "PipeParameters.xlsx",
                DefaultExt = ".xlsx"
            };

            string filePath = string.Empty;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                filePath = sfd.FileName;
            }

            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;


            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Трубы");

                int rowIndex = 0;

                sheet.SetCellValue(rowIndex, 0,"Тип трубы" );
                sheet.SetCellValue(rowIndex, 1,"Наружный Д" );
                sheet.SetCellValue(rowIndex, 2,"Внутренний Д" );
                sheet.SetCellValue(rowIndex, 3,"Длина" );

                foreach (var item in pipes)
                {
                    rowIndex++;
                    sheet.SetCellValue(rowIndex, 0, item.Name);
                    sheet.SetCellValue(rowIndex, 1, UnitUtils.ConvertFromInternalUnits(item.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), UnitTypeId.Millimeters));
                    sheet.SetCellValue(rowIndex, 2, UnitUtils.ConvertFromInternalUnits(item.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(), UnitTypeId.Millimeters));
                    sheet.SetCellValue(rowIndex, 3, UnitUtils.ConvertFromInternalUnits(item.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Millimeters));
                  
                }
                workbook.Write(fs);
                workbook.Close();
            }

            System.Diagnostics.Process.Start(filePath);



            return Result.Succeeded;
        }
    }
}
