using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kompas6API5;
using KompasAPI7;
using PositionApplicability.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace PositionApplicability.ViewModels
{
    internal partial class MainWindowViewModel : ObservableObject
    {
        [ObservableProperty]
        private string? _pathFolder;

        [ObservableProperty]
        private string? _info;

        [ObservableProperty]
        private List<PosData> _posList = new();

        [ICommand]
        private void ExtractionPositions()
        {
            string[] assemblyFiles;
            if (PathFolder == null)
            {
                return;
            }

            assemblyFiles = Directory.GetFiles(PathFolder, "*.cdw", SearchOption.TopDirectoryOnly);
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            if (kompasType == null) return;
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null) return;
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            foreach (string pathfile in assemblyFiles)
            {
                IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(pathfile, false, false);
                IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                foreach (IView view in views)
                {
                    ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                    IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                    foreach (ITable table in drawingTables)
                    {
                        IText text = (IText)table.Cell[0,0].Text;
                        if (text.Str.IndexOf("Спецификация") != -1 && table.RowsCount > 2 && table.ColumnsCount == 9)
                        {
                            for (int row = 3; row < table.RowsCount; row++)
                            {
                                if (((IText)table.Cell[row, 1].Text).Str != "")
                                {
                                    PosList.Add(new PosData(table, row, kompasDocuments2D.Name));
                                }
                            }
                        }
                    }
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
            }
            kompas.Quit();
           
            Info = "Готово";
        }



        [ICommand]
        private void OpenFolderDialog()
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PathFolder = dialog.SelectedPath;
            }
        }

        [ICommand]
        private void SaveExcel()
        {
            PosList.Sort(ComparePosData);
            //Сортировка списака по номеру позиции
            static int ComparePosData(PosData x, PosData y)
            {
                double xd = double.Parse(x.Pos.Replace(".", ","));
                double yd = double.Parse(y.Pos.Replace(".", ","));
                if (x.Pos == null)
                {
                    if (y.Pos == null)
                    {
                        return 0;
                    }
                    else
                    {
                        return -1;
                    }
                }
                else if (y.Pos == null)
                {
                    return 1;
                }
                else if (xd > yd)
                {
                    return 1;
                }
                else if (xd == yd)
                {
                    return 0;
                }
                else
                {
                    return -1;
                }
            }
            XLWorkbook workbook = new();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Позиции");

            if (worksheet != null)
            {
                for (int i = 0; i < PosList.Count; i++)
                {
                    worksheet.Cell(i + 1, 1).Value = PosList[i].Mark;
                    worksheet.Cell(i + 1, 2).Value = PosList[i].Pos;
                    worksheet.Cell(i + 1, 3).Value = PosList[i].Quantity;
                    
                    /*
                    worksheet.Cell(i + 1, 4).Value = PosList[i].Size;
                    worksheet.Cell(i + 1, 5).Value = PosList[i].Leigth;
                    worksheet.Cell(i + 1, 6).Value = PosList[i].Steel;
                    worksheet.Cell(i + 1, 7).Value = PosList[i].Weight;
                    worksheet.Cell(i + 1, 8).Value = PosList[i].TotalMass;
                    worksheet.Cell(i + 1, 9).Value = PosList[i].List;
                    */

                    //worksheet.Cell(i + 1, j + 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                }
            }
            //Ширина колонки по содержимому
            //worksheet.Columns(1, export[0].Length).AdjustToContents();

            workbook.SaveAs($"{PathFolder}\\Тест.xlsx");
            Info = "Файл сохранен";
        }
    }
}
