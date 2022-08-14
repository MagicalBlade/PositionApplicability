using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kompas6API5;
using KompasAPI7;
using PositionApplicability.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PositionApplicability.ViewModels
{
    internal partial class MainWindowViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _pathFolderAssembly = Properties.Settings.Default.PathFolderAssembly;
        [ObservableProperty]
        private string _pathFolderPos = Properties.Settings.Default.PathFolderPos;
        [ObservableProperty]
        private string _strSearchTableAssembly = Properties.Settings.Default.StrSearchTableAssembly;
        [ObservableProperty]
        private string _strSearchTablePos = Properties.Settings.Default.StrSearchTablePos;
        [ObservableProperty]
        private string? _info;
        [ObservableProperty]
        private double _pBExtraction_Value = 0;
        [ObservableProperty]
        private List<PosData> _posList = new();
        
        public List<string> Log { get => _log; set => _log = value; }

        private List<string> _log = new();



        #region Извлечение позиций
        [RelayCommand(IncludeCancelCommand = true)]
        private async Task ExtractionPositions(CancellationToken token)
        {

            if (PathFolderAssembly == "")
            {
                Info = "Не выбран путь";
                return;
            }
            if (StrSearchTableAssembly == "")
            {
                Info = "Не указан текст для поиска спецификации";
                return;
            }
            Info = "";
            await Task.Run(() => ExtractionPositionsAsync(token), token);
        }
        private async Task ExtractionPositionsAsync(CancellationToken token)
        {
            Info = "Началось извлечение позиций";
            PBExtraction_Value = 1;
            string[] assemblyFiles;
            assemblyFiles = Directory.GetFiles(PathFolderAssembly, "*.cdw", SearchOption.TopDirectoryOnly);
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            PBExtraction_Value = 10;
            if (kompasType == null) return;
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null) return;
            if (token.IsCancellationRequested)
            {
                kompas.Quit();
                PBExtraction_Value = 0;
                Info = "Извлечение отменено";
                return;
            }
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            foreach (string pathfile in assemblyFiles)
            {
                IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(pathfile, false, false);
                if (kompasDocuments2D == null) continue;
                #region Получение имени марки из штампа
                ILayoutSheets layoutSheets = kompasDocuments2D.LayoutSheets;
                string NameMark = "";
                foreach (ILayoutSheet layoutSheet in layoutSheets)
                {
                    IStamp stamp = layoutSheet.Stamp;
                    IText text = stamp.Text[2];
                    NameMark = text.Str.Split(" ")[^1];
                    break;
                }
                #endregion
                IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                foreach (IView view in views)
                {
                    ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                    IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                    foreach (ITable table in drawingTables)
                    {
                        IText text = (IText)table.Cell[0, 0].Text;
                        if (text.Str.IndexOf(StrSearchTableAssembly) != -1 && table.RowsCount > 2 && table.ColumnsCount == 9)
                        {
                            for (int row = 3; row < table.RowsCount; row++)
                            {
                                if (((IText)table.Cell[row, 1].Text).Str != "")
                                {
                                    PosList.Add(new PosData(table, row, NameMark));
                                }
                            }
                        }
                    }
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                PBExtraction_Value += 90 / assemblyFiles.Length;
            }
            kompas.Quit();
            PBExtraction_Value = 100;
            Info = "Позиции извлечены";
        }
        /// <summary>
        /// Отмена извлечения позиций
        /// </summary>
        [RelayCommand]
        private void CancelExtractionPositions()
        {
            ExtractionPositionsCommand.Cancel();
        } 
        #endregion

        [RelayCommand]
        private void OpenFolderDialogAssembly()
        {
            FolderBrowserDialog dialog = new();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PathFolderAssembly = dialog.SelectedPath;
            }
        }
        [RelayCommand]
        private void OpenFolderDialogPos()
        {
            FolderBrowserDialog dialog = new();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                PathFolderPos = dialog.SelectedPath;
            }
        }
        #region Заполнение деталировки
        [RelayCommand]
        private void FillPos()
        {
            if (PosList.Count == 0)
            {
                Info = "Вначале извлеките позиции";
                return;
            }
            if (PathFolderPos == "")
            {
                Info = "Не выбран путь";
                return;
            }
            if (StrSearchTablePos == "")
            {
                Info = "Не указан текст для поиска таблицы применяемости";
                return;
            }
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            if (kompasType == null) return;
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null) return;
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            string pathPos = $"{PathFolderPos}\\1252 Поз.1047.cdw";
            IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(pathPos, false, false);

            IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            foreach (IView view in views)
            {
                ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                foreach (IDrawingTable drawingTable in drawingTables)
                {
                    ITable table = (ITable)drawingTable;
                    IText text = (IText)table.Cell[0, 1].Text;
                    if (text.Str.IndexOf(StrSearchTablePos) != -1 && table.RowsCount > 1 && table.ColumnsCount == 2)
                    {
                        ((IText)table.Cell[1, 0].Text).Str = PosList[0].Quantity;
                        ((IText)table.Cell[1, 1].Text).Str = PosList[0].Mark;
                        drawingTable.Update();
                    }
                }
            }
            kompasDocuments2D.Save();
            kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
            kompas.Quit();
            Info = "Таблица применяемости заполнена";
        } 
        #endregion
        [RelayCommand]
        private void SaveExcel()
        {
            if (PosList.Count == 0)
            {
                Info = "Вначале извлеките позиции";
                return;
            }
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

            string Pos = "";
            XLWorkbook workbook = new();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Позиции");

            if (worksheet != null)
            {
                for (int i = 0; i < PosList.Count; i++)
                {
                    worksheet.Cell(i + 1, 1).Value = PosList[i].Mark;
                    if (Pos != PosList[i].Pos)
                    {
                        worksheet.Cell(i + 1, 2).Value = PosList[i].Pos;
                        Pos = PosList[i].Pos;
                    }
                    worksheet.Cell(i + 1, 3).Value = PosList[i].Quantity;
                    worksheet.Cell(i + 1, 4).Value = PosList[i].Size;
                    worksheet.Cell(i + 1, 5).Value = PosList[i].Leigth;
                    worksheet.Cell(i + 1, 6).Value = PosList[i].Steel;
                    worksheet.Cell(i + 1, 7).Value = PosList[i].Weight;
                    worksheet.Cell(i + 1, 8).Value = PosList[i].TotalMass;
                    worksheet.Cell(i + 1, 9).Value = PosList[i].List;
                }
            }
            //Ширина колонки по содержимому
            worksheet.Columns(1, PosList.Count).AdjustToContents();
            worksheet.Columns(1, PosList.Count).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            try
            {
                workbook.SaveAs($"{PathFolderAssembly}\\Тест.xlsx");
            }
            catch (Exception)
            {
                Info = "Не удалось сохранить файл";
                return;
            }
            Info = "Файл сохранен";
            
        }

        [RelayCommand]
        private void Closing()
        {
            Properties.Settings.Default.PathFolderAssembly = PathFolderAssembly;
            Properties.Settings.Default.PathFolderPos = PathFolderPos;
            Properties.Settings.Default.StrSearchTableAssembly = StrSearchTableAssembly;
            Properties.Settings.Default.StrSearchTablePos = StrSearchTablePos;
            Properties.Settings.Default.Save();
        }
    }
}
