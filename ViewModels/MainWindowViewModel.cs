using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kompas6API5;
using KompasAPI7;
using PositionApplicability.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
        /// <summary>
        /// ProgressBar извлечения позиций
        /// </summary>
        [ObservableProperty]
        private double _pBExtraction_Value = 0;
        /// <summary>
        /// ProgressBar заполнения применяемости
        /// </summary>
        [ObservableProperty]
        private double _pBFill_Value = 0;
        [ObservableProperty]
        private List<PosData> _posList = new();
        
        public List<string> ExtractionLog { get => _Extractionlog; set => _Extractionlog = value; }
        private List<string> _Extractionlog = new();

        public List<string> FillLog { get => _Filllog; set => _Filllog = value; }
        private List<string> _Filllog = new();


        #region Извлечение позиций
        [RelayCommand(IncludeCancelCommand = true)]
        private async Task ExtractionPositions(CancellationToken token)
        {

            if (!Directory.Exists(PathFolderAssembly))
            {
                Info = "Не верный путь к сборкам";
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
            ExtractionLog.Clear();
            Info = "Началось извлечение позиций";
            PBExtraction_Value = 1;
            string[] assemblyFiles = Directory.GetFiles(PathFolderAssembly, "*.cdw", SearchOption.TopDirectoryOnly);
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
                if (kompasDocuments2D == null)
                {
                    ExtractionLog.Add($"{pathfile} - не удалось открыть чертеж");
                    continue;
                }
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
                bool foundTable = false;
                foreach (IView view in views)
                {
                    ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                    IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                    foreach (ITable table in drawingTables)
                    {
                        IText text = (IText)table.Cell[0, 0].Text;
                        if (text.Str.IndexOf(StrSearchTableAssembly) != -1 && table.RowsCount > 2 && table.ColumnsCount == 10)
                        {
                            foundTable = true;
                            for (int row = 3; row < table.RowsCount; row++)
                            {
                                if (((IText)table.Cell[row, 0].Text).Str != "" && (((IText)table.Cell[row, 1].Text).Str != "" || ((IText)table.Cell[row, 2].Text).Str != ""))
                                {
                                    foundTable = true;
                                    int markIndex = PosList.FindIndex(x => x.Pos == ((IText)table.Cell[row, 0].Text).Str);
                                    if (markIndex != -1)
                                    {
                                        PosList[markIndex].AddMark(table, row, NameMark);
                                    }
                                    else
                                    {
                                        PosList.Add(new PosData(table, row, NameMark));
                                    }
                                }
                            }
                        }
                    }
                }
                if (!foundTable)
                {
                    ExtractionLog.Add($"{kompasDocuments2D.Name} - таблица не соответствует формату или не найдена");
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                if (token.IsCancellationRequested)
                {
                    kompas.Quit();
                    PBExtraction_Value = 0;
                    Info = "Извлечение отменено";
                    return;
                }
                PBExtraction_Value += 90 / assemblyFiles.Length;
            }
            kompas.Quit();
            PBExtraction_Value = 100;
            WriteLog(ExtractionLog, "ExtractionLog");
            Info = "Позиции извлечены";
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
        [RelayCommand(IncludeCancelCommand = true)]
        private async Task FillPos(CancellationToken token)
        {
            if (PosList.Count == 0)
            {
                Info = "Вначале извлеките позиции";
                return;
            }
            if (!Directory.Exists(PathFolderPos))
            {
                Info = "Не верный путь к деталям";
                return;
            }
            if (StrSearchTablePos == "")
            {
                Info = "Не указан текст для поиска таблицы применяемости";
                return;
            }

            await Task.Run(() => FillPosAsync(token));
        }
        private async Task FillPosAsync(CancellationToken token)
        {
            FillLog.Clear();
            Info = "Началось заполнение деталировки";
            PBFill_Value = 1;
            string[] posFiles = Directory.GetFiles(PathFolderPos, "*.cdw", SearchOption.TopDirectoryOnly);
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            if (kompasType == null) return;
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null) return;
            if (token.IsCancellationRequested)
            {
                kompas.Quit();
                PBFill_Value = 0;
                Info = "Заполнение деталировки отменено";
                return;
            }
            PBFill_Value = 10;
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            foreach (PosData pos in PosList)
            {
                Regex re = new Regex($@"поз.*\D{pos.Pos}\D.*cdw", RegexOptions.IgnoreCase);
                string[] path = Directory.GetFiles(PathFolderPos, $"*поз*{pos.Pos}*.cdw")
                    .Where(path => re.IsMatch(path))
                    .ToArray();

                if (path.Length == 0)
                {
                    FillLog.Add($"поз. {pos.Pos} - деталировка не найдена");
                    continue;
                }
                else if(path.Length > 1)
                {
                    FillLog.Add($"поз. {pos.Pos} - найдено более одного чертежа деталировки");
                    continue;
                }
                IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(path[0], false, false);
                if (kompasDocuments2D == null)
                {
                    FillLog.Add($"поз. {pos.Pos} - не удалось открыть чертеж");
                    continue;
                }
                IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                bool foundTable = false;
                foreach (IView view in views)
                {
                    ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                    IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                    foreach (IDrawingTable drawingTable in drawingTables)
                    {
                        ITable table = (ITable)drawingTable;
                        IText text = (IText)table.Cell[0, 0].Text;
                        if (text.Str.IndexOf(StrSearchTablePos) != -1 && table.RowsCount > 1 && table.ColumnsCount == 6)
                        {
                            foundTable = true;
                            double sumWeight = 0;
                            for (int indexrow = 0; table.RowsCount <= pos.Mark.Count + 3; indexrow++)
                            {
                                table.AddRow(indexrow + 3, true);
                            }
                            for (int markIndex = 0; markIndex < pos.Mark.Count; markIndex++)
                            {
                                ((IText)table.Cell[markIndex + 3, 0].Text).Str = pos.Mark[markIndex][0];
                                ((IText)table.Cell[markIndex + 3, 1].Text).Str = pos.Mark[markIndex][1];
                                ((IText)table.Cell[markIndex + 3, 2].Text).Str = pos.Mark[markIndex][2];
                                ((IText)table.Cell[markIndex + 3, 3].Text).Str = pos.Mark[markIndex][3];
                                ((IText)table.Cell[markIndex + 3, 4].Text).Str = pos.Mark[markIndex][4];
                                try
                                {
                                    sumWeight += double.Parse(pos.Mark[markIndex][4]);
                                }
                                catch (Exception)
                                {
                                    FillLog.Add($"{pos.Mark[markIndex][0]} - поз.{pos.Pos} - не корректная запись массы");
                                }
                            }
                            ((IText)table.Cell[table.RowsCount - 1, 4].Text).Str = sumWeight.ToString();
                            drawingTable.Update();
                        }
                    }
                }

                if (!foundTable)
                {
                    FillLog.Add($"{kompasDocuments2D.Name} - таблица не соответствует формату или не найдена");
                }
                kompasDocuments2D.Save();
                if (kompasDocuments2D.Changed)
                {
                    FillLog.Add($"{kompasDocuments2D.Name} - не удалось сохранить");
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                if (token.IsCancellationRequested)
                {
                    kompas.Quit();
                    PBFill_Value = 0;
                    Info = "Заполнение деталировки отменено";
                    return;
                }
                PBFill_Value += 90 / posFiles.Length;
            }
            kompas.Quit();
            WriteLog(FillLog, "FillLog");
            PBFill_Value = 100;
            Info = "Заполнение деталировки завершено";
            if (FillLog.Count > 0)
            {
                Info += ". Есть ошибки, посмотрите журнал.";
            }
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
                if (x.Pos == null || x.Pos == "")
                {
                    if (y.Pos == null || y.Pos == "")
                    {
                        return 0;
                    }
                    else
                    {
                        return -1;
                    }
                }
                else if (y.Pos == null || y.Pos == "")
                {
                    return 1;
                }
                double xd = double.Parse(x.Pos.Replace(".", ","));
                double yd = double.Parse(y.Pos.Replace(".", ","));
                if (xd > yd)
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
            int incrementRow = 0;
            if (worksheet != null)
            {
                for (int i = 0; i < PosList.Count; i++)
                {
                    for (int markIndex = 0; markIndex < PosList[i].Mark.Count; markIndex++)
                    {
                        
                        worksheet.Cell(i + incrementRow + 1, 1).Value = PosList[i].Pos;
                        worksheet.Cell(i + incrementRow + 1, 2).Value = PosList[i].Mark[markIndex][2];
                        worksheet.Cell(i + incrementRow + 1, 3).Value = PosList[i].Mark[markIndex][3];
                        worksheet.Cell(i + incrementRow + 1, 4).Value = PosList[i].Thickness;
                        worksheet.Cell(i + incrementRow + 1, 5).Value = PosList[i].Width;
                        worksheet.Cell(i + incrementRow + 1, 6).Value = PosList[i].Leigth;
                        worksheet.Cell(i + incrementRow + 1, 7).Value = PosList[i].Mark[markIndex][1];
                        worksheet.Cell(i + incrementRow + 1, 8).Value = PosList[i].Mark[markIndex][4];
                        worksheet.Cell(i + incrementRow + 1, 9).Value = PosList[i].Steel;
                        worksheet.Cell(i + incrementRow + 1, 10).Value = PosList[i].List;
                        worksheet.Cell(i + incrementRow + 1, 11).Value = PosList[i].Mark[markIndex][0];
                        incrementRow++;
                    }
                    incrementRow--;
                }
                worksheet.DataType = XLDataType.Text;
                //Ширина колонки по содержимому
                worksheet.Columns(1, PosList.Count).AdjustToContents();
                worksheet.Columns(1, PosList.Count).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            }
            
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

        /// <summary>
        /// Открыть файл журнала
        /// </summary>
        /// <param name="namelog"></param>
        [RelayCommand]
        private void OpenLog(string namelog)
        {
            Info = "";
            if (File.Exists($"{namelog}.txt"))
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo($"{namelog}.txt")
                {
                    UseShellExecute = true,
                };
                process.Start();
            }
            else
            {
                Info = "Файл журнала не найден.";
            }
        }

        /// <summary>
        /// Запись логов
        /// </summary>
        /// <param name="log"></param>
        private void WriteLog(List<string> log, string nameLog)
        {
            try
            {
                using (StreamWriter sw = new($"{nameLog}.txt", false))
                {
                    foreach (var item in log)
                    {
                        sw.WriteLine(item);
                    }
                    sw.Close();
                }
            }
            catch (Exception)
            {
                Info = $"Не удалось сохранить файл журнала {nameLog}";
            }
        }
    }
}
