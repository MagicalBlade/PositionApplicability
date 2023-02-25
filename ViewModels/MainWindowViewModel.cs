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
        #region Параметры окна
        [ObservableProperty]
        private double _heightWindow = Properties.Settings.Default.Height;
        [ObservableProperty]
        private double _widthWindow = Properties.Settings.Default.Width;
        #endregion
        /// <summary>
        /// Путь к сборкам
        /// </summary>
        [ObservableProperty]
        private string _pathFolderAssembly = Properties.Settings.Default.PathFolderAssembly;
        /// <summary>
        /// Путь к деталировкам
        /// </summary>
        [ObservableProperty]
        private string _pathFolderPos = Properties.Settings.Default.PathFolderPos;
        //Спецификация
        [ObservableProperty]
        private string _strSearchTableAssembly = Properties.Settings.Default.StrSearchTableAssembly;
        //Ведомость отправочных марок
        [ObservableProperty]
        private string _strSearchTableMark = Properties.Settings.Default.StrSearchTableMark;
        [ObservableProperty]
        private bool _isAllDirectoryExtraction = Properties.Settings.Default.IsAllDirectoryExtraction;
        [ObservableProperty]
        private bool _isAllDirectoryFill = Properties.Settings.Default.IsAllDirectoryFill;
        [ObservableProperty]
        private string? _info;
        /// <summary>
        /// ProgressBar извлечение позиций
        /// </summary>
        [ObservableProperty]
        private double _pBExtraction_Value = 0;
        /// <summary>
        /// ProgressBar заполнения применяемости
        /// </summary>
        [ObservableProperty]
        private double _pBFill_Value = 0;
        /// <summary>
        /// Массив позиций с данными по ним
        /// </summary>
        [ObservableProperty]
        private List<PosData> _posList = new();
        /// <summary>
        /// Данные по маркам для ММС
        /// </summary>
        [ObservableProperty]
        private List<string[]> _marksforMMC = new List<string[]>();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(OpenLogCommand))]
        private List<string> _log = new();
        


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
            OpenLogCommand.NotifyCanExecuteChanged();
        }
        private async Task ExtractionPositionsAsync(CancellationToken token)
        {
            PosList.Clear();
            MarksforMMC.Clear();
            Log.Clear();
            Info = "Началось извлечение позиций";
            PBExtraction_Value = 1;
            string[] assemblyFiles;
            if (IsAllDirectoryExtraction)
            {
                assemblyFiles = Directory.GetFiles(PathFolderAssembly, "*.cdw", SearchOption.AllDirectories);
            }
            else
            {
                assemblyFiles = Directory.GetFiles(PathFolderAssembly, "*.cdw", SearchOption.TopDirectoryOnly);
            }
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
                string mark = "";
                string markName = "";
                int markCountT = 0;
                int markCountN = 0;
                string markWeight = "";
                string markTotalWeight = "";
                string markSheet = "";
                IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(pathfile, false, false);
                if (kompasDocuments2D == null)
                {
                    Log.Add($"{pathfile} - не удалось открыть чертеж");
                    continue;
                }
                #region Получение имени марки из штампа
                ILayoutSheets layoutSheets = kompasDocuments2D.LayoutSheets;
                foreach (ILayoutSheet layoutSheet in layoutSheets)
                {
                    IStamp stamp = layoutSheet.Stamp;
                    IText text2 = stamp.Text[2]; //Текст из ячейки "Обозначения документа"
                    string[] text2Split = text2.Str.Split(" ");
                    mark = text2Split[^1];
                    //Записываем название марки
                    for (int i = 0; i < text2Split.Length - 1; i++)
                    {
                        markName += text2Split[i] + " ";
                    }
                    IText text16001 = stamp.Text[16001]; //Текст из ячейки "Лист"
                    markSheet = text16001.Str;
                    break;
                }
                #endregion
                IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                bool foundTableAssemble = false;
                bool foundTableMark = false;
                ITable? specTable = null;
                foreach (IView view in views)
                {
                    ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                    IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                    //Ведомость отправочных марок
                   
                    foreach (ITable table in drawingTables)
                    {
                        IText text = (IText)table.Cell[0, 0].Text;
                        if (text.Str.IndexOf(StrSearchTableMark) != -1 && table.RowsCount > 3 && table.ColumnsCount == 5)
                        {
                            if (((IText)table.Cell[3, 0].Text).Str != "" && (((IText)table.Cell[3, 1].Text).Str != "" || ((IText)table.Cell[3, 2].Text).Str != ""))
                            {
                                foundTableMark = true;
                                try
                                {
                                    if (((IText)table.Cell[3, 1].Text).Str != "")
                                    {
                                        markCountT = int.Parse(((IText)table.Cell[3, 1].Text).Str);
                                    }
                                    if (((IText)table.Cell[3, 2].Text).Str != "")
                                    {
                                        markCountN = int.Parse(((IText)table.Cell[3, 2].Text).Str);
                                    }
                                }
                                catch (Exception)
                                {
                                    Log.Add($"{mark} - не корректная запись количества марки");
                                }
                                //Запись масс марки
                                markWeight = ((IText)table.Cell[3, 3].Text).Str;
                                markTotalWeight = ((IText)table.Cell[3, 4].Text).Str;
                            }
                        }
                    }
                    foreach (ITable table in drawingTables)
                    {
                        IText text = (IText)table.Cell[0, 0].Text;
                        //Спецификация
                        if (text.Str.IndexOf(StrSearchTableAssembly) != -1 && table.RowsCount > 2 && table.ColumnsCount == 10)
                        {
                            foundTableAssemble = true;
                            specTable = table;
                        }
                    }
                }
                if (specTable != null)
                {
                    for (int row = 3; row < specTable.RowsCount; row++)
                    {
                        if (((IText)specTable.Cell[row, 0].Text).Str != "" && (((IText)specTable.Cell[row, 1].Text).Str != "" || ((IText)specTable.Cell[row, 2].Text).Str != ""))
                        {
                            double weight = 0;
                            int qantityT = 0;
                            int qantityN = 0;
                            double totalWeight = 0;
                            try
                            {
                                weight = double.Parse(((IText)specTable.Cell[row, 6].Text).Str); //Масса одной позиции
                            }
                            catch (Exception)
                            {
                                Log.Add($"{mark} - поз.{((IText)specTable.Cell[row, 0].Text).Str} - не корректная запись массы");
                            }
                            if (((IText)specTable.Cell[row, 1].Text).Str != "")
                            {
                                try
                                {
                                    qantityT = int.Parse(((IText)specTable.Cell[row, 1].Text).Str); //Количество таковских позиций
                                
                                }
                                catch (Exception)
                                {
                                    Log.Add($"{mark} - поз.{((IText)specTable.Cell[row, 0].Text).Str} - не корректная запись таковских позиций");
                                }
                            }
                            if (((IText)specTable.Cell[row, 2].Text).Str != "")
                            {
                                try
                                {
                                    qantityN = int.Parse(((IText)specTable.Cell[row, 2].Text).Str); //Количество наоборотовских позиций

                                }
                                catch (Exception)
                                {
                                    Log.Add($"{mark} - поз.{((IText)specTable.Cell[row, 0].Text).Str} - не корректная запись наоборотовских позиций");
                                }
                            }
                            try
                            {
                                totalWeight = double.Parse(((IText)specTable.Cell[row, 7].Text).Str); //Общая масса

                            }
                            catch (Exception)
                            {
                                Log.Add($"{mark} - поз.{((IText)specTable.Cell[row, 0].Text).Str} - не корректная запись общей массы");
                            }
                            foundTableAssemble = true;
                            int markIndex = PosList.FindIndex(x => x.Pos == ((IText)specTable.Cell[row, 0].Text).Str);
                            if (markIndex != -1)
                            {
                                //Проверка на ошибку в нумерации позиций, повторяющиеся номера позиций или ошибки в заполнении.
                                if (PosList[markIndex].Mark[0][6] != ((IText)specTable.Cell[row, 3].Text).Str)
                                {
                                    Log.Add($"поз.{((IText)specTable.Cell[row, 0].Text).Str} толщина различается! Проверьте соответствие по всем маркам.");
                                    PosList[markIndex].IsErrorThickness = true;
                                }
                                if (PosList[markIndex].Mark[0][7] != ((IText)specTable.Cell[row, 4].Text).Str)
                                {
                                    Log.Add($"поз.{((IText)specTable.Cell[row, 0].Text).Str} ширина различается! Проверьте соответствие по всем маркам.");
                                    PosList[markIndex].IsErrorWidth = true;
                                }
                                if (PosList[markIndex].Mark[0][8] != ((IText)specTable.Cell[row, 5].Text).Str)
                                {
                                    Log.Add($"поз.{((IText)specTable.Cell[row, 0].Text).Str} длина различается! Проверьте соответствие по всем маркам.");
                                    PosList[markIndex].IsErrorLength = true;
                                }
                                if (PosList[markIndex].Mark[0][9] != ((IText)specTable.Cell[row, 8].Text).Str)
                                {
                                    Log.Add($"поз.{((IText)specTable.Cell[row, 0].Text).Str} сталь различается! Проверьте соответствие по всем маркам.");
                                    PosList[markIndex].IsErrorSteel = true;
                                }
                                if (PosList[markIndex].Mark[0][1] != weight)
                                {
                                    Log.Add($"поз.{((IText)specTable.Cell[row, 0].Text).Str} вес различается! Проверьте соответствие по всем маркам.");
                                    PosList[markIndex].IsErrorWeight = true;
                                }

                                PosList[markIndex].AddMark(specTable, row, mark, markCountN + markCountT, weight, qantityT, qantityN, totalWeight);
                            }
                            else
                            {
                                PosList.Add(new PosData(specTable, row, mark, markCountN + markCountT, weight, qantityT, qantityN, totalWeight));
                            }
                        }
                    }
                }
                if (!foundTableAssemble)
                {
                    Log.Add($"{kompasDocuments2D.Name} - спецификация не соответствует формату или не найдена");
                }
                if (!foundTableMark)
                {
                    Log.Add($"{kompasDocuments2D.Name} - ведомость отправочных марок не соответствует формату или не найдена");
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                if (token.IsCancellationRequested)
                {
                    kompas.Quit();
                    PBExtraction_Value = 0;
                    Info = "Извлечение отменено";
                    return;
                }

                //Подготовка данных по маркам для ММС
                MarksforMMC.Add(new string[7] 
                {
                    mark,
                    markName,
                    markCountT.ToString(),
                    markCountN.ToString(),
                    markWeight,
                    markTotalWeight,
                    markSheet
                });

                PBExtraction_Value += 90 / assemblyFiles.Length;
            }
            kompas.Quit();
            PBExtraction_Value = 100;
            WriteLog();
            Info = "Позиции извлечены";
            if (Log.Count > 0)
            {
                Info += ". Есть ошибки, посмотрите журнал.";
            }
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

            await Task.Run(() => FillPosAsync(token));
            OpenLogCommand.NotifyCanExecuteChanged();

        }
        private async Task FillPosAsync(CancellationToken token)
        {
            Log.Clear();
            if (!File.Exists($"{Directory.GetCurrentDirectory()}\\Resources\\Ведомость отправочных марок.frw"))
            {
                Info = "Не найден файл 'Ведомость отправочных марок.frw' в папке Resources";
                return;
            }
            Info = "Началось заполнение деталировки";
            PBFill_Value = 1;
            SearchOption searchOptionFill;
            if (IsAllDirectoryFill)
            {
                searchOptionFill = SearchOption.AllDirectories;
            }
            else
            {
                searchOptionFill = SearchOption.TopDirectoryOnly;
            }
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
                string[] path = Directory.GetFiles(PathFolderPos, $"*поз*{pos.Pos}*.cdw", searchOptionFill)
                    .Where(path => re.IsMatch(path))
                    .ToArray();

                if (path.Length == 0)
                {
                    Log.Add($"поз. {pos.Pos} - деталировка не найдена");
                    continue;
                }
                else if(path.Length > 1)
                {
                    Log.Add($"поз. {pos.Pos} - найдено более одного чертежа деталировки");
                    continue;
                }
                IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open(path[0], false, false);
                if (kompasDocuments2D == null)
                {
                    Log.Add($"поз. {pos.Pos} - не удалось открыть чертеж");
                    continue;
                }
                #region Вставка таблицы "Ведомость отправочных марок" в чертеж деталировки
                double xSetPlacementTable = 0;
                double ySetPlacementTable = 0;
                ILayoutSheets layoutSheets = kompasDocuments2D.LayoutSheets;
                ILayoutSheet layoutSheet = layoutSheets.ItemByNumber[1];
                ISheetFormat sheetFormat = layoutSheet.Format;
                switch (sheetFormat.Format)
                {
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA0:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 836;
                            ySetPlacementTable = 1184;
                        }
                        else
                        {
                            xSetPlacementTable = 1184;
                            ySetPlacementTable = 836;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA1:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 589;
                            ySetPlacementTable = 836;
                        }
                        else
                        {
                            xSetPlacementTable = 836;
                            ySetPlacementTable = 589;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA2:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 415;
                            ySetPlacementTable = 589;
                        }
                        else
                        {
                            xSetPlacementTable = 589;
                            ySetPlacementTable = 415;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA3:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 292;
                            ySetPlacementTable = 415;
                        }
                        else
                        {
                            xSetPlacementTable = 415;
                            ySetPlacementTable = 292;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA4:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 205;
                            ySetPlacementTable = 292;
                        }
                        else
                        {
                            xSetPlacementTable = 292;
                            ySetPlacementTable = 205;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatA5:
                        if (sheetFormat.VerticalOrientation)
                        {
                            xSetPlacementTable = 143.5;
                            ySetPlacementTable = 205;
                        }
                        else
                        {
                            xSetPlacementTable = 205;
                            ySetPlacementTable = 143;
                        }
                        break;
                    case Kompas6Constants.ksDocumentFormatEnum.ksFormatUser:
                        xSetPlacementTable = sheetFormat.FormatWidth - 5;
                        ySetPlacementTable = sheetFormat.FormatHeight - 5;
                        break;
                    default:
                        break;
                }
                IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.View["Системный вид"];
                view.Current = true;
                view.Update();
                //Поиск существующих таблиц "Ведомость отправочных марок" и их удаление
                ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;
                IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                foreach (IDrawingTable item in drawingTables)
                {
                    ITable tableSearch = (ITable)item;
                    if (tableSearch.ColumnsCount == 2 && ((IText)tableSearch.Cell[0,0].Text).Str == "Кол-во" && ((IText)tableSearch.Cell[0, 1].Text).Str == "Марка")
                    {
                        item.Delete();
                    }
                }
                IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kompasDocuments2D;
                IDrawingGroups drawingGroups = kompasDocument2D1.DrawingGroups;
                IDrawingGroup drawingGroup = drawingGroups.Add(true, "");
                drawingGroup.ReadFragment(
                    $"{Directory.GetCurrentDirectory()}\\Resources\\Ведомость отправочных марок.frw",
                    true, 0, 0, 1, 0, false);
                ksDocument2D ksDocument2D = kompas.TransferInterface(kompasDocuments2D,1 ,0);
                //ksDocument2D.ksMoveObj(drawingGroup.Reference, xSetPlacementTable, ySetPlacementTable);
                IDrawingTable drawingTable = drawingGroup.Objects[0]; //Таблица
                ITable table = (ITable)drawingTable;
                if (table.ColumnsCount != 2 || table.RowsCount < 2)
                {
                    kompas.Quit();
                    PBFill_Value = 0;
                    Info = "Не корректная таблица 'Ведомость отправочных марок.frw' в папке Resources";
                    return;
                }
                double[] sumWeight = new double[pos.Mark.Count];
                //Создание строк таблицы
                for (int indexrow = 0; table.RowsCount < pos.Mark.Count + 1; indexrow++)
                {
                    table.AddRow(indexrow + 1, true);
                }
                //Заполнение строк таблицы
                for (int markIndex = 0; markIndex < pos.Mark.Count; markIndex++)
                {
                    if (pos.Mark[markIndex][2] != 0 || pos.Mark[markIndex][3] != 0)
                    {
                        ((IText)table.Cell[markIndex + 1, 0].Text).Str = $"{(pos.Mark[markIndex][2] + pos.Mark[markIndex][3]) * pos.Mark[markIndex][5]}"; //Количество позиций всего
                    }
                    ((IText)table.Cell[markIndex + 1, 1].Text).Str = pos.Mark[markIndex][0];
                }
                drawingTable.Update();
                ksRectParam ksRectangleParam = kompas.GetParamStruct(15);
                ksMathPointParam botPoint = ksRectangleParam.GetpBot();
                ksDocument2D.ksGetObjGabaritRect(drawingGroup.Reference, ksRectangleParam);
                
                ksDocument2D.ksMoveObj(drawingGroup.Reference, xSetPlacementTable, Math.Abs(botPoint.y) + 70);
                drawingGroup.Store();
                #endregion

                kompasDocuments2D.Save();
                if (kompasDocuments2D.Changed)
                {
                    Log.Add($"{kompasDocuments2D.Name} - не удалось сохранить");
                }
                kompasDocuments2D.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                if (token.IsCancellationRequested)
                {
                    kompas.Quit();
                    PBFill_Value = 0;
                    Info = "Заполнение деталировки отменено";
                    return;
                }
                PBFill_Value += 90 / PosList.Count;
            }
            kompas.Quit();
            WriteLog();
            PBFill_Value = 100;
            Info = "Заполнение деталировки завершено";
            if (Log.Count > 0)
            {
                Info += ". Есть ошибки, посмотрите журнал.";
            }
        }
        #endregion

        #region Заполнение неуказанной шероховатости в деталировке
        [RelayCommand(IncludeCancelCommand = true)]
        private async Task SpecRough(CancellationToken token)
        {
            if (!Directory.Exists(PathFolderPos))
            {
                Info = "Не верный путь к деталям";
                return;
            }

            await Task.Run(() => SpecRoughAsync(token));
            OpenLogCommand.NotifyCanExecuteChanged();

        }
        private async Task SpecRoughAsync(CancellationToken token)
        {
            Log.Clear();
            string pathRough = $@"{Directory.GetCurrentDirectory()}\Resources\Шероховатость.txt";
            string valuesRoughStr = "";
            if (!File.Exists(pathRough))
            {
                return;
            }
            if (!Directory.Exists(PathFolderPos))
            {
                return;
            }
            Info = "Начало заполнения шероховатости";
            PBFill_Value = 1;
            SearchOption searchOptionFill;
            using (StreamReader reader = new StreamReader(pathRough))
            {
                valuesRoughStr = reader.ReadToEnd();
            }
            if (valuesRoughStr == "")
            {
                PBFill_Value = 0;
                Info = "Файл со значениями шероховатости пуст";
                return;
            }
            Dictionary<string, string> valuesRough = new Dictionary<string, string>();
            foreach (string item in valuesRoughStr.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
            {
                string[] line = item.Split(' ', StringSplitOptions.None);
                if (line.Length > 1)
                {
                    valuesRough.Add(line[0], line[1]);
                }
            }
            if (IsAllDirectoryFill)
            {
                searchOptionFill = SearchOption.AllDirectories;
            }
            else
            {
                searchOptionFill = SearchOption.TopDirectoryOnly;
            }
            string[] filesDetailing = Directory.GetFiles(PathFolderPos, "*.cdw",searchOptionFill);
            if (filesDetailing.Length == 0)
            {
                PBFill_Value = 0;
                Info = "Не найдена деталировка";
                return;
            }
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            if (kompasType == null)
            {
                Info = "Не удалось запустить компас";
                return;
            }
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null)
            {
                Info = "Не удалось запустить компас";
                return;
            }
            if (token.IsCancellationRequested)
            {
                kompas.Quit();
                PBFill_Value = 0;
                Info = "Заполнение шероховатости отменено";
                return;
            }
            PBFill_Value = 10;
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            foreach (string path in filesDetailing)
            {
                IDrawingDocument kompasDocument = (IDrawingDocument)documents.Open(path, false, false);
                if (kompasDocument == null)
                {
                    PBFill_Value += 90 / filesDetailing.Length;
                    Log.Add($"{path} - не удалось открыть");
                    continue;
                }
                ILayoutSheets layoutSheets = kompasDocument.LayoutSheets;
                ILayoutSheet layoutSheet = layoutSheets.ItemByNumber[1];
                IStamp stamp = layoutSheet.Stamp;
                IText text3 = stamp.Text[3];
                string text3Str = text3.Str;
                string thickness = "";
                if (text3Str != "")
                {
                    string[] profile = text3Str.Split("$dsm; ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    if (profile.Length > 4)
                    {
                        thickness = profile[1];
                    }
                }
                ISpecRough specRough = kompasDocument.SpecRough;
                if (specRough != null)
                {
                    if (valuesRough.ContainsKey(thickness))
                    {
                        specRough.Text = $"Rz {valuesRough[thickness]}";
                        specRough.AddSign = false;
                        specRough.SignType = Kompas6Constants.ksRoughSignEnum.ksNoProcessingType;
                        specRough.Distance = 2;
                        specRough.Update();
                    }
                    else
                    {
                        Log.Add($"{path} - не найдена толщина. Проверьте правильность записи в штампе или в файле Resources\\Шероховатость.txt");
                        kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                    }
                }
                kompasDocument.Save();
                if (kompasDocument.Changed)
                {
                    Log.Add($"{path} - не удалось сохранить");
                }
                kompasDocument.Close(Kompas6Constants.DocumentCloseOptions.kdDoNotSaveChanges);
                PBFill_Value += 90 / filesDetailing.Length;
                if (token.IsCancellationRequested)
                {
                    kompas.Quit();
                    PBFill_Value = 0;
                    Info = "Заполнение шероховатости отменено";
                    return;
                }
            }
            kompas.Quit();
            PBFill_Value = 100;
            WriteLog();
            Info = "Заполнение шероховатости завершено";
            if (Log.Count > 0)
            {
                Info += ". Есть ошибки, посмотрите журнал.";
            }
        }
        #endregion
        [RelayCommand]
        private void OpenTxT(string file)
        {
            Info = "";
            if (File.Exists($"{file}"))
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo($"{file}")
                {
                    UseShellExecute = true,
                };
                process.Start();
            }
            else
            {
                Info = $"Файл {file} не найден";
            }
        }
        /// <summary>
        /// Сохранить файл отчёта
        /// </summary>
        [RelayCommand]
        private void SaveExcel()
        {
            if (PosList.Count == 0)
            {
                Info = "Вначале извлеките позиции";
                return;
            }
            //PosList.Sort(ComparePosData);
            //Сортировка списка по номеру позиции
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
            #region Лист "Позиции"
            IXLWorksheet worksheetPos = workbook.Worksheets.Add("Позиции");
            int incrementRow = 3; //Начальная строка
            #region Формирование шапки листа
            worksheetPos.Cell(1, 1).SetValue("Поз.");
            worksheetPos.Cell(1, 2).SetValue("Кол-во");
            worksheetPos.Cell(1, 4).SetValue("Сечение, мм");
            worksheetPos.Cell(1, 7).SetValue("Масса, кг");
            worksheetPos.Cell(1, 9).SetValue("Материал");
            worksheetPos.Cell(1, 10).SetValue("Примечание");
            worksheetPos.Cell(1, 11).SetValue("Марка");
            worksheetPos.Cell(1, 12).SetValue("Кол-во");
            worksheetPos.Cell(2, 2).SetValue("т");
            worksheetPos.Cell(2, 3).SetValue("н");
            worksheetPos.Cell(2, 4).SetValue("толщина");
            worksheetPos.Cell(2, 5).SetValue("ширина");
            worksheetPos.Cell(2, 6).SetValue("длина");
            worksheetPos.Cell(2, 7).SetValue("шт.");
            worksheetPos.Cell(2, 8).SetValue("общ.");
            worksheetPos.Cell(2, 12).SetValue("Марок");
            #endregion

            if (worksheetPos != null)
            {
                for (int i = 0; i < PosList.Count; i++)
                {
                    for (int markIndex = 0; markIndex < PosList[i].Mark.Count; markIndex++)
                    {
                        if (PosList[i].IsErrorThickness)
                        {
                            worksheetPos.Cell(i + incrementRow, 4).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        if (PosList[i].IsErrorWidth)
                        {
                            worksheetPos.Cell(i + incrementRow, 5).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        if (PosList[i].IsErrorLength)
                        {
                            worksheetPos.Cell(i + incrementRow, 6).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        if (PosList[i].IsErrorSteel)
                        {
                            worksheetPos.Cell(i + incrementRow, 9).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        if (PosList[i].IsErrorWeight)
                        {
                            worksheetPos.Cell(i + incrementRow, 7).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        worksheetPos.Cell(i + incrementRow, 1).SetValue(PosList[i].Pos);
                        if (PosList[i].Mark[markIndex][2] != 0)
                        {
                            worksheetPos.Cell(i + incrementRow, 2).SetValue(PosList[i].Mark[markIndex][2]);
                        }
                        if (PosList[i].Mark[markIndex][3] != 0)
                        {
                            worksheetPos.Cell(i + incrementRow, 3).SetValue(PosList[i].Mark[markIndex][3]);
                        }
                        worksheetPos.Cell(i + incrementRow, 4).SetValue(PosList[i].Mark[markIndex][6]); //Толщина
                        worksheetPos.Cell(i + incrementRow, 5).SetValue(PosList[i].Mark[markIndex][7]); //Ширина
                        worksheetPos.Cell(i + incrementRow, 6).SetValue(PosList[i].Mark[markIndex][8]); //Длина
                        worksheetPos.Cell(i + incrementRow, 7).SetValue(PosList[i].Mark[markIndex][1]); //Масса одной позиции
                        worksheetPos.Cell(i + incrementRow, 8).SetValue(PosList[i].Mark[markIndex][4]); //Общая масса
                        worksheetPos.Cell(i + incrementRow, 9).SetValue(PosList[i].Mark[markIndex][9]); //Сталь
                        worksheetPos.Cell(i + incrementRow, 10).SetValue(PosList[i].List); //Примечание
                        worksheetPos.Cell(i + incrementRow, 11).SetValue(PosList[i].Mark[markIndex][0]); //Название марки
                        worksheetPos.Cell(i + incrementRow, 12).SetValue(PosList[i].Mark[markIndex][5]); //Количество марок
                        incrementRow++;
                    }
                    incrementRow--;
                }
                worksheetPos.DataType = XLDataType.Text;
                //Ширина колонки по содержимому
                worksheetPos.Columns(1, PosList.Count).AdjustToContents(5.0, 100.0);
                #region Объединение ячеек
                worksheetPos.Range("B1:C1").Row(1).Merge();
                worksheetPos.Range("D1:F1").Row(1).Merge();
                worksheetPos.Range("G1:H1").Row(1).Merge();
                worksheetPos.Range("A1:A2").Column(1).Merge();
                worksheetPos.Range("I1:I2").Column(1).Merge();
                worksheetPos.Range("J1:J2").Column(1).Merge();
                worksheetPos.Range("K1:K2").Column(1).Merge();
                #endregion
                worksheetPos.Columns(1, PosList.Count).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheetPos.Columns(1, PosList.Count).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            }
            #endregion

            #region Лист "ММС"
            IXLWorksheet worksheetMMC = workbook.Worksheets.Add("ММС");
            int incrementRowMMC = 3;
            #region Формирование шапки листа
            worksheetMMC.Cell(1, 1).SetValue("Отпр.");
            worksheetMMC.Cell(1, 2).SetValue("Наименование");
            worksheetMMC.Cell(1, 3).SetValue("Кол-во");
            worksheetMMC.Cell(1, 5).SetValue("Масса, кг");
            worksheetMMC.Cell(1, 7).SetValue("№ черт.");
            worksheetMMC.Cell(2, 1).SetValue("марка");
            worksheetMMC.Cell(2, 3).SetValue("т");
            worksheetMMC.Cell(2, 4).SetValue("н");
            worksheetMMC.Cell(2, 5).SetValue("шт.");
            worksheetMMC.Cell(2, 6).SetValue("общ.");
            #endregion

            for (int line = 0; line < MarksforMMC.Count; line++)
            {
                worksheetMMC.Cell(line + incrementRowMMC, 1).SetValue(MarksforMMC[line][0]); //Марка
                worksheetMMC.Cell(line + incrementRowMMC, 2).SetValue(MarksforMMC[line][1].Trim()); //Название марки
                if (MarksforMMC[line][2] != "0")
                {
                    worksheetMMC.Cell(line + incrementRowMMC, 3).SetValue(MarksforMMC[line][2]); //Таковское количество
                }
                if (MarksforMMC[line][3] != "0")
                {
                    worksheetMMC.Cell(line + incrementRowMMC, 4).SetValue(MarksforMMC[line][3]); //Наоборотовское количество
                }
                worksheetMMC.Cell(line + incrementRowMMC, 5).SetValue(MarksforMMC[line][4]); //Единичная масса
                worksheetMMC.Cell(line + incrementRowMMC, 6).SetValue(MarksforMMC[line][5]); //Общая масса
                worksheetMMC.Cell(line + incrementRowMMC, 7).SetValue(MarksforMMC[line][6]); //Номер листа
                var cellWithFormulaA1 = worksheetMMC.Cell(line + incrementRowMMC, 8);
                cellWithFormulaA1.FormulaA1 = $@"==IF(((C{line + incrementRowMMC}+D{line + incrementRowMMC})*E{line + incrementRowMMC})=F{line + incrementRowMMC}, True, False)";
                if (cellWithFormulaA1.Value.ToString() == "False")
                {
                    worksheetMMC.Cell(line + incrementRowMMC, 6).Style.Fill.BackgroundColor = XLColor.Red;
                }
                worksheetMMC.Cell(line + incrementRowMMC, 8).Clear();
            }

            //worksheetMMC.DataType = XLDataType.Text;
            //Ширина колонки по содержимому
            worksheetMMC.Columns(1, 7).AdjustToContents(5.0, 100.0);
            #region Объединение ячеек
            worksheetMMC.Range("C1:D1").Row(1).Merge();
            worksheetMMC.Range("E1:F1").Row(1).Merge();
            worksheetMMC.Range("B1:B2").Column(1).Merge();
            worksheetMMC.Range("G1:G2").Column(1).Merge();
            #endregion
            worksheetMMC.Columns(1, 7).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheetMMC.Columns(1, 7).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            #endregion

            try
            {
                workbook.SaveAs($"{PathFolderAssembly}\\Отчёт.xlsx");
            }
            catch (Exception)
            {
                Info = "Не удалось сохранить файл";
                return;
            }
            Info = "Файл сохранен";
            
        }
        /// <summary>
        /// Открыть файл отчёта
        /// </summary>
        [RelayCommand]
        private void OpenExcel()
        {
            Info = "";
            if (File.Exists($@"{PathFolderAssembly}\Отчёт.xlsx"))
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo($@"{PathFolderAssembly}\Отчёт.xlsx")
                {
                    UseShellExecute = true,
                };
                process.Start();
            }
            else
            {
                Info = $"Файл {$@"{PathFolderAssembly}\Отчёт.xlsx"} не найден. Сохраните отчёт.";
            }
        }

        [RelayCommand]
        private void Closing()
        {
            Properties.Settings.Default.Height = HeightWindow;
            Properties.Settings.Default.Width = WidthWindow;
            Properties.Settings.Default.PathFolderAssembly = PathFolderAssembly;
            Properties.Settings.Default.PathFolderPos = PathFolderPos;
            Properties.Settings.Default.StrSearchTableAssembly = StrSearchTableAssembly;
            Properties.Settings.Default.IsAllDirectoryExtraction = IsAllDirectoryExtraction;
            Properties.Settings.Default.IsAllDirectoryFill = IsAllDirectoryFill;
            Properties.Settings.Default.StrSearchTableMark = StrSearchTableMark;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Открыть файл журнала
        /// </summary>
        /// <param name="log"></param>
        [RelayCommand(CanExecute = nameof(CanOpenLog))]
        private void OpenLog()
        {
            Info = "";
            if (File.Exists($"Log.txt"))
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo($"Log.txt")
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
        private bool CanOpenLog()
        {
            if (Log == null || Log.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Запись логов
        /// </summary>
        /// <param name="Log"></param>
        private void WriteLog()
        {
            try
            {
                using (StreamWriter sw = new($"Log.txt", false))
                {
                    foreach (var item in Log)
                    {
                        sw.WriteLine(item);
                    }
                    sw.Close();
                }
            }
            catch (Exception)
            {
                Info = $"Не удалось сохранить файл журнала";
            }
        }
    }
}
