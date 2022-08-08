using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Windows;
using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using System.Runtime.InteropServices;
using PositionApplicability.Data;
using Kompas6Constants3D;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PositionApplicability.ViewModels
{
    internal partial class MainWindowViewModel : ObservableObject
    {
        [ObservableProperty]
        private string? _pathFolder;

        [ObservableProperty]
        private List<PosData> _posList = new();

        [ICommand]
        private void Submit()
        {
            Type? kompasType = Type.GetTypeFromProgID("Kompas.Application.5", true);
            if (kompasType == null) return;
            KompasObject? kompas = Activator.CreateInstance(kompasType) as KompasObject; //Запуск компаса
            if (kompas == null) return;
            IApplication application = (IApplication)kompas.ksGetApplication7();
            IDocuments documents = application.Documents;
            IKompasDocument2D kompasDocuments2D = (IKompasDocument2D)documents.Open("d:\\Работа\\Хохлы\\КМД\\Блок К1т\\Блок К1т.cdw", false, false);
            IViewsAndLayersManager viewsAndLayersManager = kompasDocuments2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;

            foreach (IView item in views)
            {
                ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)item;
                IDrawingTables drawingTables = symbols2DContainer.DrawingTables;
                foreach (ITable table in drawingTables)
                {
                    IText text = (IText)table.Cell[0,0].Text;
                    if (text.Str.IndexOf("Спецификация") != -1 && table.RowsCount > 2 && table.ColumnsCount == 9)
                    {
                        for (int row = 3; row < table.RowsCount; row++)
                        {
                            PosList.Add(new PosData(table, row, kompasDocuments2D.Name));
                        }
                    }
                }
            }
            PosList.Sort();
            kompas.Quit();
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
    }
}
