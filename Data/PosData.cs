using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PositionApplicability.Data
{
    internal class PosData
    {
        /// <summary>
        /// Наименование марки
        /// </summary>
        public string? Mark { get => _mark; set => _mark = value; }
        private string? _mark;
        /// <summary>
        /// Номер позиции
        /// </summary>
        public string? Pos { get => _pos; set => _pos = value; }
        private string? _pos;
        /// <summary>
        /// Количество позиций
        /// </summary>
        public string? Quantity { get => _quantity; set => _quantity = value; }
        private string? _quantity;
        /// <summary>
        /// Сечение позиции (толщина х ширина)
        /// </summary>
        public string? Size { get => _size; set => _size = value; }
        private string? _size;
        /// <summary>
        /// Длина детали
        /// </summary>
        public string? Leigth { get => _leigth; set => _leigth = value; }
        private string? _leigth;
        /// <summary>
        /// Марка стали позиции
        /// </summary>
        public string? Steel { get => steel; set => steel = value; }
        private string? steel;

        /// <summary>
        /// Масса позиции
        /// </summary>
        public string? Weight { get => _weight; set => _weight = value; }
        private string? _weight;
        /// <summary>
        /// Общая масса позиции
        /// </summary>
        public string? TotalMass { get => _totalMass; set => _totalMass = value; }
        private string? _totalMass;
        /// <summary>
        /// Номер листа чертежа
        /// </summary>
        public string? List { get => _list; set => _list = value; }
        private string? _list;

        public PosData(ITable table, int row, string nameFile)
        {
            Mark = nameFile;
            Pos = ((IText)table.Cell[row, 0].Text).Str;
            Quantity = ((IText)table.Cell[row, 1].Text).Str;
            Size = ((IText)table.Cell[row, 3].Text).Str;
            Leigth = ((IText)table.Cell[row, 4].Text).Str;
            Steel = ((IText)table.Cell[row, 5].Text).Str;
            Weight = ((IText)table.Cell[row, 6].Text).Str;
            TotalMass = ((IText)table.Cell[row, 7].Text).Str;
            List = ((IText)table.Cell[row, 8].Text).Str;
        }
    }
}
