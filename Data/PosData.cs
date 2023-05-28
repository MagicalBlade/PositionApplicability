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
        /// Номер позиции
        /// </summary>
        public string? Pos { get => _pos; set => _pos = value; }
        private string? _pos;
        /// <summary>
        /// Наименование марки
        /// </summary>
        public List<dynamic[]> Mark { get => _mark; set => _mark = value; }
        private List<dynamic[]> _mark = new();

        /// <summary>
        /// Масса позиции
        /// </summary>
        public double Weight { get => _weight; set => _weight = value; }
        private double _weight;
        /// <summary>
        /// Общая масса позиции
        /// </summary>
        public string? TotalMass { get => _totalMass; set => _totalMass = value; }
        private string? _totalMass;
        /// <summary>
        /// Есть ошибка в заполнении толщины?
        /// </summary>
        public bool IsErrorThickness { get => _isErrorThickness; set => _isErrorThickness = value; }
        private bool _isErrorThickness = false;
        /// <summary>
        /// Есть ошибки в заполнении ширины?
        /// </summary>
        public bool IsErrorWidth { get => _isErrorWidth; set => _isErrorWidth = value; }
        private bool _isErrorWidth = false;
        /// <summary>
        /// Есть ошибки в заполнении длины?
        /// </summary>
        public bool IsErrorLength { get => _isErrorLength; set => _isErrorLength = value; }
        private bool _isErrorLength = false;
        /// <summary>
        /// Есть ошибки в заполнении стали?
        /// </summary>
        public bool IsErrorSteel { get => _isErrorSteel; set => _isErrorSteel = value; }

        private bool _isErrorSteel = false;
        /// <summary>
        /// Есть ошибки в заполнении массы одной позиции
        /// </summary>
        public bool IsErrorWeight { get => _isErrorWeight; set => _isErrorWeight = value; }
        private bool _isErrorWeight = false;


        public PosData(ITable table, int row, string nameMark, int markcount, double weight, double qantityT, double qantityN, double totalWeight)
        {
            Pos = ((IText)table.Cell[row, 0].Text).Str;
            this.AddMark(table, row, nameMark, markcount, weight, qantityT, qantityN, totalWeight);
        }
        public bool AddMark(ITable table, int row, string nameMark, int markcount, double weight, double qantityT, double qantityN, double totalWeight)
        {
            Mark.Add(new dynamic[11]
            {
                nameMark,
                weight, //Масса одной позиции
                qantityT, //Количество таковских позиций
                qantityN, //Количество наоборотовских позиций
                totalWeight, //Общая масса
                markcount, // Количество марок
                ((IText)table.Cell[row, 3].Text).Str, // Толщина
                ((IText)table.Cell[row, 4].Text).Str, // Ширина
                ((IText)table.Cell[row, 5].Text).Str, // Длина
                ((IText)table.Cell[row, 8].Text).Str, // Сталь
                ((IText)table.Cell[row, 9].Text).Str, // Примечание

            });
            return true;
        }
    }
}
