using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context) {
        }

    public override void Run() {
    }

    public void ТестированиеСозданияЯчеек() {
        Message("Информация", "Тестирование прошло успешно");
    }

    public void ТестированиеСозданияСтрок() {
        Message("Информация", "Тестирование прошло успешно");
    }

    public void ТестированиеСозданияТаблицы() {
        // Создание новой таблицы с четырьмя колонками
        DCTable table = new DCTable(new List<int>() { 10, 20, 10, 10 });
        Message("Информация о созданной таблице", $"Создана таблица: Колонок {table.ColsCount}, Строк {table.RowsCount}");
        Message("Информация", "Тестирование прошло успешно");
    }

    // TODO: Разработать класс для переноса таблиц из макроса в менеджер отчета
    // - Класс должен позволять объединять ячейки по горизонтали и по вертикали
    // - Класс должен позволять задавать размеры ячейки (высоту и ширину)
    // - Класс должен производить верификацию данных при формировании
    private class DCTable {
        public List<int> WidthOfColumns { get; private set; }
        public List<DCRow> Rows { get; private set; }

        private DCRow CurrentRow { get; set; }

        public int RowsCount => this.Rows.Count;
        public int ColsCount => this.WidthOfColumns.Count;

        public DCTable(List<int> widthOfColumns) {
            this.WidthOfColumns = widthOfColumns;
            this.Rows = new List<DCRow>();
        }

        // TODO: Реализовать код завершения редактирования таблицы
        public void EndEdit() {

        }

        // TODO: Реализовать код проверки корректности строки
        private void VerifyRow() {
        }

        public void StartRow(int height = 10) {
            if ((this.CurrentRow != null) && (!this.CurrentRow.IsEnded)) {
                throw new Exception("Перед началом работы с новой строкой сначала завершите работу с предыдущей");
            }

            DCRow row = new DCRow(this, this.Rows.Count, height);
            this.Rows.Add(row);
            this.CurrentRow = row;
        }

        // TODO: Реализовать завершение работы со строкой
        public void EndRow() {
        }

        // TODO: Реализовать сериализацию объекта в строку
        public string Serialize() {
            return string.Empty;
        }

        // TODO: Реализовать чтение таблицы из сериализованной строки
        public static DCTable Parse(string inputString) {
            return null;
        }

    }

    private class DCRow {
        public DCTable Table { get; private set; }
        public int IndexRow {get; private set; }
        public List<DCCell> Cells { get; private set; }
        public bool IsEnded { get; private set; } = false;
        public int Height { get; private set; }

        public DCRow(DCTable table, int index, int height) {
            this.Table = table;
            this.IndexRow = index;
            this.Height = height;
            this.Cells = new List<DCCell>();
        }

        public void AddCell(string text) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.None, 0, 0);
            this.Cells.Add(cell);
        }

        public void AddVerSpanCell(int span, string text) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.Vertical, 0, span);
            this.Cells.Add(cell);
        }

        public void AddHorSpanCell(int span, string text) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.Horizontal, span, 0);
            this.Cells.Add(cell);
        }

        public void AddRecSpanSell(int horSpan, int verSpan, string text) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.Rectangular, horSpan, verSpan);
            this.Cells.Add(cell);
        }

        private void CheckRowForEdit() {
            if (this.IsEnded)
                throw new Exception($"Попытка добавить новую ячейку в строку, редактирование которой завершено (Индекс строки {this.IndexRow})");
        }

        public void EndEdit() {
            VerifyRow();
            this.IsEnded = true;
        }

        private void VerifyRow() {
            // Производим проверку на соответствие количества заявленных ячеек с количеством введенных пользователем
            int cellCount = 0;
            foreach (DCCell cell in this.Cells) {
                cellCount += 1 + cell.HorizontalSpanValue;
            }

            if (cellCount != this.Table.WidthOfColumns.Count)
                throw new Exception($"Количество введенных ячеек не соответствует количеству указанных при инициализации колонок (Есть: {cellCount} => Должно быть: {this.Table.WidthOfColumns.Count})");
        }
    }

    private class DCCell {

        public DCTable Table { get; private set; }
        public DCRow Row { get; private set; }
        public int IndexCell { get; private set; }
        public int IndexRow => this.Row.IndexRow;

        public string Text { get; set; }

        // Параметры объединения ячеек
        public TypeOfSpan SpanType { get; private set; }
        public int VerticalSpanValue { get; private set; }
        public int HorizontalSpanValue { get; private set; }
        public bool HasSpan => this.SpanType == TypeOfSpan.None ? false : true;
        public bool HasSpanned { get; private set; }
        
        // Параметры размера ячейки
        public int Width =>
            (this.SpanType != TypeOfSpan.Horizontal) && (this.SpanType != TypeOfSpan.Rectangular) ?
                this.Table.WidthOfColumns[this.IndexRow] :
                this.Table.WidthOfColumns
                    .Skip(this.IndexCell)
                    .Take(this.HorizontalSpanValue)
                    .Sum();
        public int Height =>
            (this.SpanType != TypeOfSpan.Vertical) && (this.SpanType != TypeOfSpan.Rectangular) ?
                this.Row.Height :
                this.Table.Rows
                    .Skip(this.IndexRow)
                    .Take(this.VerticalSpanValue)
                    .Select(row => row.Height)
                    .Sum();



        public DCCell(DCTable table, DCRow row, int index, string text, TypeOfSpan typeOfSpan, int horizontalSpanValue, int verticalSpanValue) {
            this.Table = table;
            this.Row = row;
            this.IndexCell = index;
            this.Text = text;

            this.SpanType = typeOfSpan;
            this.HorizontalSpanValue = horizontalSpanValue;
            this.HorizontalSpanValue = verticalSpanValue;
        }
    }

    private enum TypeOfSpan {
        None,
        Vertical,
        Horizontal,
        Rectangular
    }

 
}


