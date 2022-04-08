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
        ТестовоеСозданиеПростойТаблицы();
    }

    public void ТестированиеСозданияЯчеек() {
        Message("Информация", "Тестирование прошло успешно");
    }

    public void ТестированиеСозданияСтрок() {
        Message("Информация", "Тестирование прошло успешно");
    }

    public void ТестовоеСозданиеПростойТаблицы() {
        // Создание новой таблицы с четырьмя колонками
        DCTable table = new DCTable(new List<int>() { 10, 20, 10, 10 });
        table.StartNewRow();
        table.AddCell("Первая ячейка");
        table.AddCell("Вторая ячейка");
        table.AddCell("Третья ячейка");
        table.AddCell("Четверная ячейка");
        table.StartNewRow();
        table.AddCell("Пятая ячейка");
        table.AddCell("Шестая ячейка");
        table.AddCell("Седьмая ячейка");
        table.AddCell("Восьмая ячейка");
        table.EndTable();

        Message("Содержимое таблицы", "\n" + table.ToString());


        string serializedTable = table.Serialize();
        
        
        Message("Информация о созданной таблице", table.GetInfo());

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

        public bool IsEnded { get; private set; }

        public int RowsCount => this.Rows.Count;
        public int ColsCount => this.WidthOfColumns.Count;

        public DCTable(List<int> widthOfColumns) {
            this.IsEnded = false;
            this.WidthOfColumns = widthOfColumns;
            this.Rows = new List<DCRow>();
        }

        public void EndTable() {
            // TODO: Реализовать код завершения редактирования таблицы
            VerifyTable();
            this.IsEnded = true;
        }

        public void StartNewRow(int height = 10) {

            // Если предыдущая строка не завершена, пытаемся ее завершить
            if ((this.CurrentRow != null) && (!this.CurrentRow.IsEnded))
                this.EndRow();

            DCRow row = new DCRow(this, this.Rows.Count, height);
            this.Rows.Add(row);
            this.CurrentRow = row;
        }

        public void AddCell(string text) {
            CheckPossibilityOfAddingCell();
            this.CurrentRow.AddCell(text);
        }

        public void AddCell(string text, TypeOfSpan span, int spanValue) {
            CheckPossibilityOfAddingCell();
            switch (span) {
                case TypeOfSpan.Horizontal:
                    this.CurrentRow.AddHorSpanCell(text, spanValue);
                    break;
                case TypeOfSpan.Vertical:
                    this.CurrentRow.AddVerSpanCell(text, spanValue);
                    break;
                default:
                    throw new Exception($"Данный метод добавления ячейки не поддерживает тип объекдинения {span.ToString()}. Воспользуйтесь другой перегрузкой");
            }
        }

        public void AddCell(string text, TypeOfSpan span, int horValue, int verValue) {
            CheckPossibilityOfAddingCell();
            switch (span) {
                case TypeOfSpan.Rectangular:
                    this.CurrentRow.AddRecSpanCell(text, horValue, verValue);
                    break;
                default:
                    throw new Exception($"Данный метод добавления ячейки не поддерживает тип объекдинения {span.ToString()}. Воспользуйтесь другой перегрузкой");
            }
        }

        public void AddEmptyCell(int quantity) {
            // TODO: Реализовать метод добавления пустых строк
        }

        private void CheckPossibilityOfAddingCell() {
            if (this.CurrentRow == null)
                throw new Exception("Перед добавлением ячейки необходимо создать новыю строку. Используйте метод StartNewRow для добавления ячейки");

            if (this.CurrentRow.IsEnded) {
                throw new Exception("Попытка добавить ячейку в завершенную строку. Используйте метод StartNewRow для добавления ячейки");
            }
        }

        private void EndRow() {
            // TODO: Реализовать завершение работы со строкой
            VerifyRow();
            this.CurrentRow = null;
        }

        private void VerifyRow() {
            // TODO: Реализовать код проверки корректности строки
        }

        private void VerifyTable() {
            //TODO: Реализовать таблицу на корректность перед завершением редактирования
        }

        public string Serialize() {
            if (!this.IsEnded)
                throw new Exception("Попытка сериализовать незаконченную таблицу. Перед вызовом метода Serialize вызовите метод EndTable");
            
            // TODO: Реализовать сериализацию объекта в строковое представление
            return string.Empty;
        }

        public static DCTable Parse(string inputString) {
            // TODO: Реализовать чтение таблицы из сериализованной строки
            return null;
        }

        public string GetInfo() {
            string result = string.Empty;
            result += $"Параметры таблицы:\nКоличество колонок: {this.ColsCount}\nРазмеры колонок: {string.Join(" ", this.WidthOfColumns.Select(width => width.ToString()))}\nКоличество строк: {this.RowsCount}\n";

            return result;
        }

        public override string ToString() {
            return string.Join("\n", this.Rows.Select(row => row.ToString()));
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

        public void AddVerSpanCell(string text, int span) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.Vertical, 0, span);
            this.Cells.Add(cell);
        }

        public void AddHorSpanCell(string text, int span) {
            this.CheckRowForEdit();
            DCCell cell = new DCCell(this.Table, this, this.Cells.Count, text, TypeOfSpan.Horizontal, span, 0);
            this.Cells.Add(cell);
        }

        public void AddRecSpanCell(string text, int horSpan, int verSpan) {
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

        public override string ToString() {
            return $"|{string.Join("|", this.Cells.Select(cell => cell.ToString()))}|";
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

        public override string ToString() {
            return $"{this.Text, 20}";
        }
    }

    private enum TypeOfSpan {
        None,
        Vertical,
        Horizontal,
        Rectangular
    }

 
}


