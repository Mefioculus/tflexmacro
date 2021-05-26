using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;



private void xtraReport1_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e) {
    
    #region Основные переменные
    
    // Основные ячейки, которые могут пригодиться по ходу написания скрипта
    XRTableCell cellWithPositionEngineer = tableCell21;

    #endregion Основные переменные

    #region Подготовительные работы
    // Получаем данные, которые содержатся в списке полей
    DataSet ds = (DataSet)DetailRegularPart.DataSource;
    if (ds == null) {
        MessageBox.Show("ds is null");
        return;
    }

    // Находим нужный параметр
    var index = ds.Tables.IndexOf("Табличные данные (Табличные данные)");
    if (index == null) {
        MessageBox.Show("index is null");
        return;
    }
    var tableWithData = ds.Tables[index];
    if (tableWithData == null) {
        MessageBox.Show("tableWithData is null");
        return;
    }


    // Данный код нужен в случае, если отчет формируется без данных
    // (к примеру, для того, чтобы делать предпросмотр отчета)
    // в этом случае данных в таблице не будет, а, следовательно,
    // следующий за проверкой код будет вызывать ошибку
    if (tableWithData.Rows.Count == 0)
        return;
    string tableString = Convert.ToString(tableWithData.Rows[0]["DataTable"]);

    // В целях сокращения незанятого пространства

    #endregion Подготовительные работы

    #region Редактирование табличных данных

    // Приступаем к корректированию табличных данных
    RegularDataTable.BeginInit();

    if (tableString != string.Empty) {

        List<List<string>> rowsOfTable = ParseDataTable(tableString);

        /*
        // Проверка полученных табличных даннах
        foreach (List<string> rowOfData in rowsOfTable) {
            MessageBox.Show(string.Join('^', rowOfData));
        }
        */

        const int widthOfOrderColumn = 30;

        // Размещаем полученные данные в итоговой таблице
        int counter = 1; // Счетчик параметров
        int rowIndex = 0; // Счетчик строк

        // Создаем список для хранения всех созданных ячеек
        TableOfCells tableOfCells = new TableOfCells();

        // (для того, чтобы к каждой из них можно было вернуться
        // и изменить параметр Span)

        // Создаем таблицу
        foreach (List <string> row in rowsOfTable) {
            // Обрабатываем строки
            // Начинаем с первой строки, заполняя шапку таблицы
            if (rowIndex == 0) {
                // Заполняем самую первую ячейку (которая содержит название колонки
                // с порядковыми номерами
                TableCellInit.Text = "n/n";
                TableCellInit.WidthF = widthOfOrderColumn;
                foreach (string value in row) {
                    // Добавляем ячейки в строку в соответствием с данными
                    XRTableCell cell = new XRTableCell();
                    cell.Text = value;
                    RegularDataTable.Rows[rowIndex].Cells.Add(cell);
                }
            }
            else {
                // Обрабатываем последующие строки
                // Добавляем новую строку
                tableOfCells.AddNewRow();

                // Создаем новую строку в таблице
                XRTableRow tableRow = new XRTableRow();
                XRTableCell orderCell = new XRTableCell();

                // Добавляем созданную ячейку в таблицу для хранения параметра Span
                tableOfCells.AddCell(orderCell);

                // Присваиваем ячейке значение
                if (row[0] != "---") {
                    orderCell.Text = counter.ToString();
                    counter++;
                }
                else
                    orderCell.Text = "---";

                // Установка ширины ячейки и добавление ячейки в строку таблицы
                orderCell.WidthF = widthOfOrderColumn;
                tableRow.Cells.Add(orderCell);

                foreach (string value in row) {
                    // Создаем новую ячейку
                    XRTableCell cell = new XRTableCell();

                    // Добавляем ячейку в таблицу для хранения параметра Span
                    tableOfCells.AddCell(cell);

                    // Заполнение значения ячейки
                    // Производим проверку на нулевое значение.
                    // Если значение нулевое, ставим прочерк
                    if (value != "0")
                        cell.Text = value;
                    else
                        cell.Text = "-";

                    // Дабавление ячейки в таблицу отчета
                    tableRow.Cells.Add(cell);
                }
                // Добавляем новую строку к таблице
                RegularDataTable.Rows.Add(tableRow);
            }

            rowIndex++;
        }

        // Производим объединение ячеек
        for (int column = 0; column < tableOfCells.ColumnCount; column++) {
            int counterOfSpan = 1;
            for (int row = (tableOfCells.RowCount - 1); row >= 0; row--) {
                if (tableOfCells[row, column].Text == "---")
                    counterOfSpan++;
                else {
                    tableOfCells[row, column].RowSpan = counterOfSpan;
                    counterOfSpan = 1;
                }
            }
        }
    }
    else {
        // Удаляем заготовку таблицы с листа отчета
       RegularDataTable.Rows[0].Cells.Remove(TableCellInit);
       // Изменяем размер реглярной части раздела, для того, чтобы не было большого пустого места в отчете
       Detail2.HeightF = 0;
    }

    RegularDataTable.EndInit();

#endregion Редактирование табличных данных

#region Проверка подписей

// Получение содержимого подписи Ведущего инженера и утверждающего
    // Получаем данные, которые содержатся в списке полей

    // Находим нужный параметр
    index = ds.Tables.IndexOf("Подписи (Подписи)");
    if (index == null) {
        MessageBox.Show("index is null");
        return;
    }
    var tableWithSigns = ds.Tables[index];
    if (tableWithSigns == null) {
        MessageBox.Show("tableWithSigns is null");
        return;
    }

    if (tableWithSigns.Rows.Count == 0)
        return;

    string engineer = Convert.ToString(tableWithSigns.Rows[0]["Ведущий инженер ФИО"]);
    string approver = Convert.ToString(tableWithSigns.Rows[0]["Утвердил ФИО"]);


    if (engineer == approver) {
        // Заменяем значение метки, которая содержит название должности
        cellWithPositionEngineer.Text = string.Format("'{0}", cellWithPositionEngineer.Text);
    }
}

#endregion Проверка подписей

#region Service methods

#region ParseDateTable
private List<List<string>> ParseDataTable (string stringFromInput) {
    // Данный метод преобразовывает входную струку с табличными данными в отдельные строки.
    // Так же данный метод будет удалять назадействованные колонки
    List<List<string>> result = new List<List<string>>();
    
    // Получаем список строк
    string[] rows = stringFromInput.Split(';');

    // Определяем, какие колонки требуется выводить на печать
    SortedSet<int> indexes = GetIndexesOfColumnsForPrinting(rows);

    // Возвращаем данные
    foreach (string row in rows) {
        result.Add(GetRowForPrint(row, indexes));
    }
    
    return result;
}
#endregion ParseDataTable

#region GetIndexesOfColumnsForPrinting
private SortedSet<int> GetIndexesOfColumnsForPrinting(string[] table) {

    SortedSet<int> indexes = new SortedSet<int>();

    for (int i = 1; i <= table.Length - 1; i++) {
        string[] values = table[i].Split('^');
        for (int j =0; j <= values.Length - 1; j++) {
            if (NonDefault(values[j])) {
                indexes.Add(j);
            }
        }
    }

    return indexes;
}
#endregion GetIndexesOfColumnsForPrinting

#region GetRowForPrint
private List<string> GetRowForPrint(string rawRow, SortedSet<int> indexes) {
    List<string> valuesToPrint = new List<string>();
    // Разбиваем полученную строку на входящие элементы
    foreach (int index in indexes) {
        valuesToPrint.Add(rawRow.Split('^')[index]);
    }

    return valuesToPrint;

}
#endregion

#region NonDefault
private bool NonDefault(string value) {
    return (value != string.Empty) && (value != "0");
}
#endregion NonDefault

#endregion Service methods

#region service classes

#region TableOfCells class

private class TableOfCells {
    private List<List<XRTableCell>> TableOfCellsSpan = new List<List<XRTableCell>>();
    private List<XRTableCell> CurrentRow = new List<XRTableCell>();
    
    public int RowCount { get; private set; }
    public int ColumnCount { get; private set; }


    public void AddNewRow() {
        CurrentRow = new List<XRTableCell>();
        TableOfCellsSpan.Add(CurrentRow);
        RowCount++;
    }

    public void AddCell(XRTableCell cell) {
        CurrentRow.Add(cell);

        // Обновление количества колонок в таблице
        ColumnCount = (ColumnCount < CurrentRow.Count) ? CurrentRow.Count : ColumnCount;
    }

    // Итераторы для класса контейнера
    public XRTableCell this[int row, int column] {
        get {
            return TableOfCellsSpan[row][column];
        }
        set {
            TableOfCellsSpan[row][column] = value;
        }
    }

    public List<XRTableCell> this[int row] {
        get {
            return TableOfCellsSpan[row];
        }
        set {
            TableOfCellsSpan[row] = value;
        }
    }

}

#endregion TableOfCells class

#endregion service classes
