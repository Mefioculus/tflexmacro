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
    string protocolType = Convert.ToString(tableWithData.Rows[0]["ProtocolType"]);

    // В целях сокращения незанятого пространства

    #endregion Подготовительные работы

    #region Редактирование табличных данных

    switch (protocolType) {
        case "Протокол металлографической лаборатории":
            GenerateDefaultTable(tableString);
            break;
        case "Протокол физикомеханической лаборатории":
            GenerateDefaultTable(tableString);
            break;
        case "Протокол спектральной лаборатории":
            break;
        case "Протокол химической лаборатории":
            break;
        case "Протокол гальванической лаборатории":
            break;
        case "Протокол магнитной лаборатории":
            GenerateMagneteTable(tableString);
            break;
        case "Протокол электрической лаборатории":
            break;
        default:
            break;
    }



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
#region Метод для генерации таблицы по умолчанию

private void GenerateDefaultTable(string tableString) {

    // Приступаем к корректированию табличных данных
    RegularDataTable.BeginInit();

    if (tableString != string.Empty) {

        List<List<string>> rowsOfTable = ParseDataTable(tableString);

        const int widthOfOrderColumn = 30;

        // Размещаем полученные данные в итоговой таблице
        int counter = 1; // Счетчик параметров
        int rowIndex = 0; // Счетчик строк

        // Создаем список для хранения всех созданных ячеек (Необходимо для объединения ячеек)
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
}

#endregion Метод для генерации таблицы по умолчанию

#region Метод для генерации таблицы для протокола магнитной лаборатории

private void GenerateMagneteTable(string tableString) {
    RegularDataTable.BeginInit();

    // Основная особенность данной таблицы в том, что не смотря на то, что изначально она заявлялась как полностью
    // фиксированная по количеству колонок, в процессе работы выяснилось, что количество замеров магнитной индукции
    // у разных материалов может варьироваться от 2-х до 6-ти.
    
    // Следовательно возникает потребность в том, чтобы таблица менялась динамически

    // Завершаем работу метода, если строка не содержит информации с данными таблицы
    // (скорее всего это означает, что пользователь пытается сформировать отчет на протокол
    // без подключенного материала)
    if (tableString == string.Empty)
        return;

    // Получаем таблицу с данными сразу, для того, чтобы понять, какого размера нам нужно формировать таблицу
    List<List<string>> rowsOfTable = ParseDataTable(tableString);
    int numberOfDinamicColumns = rowsOfTable[0].Count - 3; // Вычитаем из общего количества колонок количество неизменных колонок


    // Создаем словарь с значением ширин колонок
    Dictionary<int, int> widthOfColumns = new Dictionary<int, int>();
    int counter = 1; // Счетчик для нумерации колонки
    int lengthOfInductionColumn = 25;
    int lengthOfInductionColumnSummary = (int)(lengthOfInductionColumn * numberOfDinamicColumns);
    
    // Заполняем первые две неизменные колонки
    widthOfColumns.Add(counter++, 50);
    widthOfColumns.Add(counter++, 25);
    // Заполняем динамически изменяемые колонки
    for (int i = 0; i < numberOfDinamicColumns; i++)
        widthOfColumns.Add(counter++, lengthOfInductionColumn);
    // Заполняем последнюю незименяемую колонку
    widthOfColumns.Add(counter++, 35);


    // Приступаем к формированию шапки таблицы
    TableCellInit.Text = "Марка материала, размер, мм";
    TableCellInit.WidthF = widthOfColumns[1];
    
    List<CellData> firstColumnNames = new List<CellData>() {
        new CellData("№ Контр. образца", widthOfColumns[2]),
        new CellData("Магнитная индукция А/м при напряжении магнитного поля, Тл", lengthOfInductionColumnSummary),
        new CellData("Коэрцитивная сила, А/м", widthOfColumns[widthOfColumns.Count])
    };

    // Добавляем колонки заголовка в таблицу
    foreach (CellData cellData in firstColumnNames) {
        XRTableCell cell = new XRTableCell();
        cell.Text = cellData.NameColumn;
        cell.WidthF = cellData.Width;
        RegularDataTable.Rows[0].Cells.Add(cell);
    }
    
    // Приступаем к формированию регуляной части таблицы
    foreach (List<string> row in rowsOfTable) {
        XRTableRow tableRow = new XRTableRow();
        RegularDataTable.Rows.Add(tableRow);

        int orderOfColumn = 1;
        foreach (string value in row) {
            XRTableCell cell = new XRTableCell();
            cell.Text = value;
            cell.Width = widthOfColumns[orderOfColumn++];
            tableRow.Cells.Add(cell);
        }
    }

    // Объединяем общие ячейки в заголовке
    // (Часть заголовка находилась в регулярной части, поэтому объединить эти колонки ранее не предствлялось возможным)
    RegularDataTable.Rows[0].Cells[0].RowSpan = 2;
    RegularDataTable.Rows[0].Cells[1].RowSpan = 2;
    RegularDataTable.Rows[0].Cells[3].RowSpan = 2;


    // Завершаем работу с таблицей
    RegularDataTable.EndInit();
}

#endregion Метод для генерации таблицы для протокола магнитной лаборатории

#region ParseDateTable
private List<List<string>> ParseDataTable (string stringFromInput, bool tableWithHeader = true) {
    // Данный метод преобразовывает входную струку с табличными данными в отдельные строки.
    // Так же данный метод будет удалять назадействованные колонки
    List<List<string>> result = new List<List<string>>();
    
    // Получаем список строк
    string[] rows = stringFromInput.Split(';');

    // Определяем, какие колонки требуется выводить на печать
    SortedSet<int> indexes = GetIndexesOfColumnsForPrinting(rows, tableWithHeader);

    // Возвращаем данные
    foreach (string row in rows) {
        result.Add(GetRowForPrint(row, indexes));
    }
    
    return result;
}
#endregion ParseDataTable

#region GetIndexesOfColumnsForPrinting
private SortedSet<int> GetIndexesOfColumnsForPrinting(string[] table, bool ignoreHeader) {

    SortedSet<int> indexes = new SortedSet<int>();
    int startIndex = 1;
    if (!ignoreHeader)
        startIndex = 0;

    for (int i = startIndex; i <= table.Length - 1; i++) {
        string[] values = table[i].Split('^');
        for (int j = 0; j <= values.Length - 1; j++) {
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

#region CellData class

private class CellData {
    public string NameColumn { get; set; }
    public int Width { get; set; }

    public CellData (string name, int width) {
        this.NameColumn = name;
        this.Width = width;
    }
}

#endregion CellData class

#endregion service classes
