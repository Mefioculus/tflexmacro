using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
// Для работы с файловым справочником
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References.Files;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    
    }
    
    public void compareSpec()
    {
        /* Метод для проведения анализа содержимого таблиц SPEC и SPEC_OUT */
        ДиалогОжидания.Показать("Пожалуйста, подождите", true);
        ДиалогОжидания.СледующийШаг("Производится подбор данных для анализа из таблиц SPEC и SPEC_OUT");
        
        // Для начала нужно получить из SPEC_OUT список изделий, которе необходимо проверить
        Объекты headItems = НайтиОбъекты("SPEC_OUT", "[IZD] = ''");

        // Получаем список всех объектов в справочнике SPEC_OUT (TFlex)
        Объекты allItemsInTflex = НайтиОбъекты("SPEC_OUT", "[ID] != 0");

        // Рекурсивным методом получаем названия всех ДСЕ из таблицы SPEC (FoxPro)
        Объекты allItemsInFoxPro = new Объекты();

        foreach (Объект item in headItems)
        {
            foreach (Объект record in getListOfItemFromSpec(item))
            {
                allItemsInFoxPro.Add(record);
            }
        }

        ДиалогОжидания.СледующийШаг("Формируются результирующие данные");

        // Переходим к сравнению данных

        // Получаем список позиций, которые отсутствуют в FoxPro
        Объекты missingItemsInFoxPro = getMissingItems(allItemsInFoxPro, allItemsInTflex);

        // Получаем список позиций, которые отсутствуют в TFlex
        Объекты missingItemsInTflex = getMissingItems(allItemsInTflex, allItemsInFoxPro);

        Объекты differItems = getMissmatchingItem(allItemsInFoxPro, allItemsInTflex);





        // Находим позиции, по которым были обнаружены несоответствия
        // getListOfMissmatchingItem();
        // Выводим полученные данные в текстовый файл
        printResult(allItemsInFoxPro, allItemsInTflex,missingItemsInFoxPro, missingItemsInTflex, differItems);
        ДиалогОжидания.Скрыть();
        Сообщение("Информация", "Макрос закончил обработку данных успешно");
    }
        public Объекты getListOfItemFromSpec(Объект headItems)
        {
            // Инициализируем переменную, которая будет хранить данные
            Объекты itemsFromSpec = new Объекты();

            // Находим список объектов, которые входят в данное изделие

            Объекты tempItems = НайтиОбъекты("SPEC", String.Format("[IZD] = '{0}'", headItems.Параметр["SHIFR"]));

            if (tempItems != null)
            {
                foreach(Объект item in tempItems)
                {
                    itemsFromSpec.Add(item);
                    Объекты listOfChild = getListOfItemFromSpec(item);
                    if (listOfChild != null)
                    {
                        itemsFromSpec = concatenateList(itemsFromSpec, listOfChild);
                    }
                }
                return itemsFromSpec;
            }
            else
            {
                return null;
            }
        }

        public Объекты getMissingItems(Объекты firstListOfItems, Объекты secondListOfItems)
        {
            /* Функция проверяет, какие записи из первого справочника отсутствуют во втором справочике
            и возвращают результирующий список */

            // Инициализируем переменную, которая будет хранить отсутствующие позиции
            Объекты missingItems = new Объекты();

            // Инициализируем список со строками, который будет содержать все имена позиций, которые встречаются в
            // первом списке
            List<string> shifrItemsInfirstItems = new List<string>();
            foreach (Объект item in firstListOfItems)
            {
                shifrItemsInfirstItems.Add(String.Format("{0} - {1}",
                                                        item.Параметр["SHIFR"].ToString().Replace(".", "").ToUpper(),
                                                        item.Параметр["IZD"].ToString().Replace(".", "").ToUpper()));
            }
            
            // Проходим циклом по всем переменным второго списка для определения записи, которая отсутствует в первом справочнике
            foreach (Объект item in secondListOfItems)
            {
                if (shifrItemsInfirstItems.Contains(String.Format("{0} - {1}",
                                                    item.Параметр["SHIFR"].ToString().Replace(".", "").ToUpper(),
                                                    item.Параметр["IZD"].ToString().Replace(".", "").ToUpper())) != true)
                {
                    missingItems.Add(item);
                }
            }

            return missingItems;
        }

        public Объекты getMissmatchingItem(Объекты firstListOfItems, Объекты secondListOfItems)
        {
            // Получаем список позиций, у которых совпало все кроме применяемости
            Объекты MissmatchingItem = new Объекты();

            foreach (Объект itemFirst in firstListOfItems)
            {
                foreach (Объект itemSecond in secondListOfItems)
                {
                    if ((itemFirst.Параметр["SHIFR"].ToString().ToUpper().Replace(".", "") == itemSecond.Параметр["SHIFR"].ToString().ToUpper().Replace(".", ""))
                        &&
                        (itemFirst.Параметр["IZD"].ToString().ToUpper().Replace(".", "") == itemSecond.Параметр["IZD"].ToString().ToUpper().Replace(".", ""))
                        &&
                        (itemFirst.Параметр["PRIM"] != itemSecond.Параметр["PRIM"]))

                        {
                            MissmatchingItem.Add(itemFirst);
                            MissmatchingItem.Add(itemSecond);
                        }

                }
            }

            // ПРИМЕЧАНИЕ: В данном списке поочередно хранятся данные двух списков, это следует держать во внимании при работе с ним
            return MissmatchingItem;
        }

        public void printResult(Объекты itemsFromFoxPro, Объекты itemsFromTflex, Объекты missingItemsFoxPro, Объекты missingItemsTflex, Объекты differItems)
        {
            // Получение пути к файлу

            string pathToResultFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "result.txt");
            
            // Для начала создаем файл на рабочем столе пользователя
            StreamWriter result = File.CreateText(pathToResultFile);
            
            //Определим порядок колонок
            
            #region Настройка отображения таблицы
            string[] nameOfColumns = new string[] {"SHIFR", "NAIM", "IZD", "POS", "PRIM"};
            string templateForColumns = "{0,-40}|{1,-40}|{2,-40}|{3,-4}|{4,-4}|";
            string splitter = "-------------------------------------------------------------------------------------------------------------------------------------";
            string header = String.Format(  templateForColumns,
                                            nameOfColumns[0],
                                            nameOfColumns[1],
                                            nameOfColumns[2],
                                            nameOfColumns[3],
                                            nameOfColumns[4]);
            #endregion
            

            // Заполнение файла
            #region Печать исходных таблиц с данными

            if (Вопрос("Печатать в результирующий файл таблицы, по которым идет сравнение?") == true)
            {
                // Печать данных, которые содержутся в таблице FoxPro (Spec)
                result.WriteLine("\nСписок изделий, полученный из таблицы с данными из FoxPro (SPEC)\n");
                // Печать заголовка
                result.WriteLine(String.Format("{1}\n{0}\n{1}", header, splitter));
                foreach(Объект item in itemsFromFoxPro)
                {
                    result.WriteLine(String.Format( templateForColumns,
                                                    item.Параметр[nameOfColumns[0]],
                                                    item.Параметр[nameOfColumns[1]],
                                                    item.Параметр[nameOfColumns[2]],
                                                    item.Параметр[nameOfColumns[3]],
                                                    item.Параметр[nameOfColumns[4]]));
                }
                result.WriteLine(splitter);


                // Печать данных, которые сожержутся в таблицу TFlex (Spec_Out)
                result.WriteLine("\nСписок изделий, полученных из таблицы с данными из T-Flex (SPEC_OUT)\n");
                // Печать заголовка
                result.WriteLine(String.Format("{1}\n{0}\n{1}", header, splitter));
                foreach(Объект item in itemsFromTflex)
                {
                    result.WriteLine(String.Format( templateForColumns,
                                                    item.Параметр[nameOfColumns[0]],
                                                    item.Параметр[nameOfColumns[1]],
                                                    item.Параметр[nameOfColumns[2]],
                                                    item.Параметр[nameOfColumns[3]],
                                                    item.Параметр[nameOfColumns[4]]));
                }
                result.WriteLine(splitter);

            }
            #endregion


            #region Печать отсутствующих позиций в таблицах
            // Перечень позиций, которые отсутствуют в файле FoxPro
            if (missingItemsFoxPro != null)
            {
                result.WriteLine("Перечень позиций, которые отсутствуют в таблице FoxPro:\n");
                // Печать заголовка
                result.WriteLine(String.Format("{1}\n{0}\n{1}", header, splitter));

                foreach(Объект item in missingItemsFoxPro)
                {
                    result.WriteLine(String.Format(templateForColumns,
                                                item.Параметр[nameOfColumns[0]],
                                                item.Параметр[nameOfColumns[1]],
                                                item.Параметр[nameOfColumns[2]],
                                                item.Параметр[nameOfColumns[3]],
                                                item.Параметр[nameOfColumns[4]]));
                }
                result.WriteLine(splitter);

            }
            else
            {
                result.WriteLine("в FoxPro нет отсутствующих позиций:\n");
            }

            // Перечень позиций, которые отсутствуют в TFlex
            if (missingItemsTflex != null)
            {
                result.WriteLine("\nПеречень позиций, которые отсутствуют в таблице TFlex:\n");
                
                // Печать заголовка
                result.WriteLine(String.Format("{1}\n{0}\n{1}", header, splitter));

                foreach(Объект item in missingItemsTflex)
                {
                    result.WriteLine(String.Format(templateForColumns,
                                                item.Параметр[nameOfColumns[0]],
                                                item.Параметр[nameOfColumns[1]],
                                                item.Параметр[nameOfColumns[2]],
                                                item.Параметр[nameOfColumns[3]],
                                                item.Параметр[nameOfColumns[4]]));
                }
                result.WriteLine(splitter);

            }
            else
            {
                result.WriteLine("в Tflex нет отсутствующих позиций:\n");
            }
            #endregion


            #region Печать совпавших позиций, по которым имеются несоответствия по применяемости

            result.WriteLine("\nПозиции, которые в FoxPro и T-Flex имеют разные значения по применяемости\n");

            if (differItems != null)
            {
                // Печать заголовка
                result.WriteLine(header);

                // Печатаем смысловую часть
                int counter = 0;
                foreach (Объект item in differItems)
                {
                    if (counter % 2 == 0)
                    {
                        result.WriteLine(splitter);
                        result.WriteLine("В FoxPro:");
                    }
                    else
                    {
                        result.WriteLine("В T-Flex:");
                    }
                    counter++;
                    result.WriteLine(String.Format( templateForColumns,
                                                    item.Параметр[nameOfColumns[0]],
                                                    item.Параметр[nameOfColumns[1]],
                                                    item.Параметр[nameOfColumns[2]],
                                                    item.Параметр[nameOfColumns[3]],
                                                    item.Параметр[nameOfColumns[4]]));
                }
                // Печатаем подводящую черту
                result.WriteLine(splitter);

            }
            else
            {
                result.WriteLine("\nПозиции, имеющие различия только по применяемости отстутствуют");
            }
            #endregion

            #region Импорт файла в файловый справочник DOCs и удаление временного файла
            // Закрытие файла
            result.Close();
            // Импорт файла в базу данных
            // Инициализируем новый экземпляр файлового справочника
            FileReference fileRef = new FileReference(Context.Connection);

            // Получаем объект папки в файловом справочнике, в котором будет храниться исходный файл
            ДиалогВыбораОбъектов диалог = СоздатьДиалогВыбораОбъектов("Файлы");
            диалог.Заголовок = "Выбор папки для сохранения";
            диалог.МножественныйВыбор = false;
            
            Объект folder = null;
            
            if (диалог.Показать() == true)
            {
                folder = диалог.ВыбранныеОбъекты[0];
            }

            // Пробуем произвести импорт файла в выбранную папку
            try
            {
                FileObject file = fileRef.AddFile(pathToResultFile, (FolderObject)folder);
                Desktop.CheckIn(file, "Сохранение изменений в файле", false);
            }
            catch
            {
                Сообщение("Ошибка","Не удалось произвести сохранение в выбранную директорию. Результирующий файл можно найти на рабочем столе");
                return;
            }


            // 


            // Удаление временного файла из системы
            File.Delete(pathToResultFile);
            #endregion
        }

        public Объекты concatenateList (Объекты firstList, Объекты secondList)
        {
            foreach (Объект item in secondList)
            {
                firstList.Add(item);
            }
        return firstList;
        }
}
