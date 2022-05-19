/* Дополнительные ссылки
PresentationCore.dll
WindowsBase.dll
System.Xaml.dll
TFlex.DOCs.Common.dll
DevExpress.Data.v19.1.dll
DevExpress.Docs.v19.1.dll
DevExpress.Office.v19.1.Core.dll
DevExpress.Pdf.v19.1.Core.dll
DevExpress.Pdf.v19.1.Drawing.dll
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.Office.Utils;
using DevExpress.Pdf;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;


namespace PDM_PrintWithMagick
{
    public class Macro : MacroProvider
    {
        #region Штамп

        private string _fontName = "Times-New-Roman"; //Times-New-Roman-Bold
        private int _stampBorderWidth = 2;

        private int _stampColor;
        private double _stampOpacity;
        private double _stampGlobalX;
        private double _stampGlobalY;
        private int _stampGlobalAngle;

        #endregion

        #region Guids

        private static class Guids
        {
            /// <summary> Guid типа "Печать файлов" справочника "Пользовательские диалоги" </summary>
            public static readonly Guid UserDialogClass = new Guid("4491b3e4-eb87-47ae-be32-d5df1af1a30f");

            /// <summary> Guid справочника "Форматы" </summary>
            public static readonly Guid PageFormatsReferenceId = new Guid("17816113-feb7-44a1-9024-7052eb3d9b52");

            /// <summary> Guid параметра "Наименование" справочника "Форматы" </summary>
            public static readonly Guid NameFormat = new Guid("a8ccef51-1d3c-4f5f-9c6b-b4fd55ff4dcf");

            /// <summary> Guid параметра "Высота" справочника "Форматы" </summary>
            public static readonly Guid HeightFormat = new Guid("0541757d-5a38-401b-a234-e073a35f509d");

            /// <summary> Guid параметра "Ширина" справочника "Форматы" </summary>
            public static readonly Guid WidthFormat = new Guid("2e1fdaf1-9d7e-4496-9da0-0c73fe85cc40");


            public static class Parameters
            {
                public static readonly Guid СокращенноеНазваниеПодразделения =
                    new Guid("d834e736-141a-42fa-b2d8-c14f58505686");

                public static readonly Guid КодПодразделения = new Guid("6d58df61-2a43-45b2-b8f6-34f3507107fd");
                public static readonly Guid ИсходныйФайл = new Guid("66ac7676-e3f5-4e6c-9b62-44213996ccc1");
                public static readonly Guid ФайлДляПечати = new Guid("03f339ed-39c3-463d-93f1-4825f3635396");
            }


            public static class Links
            {
                public static readonly Guid Подразделение = new Guid("378a7fd0-4490-490f-991a-22237e06defc");
                public static readonly Guid Форматы = new Guid("5e535c0a-309d-4975-b554-fe35659774f0");
                public static readonly Guid ДокументУчетнаяКарточка = new Guid("d708e1b4-2a1a-499c-aaaf-be5828e6377e");
            }
        }

        #endregion

        /// <summary> Путь к файлу 'magick.exe' в справочнике 'Файлы' </summary>
        public const string MagickExePath = @"Служебные файлы\Библиотека dll\magick.exe";

        /// <summary> Поддерживаемые расширения файлов </summary>
        private static readonly string[] SupportedExtensions = { ".tif", ".tiff", ".pdf" };

        DateTime _currentDateTime = DateTime.Now;

        private static string _папкаДляВременныхФайлов = Path.Combine(Path.GetTempPath(), "Temp DOCs");
        private static string _логОшибок = Path.Combine(_папкаДляВременныхФайлов, "TFlex.DOCs.PDM.Print.log");

        private string _localMagickFilePath;

        private string _сотрудник;
        private Объект _подразделениеПользователя;
        private Объекты _принтеры;
        private readonly Dictionary<string, string> _списокФорматПринтер = new Dictionary<string, string>();
        private int _текущийНомерПечатаемогоФайла;
        private List<Объект> _файлыНаПечать = new List<Объект>();

        private List<int> _списокИдентификаторовДокументов = new List<int>(0);
        private List<int> _списокИдентификаторовФайлов = new List<int>(0);


        public Macro(MacroContext context)
            : base(context)
        {
        }

        private void LoadBinaryFiles()
        {
            _localMagickFilePath = GetFileLocalPathFromFileReference(MagickExePath);
        }

        private string GetFileLocalPathFromFileReference(string relativeFilePath)
        {
            var fileReference = new FileReference(Context.Connection);

            var fileObject = fileReference.FindByRelativePath(relativeFilePath) as FileObject;

            if (fileObject == null)
                throw new MacroException(string.Format(
                    "Не найден объект в справочнике файлы по заданному относительному пути '{0}'", relativeFilePath));

            fileObject.GetHeadRevision();

            if (!File.Exists(fileObject.LocalPath))
                throw new MacroException(string.Format(
                    "Не удалось загрузить файл '{0}' из справочника файлы", fileObject.Name));

            return fileObject.LocalPath;
        }

        // короткое имя пользователя
        private string КороткоеИмяПользователя
        {
            get
            {
                if (String.IsNullOrEmpty(_сотрудник))
                    _сотрудник = ТекущийПользователь[User.Fields.ShortName.ToString()];

                return _сотрудник;
            }
        }

        // сокращённое название
        private string НаименованиеПодразделения
        {
            get
            {
                return _подразделениеПользователя != null ? (string)_подразделениеПользователя[Guids.Parameters.СокращенноеНазваниеПодразделения.ToString()] : String.Empty;
            }
        }

        private string КодПодразделения
        {
            get
            {
                return _подразделениеПользователя != null ? (string)_подразделениеПользователя[Guids.Parameters.КодПодразделения.ToString()] : String.Empty;
            }
        }

        public override void Run()
        {
        }

        public void ТолькоВыбранныеОбъекты()
        {
            Выполнение(0);
        }

        public void ВложенныеНаОдинУровень()
        {
            Выполнение(1);
        }

        public void ВложенныеНаВсехУровнях()
        {
            Выполнение(2);
        }

        private void Выполнение(int уровеньВложенности)
        {
            УдалитьВременныеФайлы();

            if (ВыбранныеОбъекты == null)
                return;

            LoadBinaryFiles();

            if (!Directory.Exists(_папкаДляВременныхФайлов))
                Directory.CreateDirectory(_папкаДляВременныхФайлов);

            _списокИдентификаторовДокументов = new List<int>();
            _списокИдентификаторовФайлов = new List<int>();

            ПользовательскийДиалог диалог = ПолучитьПользовательскийДиалог(Guids.UserDialogClass.ToString());
            диалог.Изменить();

            диалог.Caption = "Печать подлинников";
            диалог["Наименование"] = "Печать подлинников";

            // Очищаем список форматов
            foreach (Объект формат in диалог.СвязанныеОбъекты[Guids.Links.Форматы.ToString()])
            {
                формат.Удалить();
            }

            _подразделениеПользователя = диалог.СвязанныйОбъект["Подразделение"];

            // Получаем список принтеров
            _принтеры = НайтиОбъекты("Принтеры", "[ID] > 0");

            WaitingDialog.Show("Идет получение подлинников", true);
            foreach (Объект объектНоменклатуры in ВыбранныеОбъекты)
            {
                ИскатьФайлыВОбъекте(объектНоменклатуры, уровеньВложенности, диалог);
            }
            WaitingDialog.Hide();

            if (диалог.СвязанныйОбъект[Guids.Links.Подразделение.ToString()] == null)
                диалог.СвязанныйОбъект[Guids.Links.Подразделение.ToString()] = _подразделениеПользователя;

            диалог.Сохранить();

            try
            {
                диалог.ПоказатьДиалог();
            }
            finally
            {
                УдалитьВременныеФайлы();
            }
        }

        public void ИзменениеПараметра()
        {
            if (ИзмененныйПараметр == "Исходный файл")
            {
                string исходныйФайл = ТекущийОбъект[Guids.Parameters.ИсходныйФайл.ToString()];
                if (String.IsNullOrEmpty(исходныйФайл))
                {
                    ТекущийОбъект[Guids.Parameters.ФайлДляПечати.ToString()] = String.Empty;
                    return;
                }

                string файлДляПечати = String.Format(@"{0}\{1}_print{2}",
                    _папкаДляВременныхФайлов,
                    Path.GetFileNameWithoutExtension(исходныйФайл),
                    Path.GetExtension(исходныйФайл));
                File.Copy(исходныйФайл, файлДляПечати);
                ТекущийОбъект[Guids.Parameters.ФайлДляПечати.ToString()] = файлДляПечати;
            }
        }

        private void ИскатьФайлыВОбъекте(Объект объектНоменклатуры, int уровеньВложенности, ПользовательскийДиалог диалог)
        {
            if (объектНоменклатуры == null)
                return;

            if (!WaitingDialog.NextStep(String.Format("Поиск подлинников в '{0}'", объектНоменклатуры)))
                return;

            ПоискИзображенийВОбъекте(объектНоменклатуры, диалог);

            if (уровеньВложенности > 0)
            {
                foreach (Объект дочернийОбъект in объектНоменклатуры.ДочерниеОбъекты)
                {
                    if (уровеньВложенности > 1)
                        ИскатьФайлыВОбъекте(дочернийОбъект, уровеньВложенности, диалог);
                    else
                    {
                        if (!WaitingDialog.NextStep(String.Format("Поиск подлинников в '{0}'", дочернийОбъект)))
                            return;

                        ПоискИзображенийВОбъекте(дочернийОбъект, диалог);
                    }
                }
            }
        }

        //Создает список объектов в диалоге ввода, с количеством файлов, с дальнейшим выбором принтера
        private void ПоискИзображенийВОбъекте(Объект объектНоменклатуры, ПользовательскийДиалог диалог)
        {
            Объект документ = объектНоменклатуры.СвязанныйОбъект["Связанный объект"];
            if (документ == null)
                return;

            int idДокумента = документ["ID"];
            if (_списокИдентификаторовДокументов.Contains(idДокумента))
                return;

            _списокИдентификаторовДокументов.Add(idДокумента);

            // Ищем связанные файлы (подлинники) формата tiff в стадии "Хранение"
            List<Объект> файлыTiff = new List<Объект>();
            foreach (Объект файл in документ.СвязанныеОбъекты[EngineeringDocumentFields.File.ToString()])
            {
                int idФайла = файл["ID"];
                if (_списокИдентификаторовФайлов.Contains(idФайла))
                    continue;

                _списокИдентификаторовФайлов.Add(idФайла);

                if (файл["Стадия"] != "Хранение")
                    continue;

                if (ТипФайлаПоддерживается(файл["Наименование"]))
                    файлыTiff.Add(файл);
            }

            if (!файлыTiff.Any())
                return;

            // Определяем номер последнего изменения формата 'ИИ.{номер изменения}'
            int номерПоследнегоИзменения = 0;
            //Сообщение("объектНоменклатуры.СвязанныеОбъекты['Изменения']", "{0}", объектНоменклатуры.СвязанныеОбъекты["Изменения"].Count);
            foreach (Объект изменение in объектНоменклатуры.СвязанныеОбъекты["Изменения"])
            {
                string наименование = изменение["91486563-d044-4045-814b-3432b67812f1"];
                if (наименование.ToLower().StartsWith("ии."))
                {
                    string строка = наименование.Substring(3);
                    int номерИзменения;
                    if (int.TryParse(строка, out номерИзменения))
                    {
                        if (номерИзменения > номерПоследнегоИзменения)
                            номерПоследнегоИзменения = номерИзменения;
                    }
                }
            }

            StringBuilder errors = new StringBuilder();

            foreach (Объект файл in файлыTiff)
            {
                FileReferenceObject file = (FileReferenceObject)файл;
                if (file.IsAdded)
                    continue;

                string filePath = Path.Combine(_папкаДляВременныхФайлов, Path.GetFileName(file.LocalPath));
                file.GetHeadRevision(filePath);
                if (!File.Exists(filePath))
                    continue;

                string относительныйПуть = файл["Относительный путь"];

                Объект учетнаяКарточка = документ.СвязанныйОбъект[Guids.Links.ДокументУчетнаяКарточка.ToString()];
                string инвентарныйНомер = учетнаяКарточка != null ? (string)учетнаяКарточка["Инвентарный номер"] : String.Empty;

                Dictionary<string, List<string>> форматыСФайлами = GetFormatsWithFiles(filePath);
                DeleteFile(filePath);

                foreach (KeyValuePair<string, List<string>> форматСФайлами in форматыСФайлами)
                {
                    string имяФормата = форматСФайлами.Key;
                    Объект форматНаПечать = диалог.СвязанныеОбъекты[Guids.Links.Форматы.ToString()].FirstOrDefault(форматСписка => (string)форматСписка["Формат"] == имяФормата);
                    if (форматНаПечать == null)
                    {
                        // создаем объект диалога в списке "Форматы на печать"
                        форматНаПечать = диалог.СоздатьОбъектСписка(Guids.Links.Форматы.ToString(), "8f9b47d8-2abf-4b1d-b85c-4c3dd370d9ca");
                        форматНаПечать["Формат"] = имяФормата;

                        if (!String.IsNullOrEmpty(имяФормата))
                        {
                            string имяПринтера = String.Empty;

                            if (_списокФорматПринтер.ContainsKey(имяФормата))
                            {
                                имяПринтера = _списокФорматПринтер[имяФормата];
                            }
                            else
                            {
                                foreach (Объект принтерПоФормату in _принтеры.Where(принтер => принтер.СвязанныеОбъекты["Форматы"].FirstOrDefault(форматПринтера => форматПринтера["Наименование"] == имяФормата) != null))
                                {
                                    foreach (Объект подразделение in принтерПоФормату.СвязанныеОбъекты["Подразделения"])
                                    {
                                        if (_подразделениеПользователя == null)
                                        {
                                            UserReferenceObject userUnit = (UserReferenceObject)подразделение;
                                            if (userUnit.GetAllInternalUsers().FirstOrDefault(user => user == (User)ТекущийПользователь) != null)
                                            {
                                                _подразделениеПользователя = подразделение;
                                            }
                                        }

                                        if (_подразделениеПользователя != null && подразделение["ID"] == _подразделениеПользователя["ID"])
                                        {
                                            имяПринтера = принтерПоФормату["Наименование"];
                                            break;
                                        }
                                    }

                                    if (!String.IsNullOrEmpty(имяПринтера))
                                        break;
                                }

                                _списокФорматПринтер.Add(имяФормата, имяПринтера);
                            }

                            форматНаПечать["Принтер"] = имяПринтера;
                        }
                    }
                    else
                    {
                        форматНаПечать.Изменить();
                    }

                    foreach (string исходныйФайл in форматСФайлами.Value)
                    {
                        // Создаем объекты в списке Файлы на печать
                        Объект файлНаПечать = форматНаПечать.СоздатьОбъектСписка("ccb22c57-4a7c-4e17-82a4-3d123ceacae9", "0b9ea13c-d8e5-4c05-84c0-42f9c99e540a");
                        файлНаПечать["Наименование"] = относительныйПуть;

                        // Событие "Завершение изменения параметра"
                        // вызов метода ИзменениеПараметра()
                        файлНаПечать[Guids.Parameters.ИсходныйФайл.ToString()] = исходныйФайл;

                        файлНаПечать["Инвентарный номер"] = инвентарныйНомер;
                        if (номерПоследнегоИзменения > 0)
                            файлНаПечать["Изменение"] = номерПоследнегоИзменения;
                        файлНаПечать.Сохранить();
                    }

                    форматНаПечать["Количество листов"] = форматНаПечать.СвязанныеОбъекты["ccb22c57-4a7c-4e17-82a4-3d123ceacae9"].Count;

                    форматНаПечать.Сохранить();
                }
            }

            WriteLog(errors.ToString());
        }

        public void Распечатать()
        {
            // ТекущийОбъект - ПользовательскийДиалог

            var списокФорматовНаПечать = ТекущийОбъект.СвязанныеОбъекты[Guids.Links.Форматы.ToString()]
                .Where(форматНаПечать => форматНаПечать["Печатать"] == true)
                .ToArray();

            if (списокФорматовНаПечать.Length == 0)
            {
                Сообщение("Внимание", "Не выбраны форматы для печати!");
                return;
            }

            string рабочееМесто = Context.Connection.ClientView.ToString();
            _подразделениеПользователя = ТекущийОбъект.СвязанныйОбъект[Guids.Links.Подразделение.ToString()];

            _currentDateTime = DateTime.Now;

            foreach (Объект форматДляПечати in списокФорматовНаПечать)
            {
                string имяПринтера = форматДляПечати["Принтер"];
                if (String.IsNullOrEmpty(имяПринтера))
                    continue;

                short количествоКопий = форматДляПечати["Количество копий"];
                if (количествоКопий <= 0)
                    continue;

                //Получаем файлы данного формата
                _файлыНаПечать = new List<Объект>();
                foreach (Объект файлНаПечать in форматДляПечати.СвязанныеОбъекты["ccb22c57-4a7c-4e17-82a4-3d123ceacae9"])
                {
                    if (файлНаПечать["Печатать"] == false)
                        continue;

                    string filePath = файлНаПечать[Guids.Parameters.ФайлДляПечати.ToString()];
                    if (ТипФайлаПоддерживается(filePath))
                        _файлыНаПечать.Add(файлНаПечать);
                }

                if (!_файлыНаПечать.Any())
                    continue;

                //Печать страниц указанного формата на указанном принтере
                if (ПечататьНаПринтере(ref имяПринтера, ref количествоКопий))
                {
                    Dictionary<string, int> распечатанныеФайлы = new Dictionary<string, int>();

                    foreach (Объект файлНаПечать in _файлыНаПечать)
                    {
                        string наименованиеПодлинника = файлНаПечать["Наименование"];
                        if (распечатанныеФайлы.ContainsKey(наименованиеПодлинника))
                        {
                            int количествоРаспечатано = распечатанныеФайлы[наименованиеПодлинника];
                            распечатанныеФайлы[наименованиеПодлинника] = количествоРаспечатано + количествоКопий;
                        }
                        else
                            распечатанныеФайлы.Add(наименованиеПодлинника, количествоКопий);
                    }

                    foreach (KeyValuePair<string, int> распечатанныйФайл in распечатанныеФайлы)
                    {
                        //Справочник "Распечатки", тип "Распечатка"
                        Объект распечатка = СоздатьОбъект("4c4383c3-8c9e-474f-a54b-eba92e4452d6", "6bb5e90d-e205-4025-9128-c565fb04264f");
                        // Наименование файла
                        распечатка["982068f0-f524-4ab6-8602-bc5d07802df2"] = распечатанныйФайл.Key;
                        // Количество
                        распечатка["b544aaa3-508c-4488-9d12-091d62f0b535"] = распечатанныйФайл.Value;
                        распечатка["63432aaf-2431-4e99-ac20-98c2ed04c6b0"] = имяПринтера;
                        распечатка["b680711c-1b2e-412a-b20d-8c43c0e81e91"] = форматДляПечати["Формат"];
                        распечатка["26587482-58f4-420b-b881-5bf119e815b2"] = рабочееМесто;
                        распечатка.Сохранить();
                    }
                }
            }
        }

        public List<Объект> ПолучитьСписокФорматов(string pathToDocument)
        {
            var pages = DocumentPageInfo.GetDocumentPagesInfo(pathToDocument);
            return DefineFormatsInPageFormatsReference(pages, true);
        }

        public void ПоставитьШтампы()
        {
            LoadBinaryFiles();

            List<Объект> списокФорматовНаПечать = ТекущийОбъект.СвязанныеОбъекты[Guids.Links.Форматы.ToString()].ToList();
            if (!списокФорматовНаПечать.Any())
                return;

            _currentDateTime = DateTime.Now;
            _подразделениеПользователя = ТекущийОбъект.СвязанныйОбъект[Guids.Links.Подразделение.ToString()];

            _stampColor = ТекущийОбъект["Цвет"];
            _stampOpacity = ТекущийОбъект["Видимость"];
            _stampGlobalX = ТекущийОбъект["X"];
            _stampGlobalY = ТекущийОбъект["Y"];
            _stampGlobalAngle = ТекущийОбъект["Угол поворота"];

            StringBuilder errors = new StringBuilder();

            WaitingDialog.Show("Проставление штампа", true);
            foreach (Объект форматДляПечати in списокФорматовНаПечать)
            {
                foreach (Объект файлНаПечать in форматДляПечати.СвязанныеОбъекты["ccb22c57-4a7c-4e17-82a4-3d123ceacae9"])
                {
                    string filePath = файлНаПечать[Guids.Parameters.ФайлДляПечати.ToString()];
                    if (ТипФайлаПоддерживается(filePath))
                    {
                        if (!WaitingDialog.NextStep(String.Format("Заполнение штампа для '{0}'.", Path.GetFileName(filePath))))
                            return;

                        ПоставитьШтамп(файлНаПечать, ref errors);
                    }
                }
            }
            WaitingDialog.Hide();

            WriteLog(errors.ToString());
        }

        public void ПоставитьОтдельныйШтамп()
        {
            LoadBinaryFiles();

            _currentDateTime = DateTime.Now;
            _подразделениеПользователя = ТекущийОбъект.Владелец.Владелец.СвязанныйОбъект[Guids.Links.Подразделение.ToString()];

            _stampColor = ТекущийОбъект.Владелец.Владелец["Цвет"];
            _stampOpacity = ТекущийОбъект.Владелец.Владелец["Видимость"];
            _stampGlobalX = ТекущийОбъект.Владелец.Владелец["X"];
            _stampGlobalY = ТекущийОбъект.Владелец.Владелец["Y"];
            _stampGlobalAngle = ТекущийОбъект.Владелец.Владелец["Угол поворота"];

            StringBuilder errors = new StringBuilder();

            string filePath = ТекущийОбъект[Guids.Parameters.ФайлДляПечати.ToString()];
            if (ТипФайлаПоддерживается(filePath))
            {
                WaitingDialog.Show("Проставление штампа", true);
                if (!WaitingDialog.NextStep(String.Format("Заполнение штампа для '{0}'.", Path.GetFileName(filePath))))
                    return;

                ПоставитьШтамп(ТекущийОбъект, ref errors);

                WaitingDialog.Hide();
            }

            WriteLog(errors.ToString());
        }

        private bool ТипФайлаПоддерживается(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            var extension = Path.GetExtension(filePath).ToLower();
            return SupportedExtensions.Contains(extension);
        }

        private bool ПечататьНаПринтере(ref string наименованиеПринтера, ref short количествоКопий)
        {
            _текущийНомерПечатаемогоФайла = 0;

            using (var printDocument = new PrintDocument())
            {
                var printerSettings = new PrinterSettings
                {
                    PrinterName = наименованиеПринтера,
                    Copies = количествоКопий
                };

                printDocument.PrinterSettings = printerSettings;
                printDocument.PrintPage += PD_PrintPage;

                try
                {
                    if (printerSettings.IsValid)
                    {
                        printDocument.Print();
                        return true;
                    }

                    using (var printDialog = new PrintDialog
                    {
                        Document = printDocument,
                        PrinterSettings = printerSettings
                    })
                    {
                        if (printDialog.ShowDialog() != DialogResult.OK)
                            return false;

                        // сколько копий было указано в диалоге
                        количествоКопий = printDialog.PrinterSettings.Copies;
                        // на каком принтере в итоге распечатали
                        наименованиеПринтера = printDialog.PrinterSettings.PrinterName;

                        printDocument.Print();
                        return true;
                    }
                }
                catch (Exception e)
                {
                    Сообщение("Ошибка печати", e.Message);
                    return false;
                }
                finally
                {
                    printDocument.PrintPage -= PD_PrintPage;
                }
            }
        }

        private void PD_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (!_файлыНаПечать.Any())
                return;

            if (_файлыНаПечать.Count <= _текущийНомерПечатаемогоФайла)
                return;

            e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
            e.Graphics.CompositingQuality = CompositingQuality.HighQuality;
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            e.Graphics.PageUnit = GraphicsUnit.Pixel;

            var файл = _файлыНаПечать[_текущийНомерПечатаемогоФайла];
            using (Image image = Image.FromFile(файл[Guids.Parameters.ФайлДляПечати.ToString()]))
            {
                double width = image.Width;
                double height = image.Height;
                if (width > height)
                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);

                float s1 = e.Graphics.VisibleClipBounds.Width / image.Width;
                float s2 = e.Graphics.VisibleClipBounds.Height / image.Height;
                if (s1 > s2)
                    s1 = s2;

                int w = (int)(image.Width * s1);
                int h = (int)(image.Height * s1);
                int l = Convert.ToInt32(e.Graphics.VisibleClipBounds.Left + ((e.Graphics.VisibleClipBounds.Width - w) / 2));
                int t = Convert.ToInt32(e.Graphics.VisibleClipBounds.Top + ((e.Graphics.VisibleClipBounds.Height - h) / 2));

                e.PageSettings.Margins = new Margins(0, 0, 0, 0);

                e.Graphics.DrawImage(image,
                    new Rectangle(l, t, w, h),
                    new Rectangle(0, 0, image.Width, image.Height),
                    GraphicsUnit.Pixel);
            }

            _текущийНомерПечатаемогоФайла++;
            e.HasMorePages = _файлыНаПечать.Count > _текущийНомерПечатаемогоФайла;
        }

        public void УдалитьВременныеФайлы()
        {
            if (!Directory.Exists(_папкаДляВременныхФайлов))
                return;

            foreach (string filePath in Directory.EnumerateFiles(_папкаДляВременныхФайлов))
            {
                DeleteFile(filePath);
            }
        }

        private void DeleteFile(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Exists)
                {
                    fileInfo.IsReadOnly = false;
                    fileInfo.Delete();
                }
            }
            catch
            {
                // ignored
            }
        }

        private void ПоставитьШтамп(Объект файлНаПечать, ref StringBuilder errors)
        {
            string инвНомерПодлинника = файлНаПечать["2bc188a8-395a-4aa2-a7e4-8101b442c838"];
            string номерИзменения = файлНаПечать["512304e1-9a84-456c-89f7-87aad4302b6b"];

            string stampText = String.Format(
        @"КОПИЯ
с электр. подлинника инв. №{0}
{1}{2}
{3}{4}, {5}",
        инвНомерПодлинника,
        !String.IsNullOrEmpty(номерИзменения) ? String.Format("Изм. {0}, ", номерИзменения) : String.Empty,
        НаименованиеПодразделения,
        !String.IsNullOrEmpty(КодПодразделения) ? String.Format("Код подр.{0}, ", КодПодразделения) : String.Empty,
        КороткоеИмяПользователя,
        _currentDateTime.ToString("dd.MM.yyyy HH:mm"));

            string исходныйФайл = файлНаПечать[Guids.Parameters.ИсходныйФайл.ToString()];
            if (String.IsNullOrEmpty(исходныйФайл))
                Error("Не задан исходный файл");

            byte[] colors = BitConverter.GetBytes(_stampColor);

            bool отдельнаяНастройка = файлНаПечать["Отдельная настройка"];

            double stampX = отдельнаяНастройка ? (double)файлНаПечать["X"] : _stampGlobalX;
            double stampY = отдельнаяНастройка ? (double)файлНаПечать["Y"] : _stampGlobalY;
            int stampAngle = отдельнаяНастройка ? (int)файлНаПечать["Угол поворота"] : _stampGlobalAngle;

            string[] fileInfo = GetFileInfo(исходныйФайл, ref errors);

            int horisontalResolution = 300;
            int verticalResolution = 300;
            if (fileInfo.Length == 4)
            {
                int.TryParse(fileInfo[2], out horisontalResolution);
                if (horisontalResolution == 0)
                    horisontalResolution = 300;

                int.TryParse(fileInfo[3], out verticalResolution);
                if (verticalResolution == 0)
                    verticalResolution = 300;
            }

            int stampHeight = verticalResolution - (2 * _stampBorderWidth);

            // рисуем штамп с рамкой и поворотом
            string stampPath = String.Format(@"{0}\{1}_stamp.png", _папкаДляВременныхФайлов, Path.GetFileNameWithoutExtension(исходныйФайл));
            string cmd = String.Format("-background none -fill rgb(\"{2},{3},{4}\") -size x{5} -font \"{6}\" -gravity center label:\"{0}\" -bordercolor rgb(\"{2},{3},{4}\") -compose Copy -border {7} -rotate -{8} \"{1}\"",
                stampText,
                stampPath,
                colors[2],
                colors[1],
                colors[0],
                stampHeight,
                _fontName,
                _stampBorderWidth,
                stampAngle
                );


            StartProcess(_localMagickFilePath, cmd, ref errors);

            string файлДляПечати = файлНаПечать[Guids.Parameters.ФайлДляПечати.ToString()];
            int opacity = (int)(_stampOpacity * 100);

            // определяем координаты штампа
            int x = (int)((stampX / 25.4) * horisontalResolution);
            int y = (int)((stampY / 25.4) * verticalResolution);

            // накладываем штамп на исходные изображения
            //const string isQuietMode = "-quiet"; // не выводить предупреждения
            cmd = string.Format("{6}\"{0}\" \"{1}\" -compose dissolve -define compose:args=\"{2}\" -geometry +{3}+{4} -composite \"{5}\"",
                исходныйФайл, stampPath, opacity, x, y, файлДляПечати, string.Empty);

            StartProcess(_localMagickFilePath, cmd, ref errors);
        }

        private Dictionary<string, List<string>> GetFormatsWithFiles(string pathToDocument)
        {
            var formatWithFiles = new Dictionary<string, List<string>>();

            try
            {
                var pages = DocumentPageInfo.ParseDocumentToTiffFiles(pathToDocument, _папкаДляВременныхФайлов);
                DefineFormatsInPageFormatsReference(pages);

                foreach (var page in pages.Where(p => !string.IsNullOrWhiteSpace(p.FormatPage)))
                {
                    if (formatWithFiles.ContainsKey(page.FormatPage))
                    {
                        formatWithFiles[page.FormatPage].Add(page.FileInfo.FullName);
                    }
                    else
                    {
                        formatWithFiles.Add(page.FormatPage, new List<string> { page.FileInfo.FullName });
                    }
                }
            }
            catch (Exception e)
            {
                Message("Ошибка", string.Format("Ошибка получения форматов из подлинника '{0}':{1}{2}",
                    pathToDocument,
                    Environment.NewLine,
                    e.Message));
            }

            return formatWithFiles;
        }

        private List<Объект> DefineFormatsInPageFormatsReference(DocumentPageInfo[] pagesInfo, bool fillListFormats = false)
        {
            if (pagesInfo.Length == 0)
                return new List<Объект>();

            var referenceInfo = Context.Connection.ReferenceCatalog.Find(Guids.PageFormatsReferenceId);

            if (referenceInfo == null)
                throw new MacroException("Не найден справочник 'Форматы'");

            var reference = referenceInfo.CreateReference();
            // Так объектов в справочнике 'Форматы' не много, то грузим все на клиент
            reference.Objects.Load();

            var allFormats = reference.Objects.ToArray();
            var formats = new List<ReferenceObject>();

            foreach (var pageInfo in pagesInfo)
            {
                var format = allFormats.FirstOrDefault(f =>
                {
                    var width = f[Guids.WidthFormat].GetInt32();
                    var height = f[Guids.HeightFormat].GetInt32();

                    int pageWidth;
                    int pageHeight;

                    if (pageInfo.IsLandscape)
                    {
                        pageWidth = pageInfo.HeightMm;
                        pageHeight = pageInfo.WidthMm;
                    }
                    else
                    {
                        pageWidth = pageInfo.WidthMm;
                        pageHeight = pageInfo.HeightMm;
                    }

                    return width >= pageWidth - 5 &&
                           width <= pageWidth + 5 &&
                           height >= pageHeight - 5 &&
                           height <= pageHeight + 5;
                });

                if (format == null)
                    continue;

                var formatName = format[Guids.NameFormat].GetString();

                if (string.IsNullOrWhiteSpace(formatName))
                    continue;

                pageInfo.FormatPage = formatName;

                if (fillListFormats)
                    formats.Add(format);
            }

            return fillListFormats
                ? formats.Select(format => Объект.CreateInstance(format, Context)).ToList()
                : new List<Объект>();
        }


        private string[] GetFileInfo(string filePath, ref StringBuilder errors)
        {
            // определяем размеры и разрешение файла
            //identify -format "%w %h %x %y" A_1.tif
            string cmd = String.Format("identify -format \"%w %h %x %y\" \"{0}\"", filePath);
            string result = StartProcess(_localMagickFilePath, cmd, ref errors);
            return result.Split(' ');
        }

        private string StartProcess(string pathToExeFile, string cmd, ref StringBuilder errors)
        {
            errors.AppendLine();
            errors.AppendLine(cmd);

            var process = Process.Start(new ProcessStartInfo
            {
                FileName = pathToExeFile,
                Arguments = cmd,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });

            if (process == null)
            {
                errors.AppendLine("process is null");
                return string.Empty;
            }

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (!string.IsNullOrEmpty(error))
                errors.AppendLine(error);

            return output;
        }

        private static void WriteLog(string errors)
        {
            if (String.IsNullOrEmpty(errors))
                return;

            using (var logWriter = new TFlex.DOCs.Common.LogWriter())
            {
                logWriter.InitializeFileMode(_логОшибок, false);
                logWriter.Write(errors);
            }
        }


        /// <summary> Класс для получения информации о странице документа </summary>
        public class DocumentPageInfo
        {
            /// <summary> Кол-во точек на дюйм для PDF страниц по умолчанию </summary>
            private const int Dpi = 300;

            /// <summary> Коэффициент для приведения точек в дюймы </summary>
            private const int CoefficientInch = 72;

            /// <summary> Коэффициент для приведения дюймов в мм </summary>
            private const double CoefficientMm = 25.42;

            private DocumentPageInfo()
            {
            }

            /// <summary> Информация о файле </summary>
            public FileInfo FileInfo { get; private set; }

            /// <summary> Ширина в пикселях </summary>
            public int WidthPx { get; private set; }

            /// <summary> Высота в пикселях </summary>
            public int HeightPx { get; private set; }

            /// <summary> Кол-во точек на дюйм горизонтали </summary>
            public float XDpi { get; private set; }

            /// <summary> Кол-во точек на дюйм вертикали </summary>
            public float YDpi { get; private set; }

            /// <summary> Ширина в миллиметрах </summary>
            public int WidthMm { get; private set; }

            /// <summary> Высота в миллиметрах </summary>
            public int HeightMm { get; private set; }

            /// <summary> Формат страницы </summary>
            public string FormatPage { get; set; }

            /// <summary> Ориентация страницы, если true то альбомная иначе книжная </summary>
            public bool IsLandscape { get; private set; }

            #region Получить информацию о страницах документа

            public static DocumentPageInfo[] GetDocumentPagesInfo(string filePath)
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    throw new ArgumentNullException("filePath");

                var file = new FileInfo(filePath);

                if (!file.Exists)
                    throw new MacroException(string.Format("Не найден файл по заданному пути '{0}'", filePath));

                string fileExtension = file.Extension.ToLower();

                if (fileExtension == ".pdf")
                    return GetDocumentPagesInfoFromPdf(file);

                if (fileExtension == ".tif" || fileExtension == ".tiff")
                    return GetDocumentPagesInfoFromTiff(file);

                throw new MacroException(string.Format(
                    "Формат файла по заданному пути '{0}' не поддерживается",
                    filePath));
            }

            private static DocumentPageInfo[] GetDocumentPagesInfoFromPdf(FileInfo file)
            {
                var pages = new List<DocumentPageInfo>();

                using (var processor = new PdfDocumentProcessor())
                {
                    processor.LoadDocument(file.FullName);

                    foreach (var page in processor.Document.Pages)
                    {
                        var pageInfo = GetDocumentPageInfoFromPdfPage(page, file);
                        pages.Add(pageInfo);
                    }
                }

                return pages.ToArray();
            }

            private static DocumentPageInfo[] GetDocumentPagesInfoFromTiff(FileInfo file)
            {
                var pages = new List<DocumentPageInfo>();

                using (var image = Image.FromFile(file.FullName))
                {
                    var totalCount = image.GetFrameCount(FrameDimension.Page);

                    for (var i = 0; i < totalCount; i++)
                    {
                        image.SelectActiveFrame(FrameDimension.Page, i);

                        var pageInfo = GetDocumentPageInfoFromImage(image, file);
                        pages.Add(pageInfo);
                    }
                }

                return pages.ToArray();
            }

            #endregion

            #region Разбор документа по TIFF-файлам

            public static DocumentPageInfo[] ParseDocumentToTiffFiles(string filePath, string outPath)
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    throw new ArgumentNullException("filePath");

                if (string.IsNullOrWhiteSpace(outPath))
                    throw new ArgumentNullException("outPath");

                var file = new FileInfo(filePath);
                if (!file.Exists)
                    throw new MacroException(string.Format("Не найден файл по заданному пути '{0}'", filePath));

                if (!Directory.Exists(outPath))
                    throw new MacroException(string.Format("Не найден каталог по заданному пути '{0}'", outPath));

                string fileExtension = file.Extension.ToLower();

                if (fileExtension == ".pdf")
                    return ParsePdfFileToTiffFiles(file, outPath);

                if (fileExtension == ".tif" || fileExtension == ".tiff")
                    return ParseTiffFileToTiffFiles(file, outPath);

                throw new MacroException(string.Format(
                    "Формат файла по заданному пути '{0}' не поддерживается",
                    filePath));
            }

            private static DocumentPageInfo[] ParsePdfFileToTiffFiles(FileInfo file, string outPath)
            {
                var pages = new List<DocumentPageInfo>();

                using (var processor = new PdfDocumentProcessor())
                {
                    processor.LoadDocument(file.FullName);

                    var numberPage = 1;
                    var originaFileName = Path.GetFileNameWithoutExtension(file.Name);

                    foreach (var page in processor.Document.Pages)
                    {
                        var pageInfo = GetDocumentPageInfoFromPdfPage(page);

                        var leg = pageInfo.WidthPx > pageInfo.HeightPx
                            ? pageInfo.WidthPx
                            : pageInfo.HeightPx;

                        var newFilePath = Path.Combine(outPath,
                            string.Format("{0}_(pdf)_л{1}.tif", originaFileName, numberPage));

                        var image = processor.CreateBitmap(numberPage, leg);
                        image.SetResolution(pageInfo.XDpi, pageInfo.YDpi);
                        image.Save(newFilePath, ImageFormat.Tiff);

                        pageInfo.FileInfo = new FileInfo(newFilePath);
                        pages.Add(pageInfo);
                        numberPage++;
                    }
                }

                return pages.ToArray();
            }

            private static DocumentPageInfo[] ParseTiffFileToTiffFiles(FileInfo file, string outPath)
            {
                var pages = new List<DocumentPageInfo>();

                var numberPage = 1;
                var originaFileName = Path.GetFileNameWithoutExtension(file.Name);

                using (var image = Image.FromFile(file.FullName))
                {
                    var totalCount = image.GetFrameCount(FrameDimension.Page);

                    for (var i = 0; i < totalCount; i++)
                    {
                        image.SelectActiveFrame(FrameDimension.Page, i);

                        var pageInfo = GetDocumentPageInfoFromImage(image);

                        var newFilePath = Path.Combine(outPath,
                            string.Format("{0}_(tif)_л{1}.tif", originaFileName, numberPage));

                        image.Save(newFilePath, ImageFormat.Tiff);
                        pageInfo.FileInfo = new FileInfo(newFilePath);

                        pages.Add(pageInfo);
                        numberPage++;
                    }
                }

                return pages.ToArray();
            }

            #endregion

            #region Получить информацию о странице документа из файла

            private static DocumentPageInfo GetDocumentPageInfoFromPdfPage(PdfPage page, FileInfo file = null)
            {
                var docPage = new DocumentPageInfo
                {
                    FileInfo = file,
                    WidthPx = Units.PointsToPixels((int)page.MediaBox.Width, Dpi),
                    HeightPx = Units.PointsToPixels((int)page.MediaBox.Height, Dpi),
                    XDpi = Dpi,
                    YDpi = Dpi,
                    WidthMm = (int)(page.MediaBox.Width / CoefficientInch * CoefficientMm),
                    HeightMm = (int)(page.MediaBox.Height / CoefficientInch * CoefficientMm)
                };

                docPage.IsLandscape = docPage.WidthMm > docPage.HeightMm;

                return docPage;
            }

            private static DocumentPageInfo GetDocumentPageInfoFromImage(Image image, FileInfo file = null)
            {
                var docPage = new DocumentPageInfo
                {
                    FileInfo = file,
                    WidthPx = image.Width,
                    HeightPx = image.Height,
                    XDpi = image.HorizontalResolution > 0 ? image.HorizontalResolution : 300,
                    YDpi = image.VerticalResolution > 0 ? image.VerticalResolution : 300,
                    WidthMm = (int)(image.Width / (double)image.HorizontalResolution * 25.42),
                    HeightMm = (int)(image.Height / (double)image.VerticalResolution * 25.42)
                };

                docPage.IsLandscape = docPage.WidthMm > docPage.HeightMm;

                return docPage;
            }

            #endregion

        }
    }
}

