using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Nomenclature;

namespace Technology
{
    public class TechProcessMacros : MacroProvider
    {
        public TechProcessMacros(MacroContext context) : base(context)
        {
        }

        public void OnCreateTechProcess()
        {
            if (!Context.Reference.IsSlave)
                return;

            if (Context.ReferenceObject.MasterObject.Reference is not NomenclatureReference)
                return;

            Объект изделие = ТекущийОбъект.Владелец;
            if (изделие is null)
                return;

            ТекущийОбъект["Масса изготавливаемой детали"] = изделие["Масса"];
            ТекущийОбъект["Наименование"] = изделие["Наименование"];
            ТекущийОбъект["Обозначение ТП"] = изделие["Обозначение"];

            Объекты файлы = изделие.СвязанныеОбъекты["[Связанный объект].[Документы]->[Файлы]"];

            if (файлы.Count == 0)
                return;

            Объект файл = файлы.FirstOrDefault();

            if (файл is null)
                return;

            ТекущийОбъект.СвязанныйОбъект["9bc695f9-fa6f-4bfd-bc6f-9229b75938da"] = файл;
        }
    }
}
