#region Рекурсивное получение списка изделий для случайной ДСЕ
// Метод, который будет рекурсивно просматривать родителей дсе для получения списка всех изделий
private Объекты ПолучитьРодительскиеИзделияДляДСЕ (Объект дсе) {
    Объекты списокИзделий = new Объекты();

    Объекты родители = дсе.РодительскиеОбъекты;

    foreach (Объект родитель in родители) {
        // Проверяем, не является ли данный родитель изделием
        if ((родитель.РодительскиеОбъекты.Count == 0) || (СодержитПапку(родитель))) {
            списокИзделий.Add(родитель);
            // Данная ветвь будет срабатывать в том случае, если данное две является изделием (концом ветки)
        }
        else {
            // Данная ветвь будет срабатываеть, если данное изделие является промежуточным звеном
            foreach (Объект изделие in ПолучитьРодительскиеИзделияДляДСЕ(родитель)) {
                if (!списокИзделий.Contains(изделие)) {
                    списокИзделий.Add(изделие);
                }
            }

        }

    }

    return списокИзделий;
}

private bool СодержитПапку(Объект дсе) {
    foreach (Объект родитель in дсе.РодительскиеОбъекты) {
        if (родитель.Тип == "Папка") {
            return true;
        }
    }
    return false;
}
#endregion Рекурсивное получение списка изделий для случайной ДСЕ

#region Получение списка цехопереходов, которые наследуют свои доступы


// Для работы данного кода потребуются так же следующие пространства имен:
//using TFlex.DOCs.Model.References;
//using TFlex.DOCs.Model.Access;
//using TFlex.DOCs.Model.Classes;
// А так же возможно ссылки на следующие библиотеки
//Cs.UI.Objects.dll
//TFlex.DOCs.UI.Common.dll
//TFlex.DOCs.UI.Types.dll
//TFlex.DOCs.Common.dll


Reference techReference = Context.Connection.ReferenceCatalog.Find(new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb")).CreateReference();

// Получаем список технологический процессов
techReference.Objects.Load();

string message = "Список цехопереходов с явными правами:";

foreach (ReferenceObject process in techReference.Objects) {
    // Получаем список цехопереходов
    foreach (ReferenceObject perehod in process.Children) {
        AccessManager accessManager = AccessManager.GetReferenceObjectAccess(perehod);
        if (!accessManager.IsInherit) {
            message += string.Format("\nТП (id {0}): {1}; ЦЗ (id {2}): {3}", process.SystemFields.Id.ToString(), process.ToString(), perehod.SystemFields.Id.ToString(), perehod.ToString());
        }
    }
}

Message("Информация", message);

#endregion Получение списка цехопереходов, которые наследуют свои доступы
