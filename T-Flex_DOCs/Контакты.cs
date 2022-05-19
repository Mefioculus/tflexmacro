using TFlex.DOCs.Model.Macros;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {      
    }

    public void Сохранение()
    {
        string ФИО = Параметр["Фамилия"].ToString() + " ";
        string имя = Параметр["Имя"].ToString().Trim();
        if (имя.Length > 0)
            ФИО += имя[0] + ".";

        string отчество = Параметр["Отчество"].ToString().Trim();
        if (отчество.Length > 0)
            ФИО += отчество[0] + ".";

        Параметр["ФИО"] = ФИО;

        string приветствие = Параметр["Приветствие"].ToString();
        if (приветствие.Length == 0)
        {
            приветствие = "Уважаемый " + имя + " " + отчество;
            Параметр["Приветствие"] = приветствие;
        }
    }
}
