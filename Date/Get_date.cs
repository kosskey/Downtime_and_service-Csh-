namespace Date;
public class Get_date
{
    private string d_reverse;
    public string day;
    public string month;
    public string yaer;
    public string d_full;
    public string d_briefly;
    public Dictionary<string, string> month_name = new Dictionary<string, string>()
    {
        ["01"] = "Январь",
        ["02"] = "Февраль",
        ["03"] = "Март",
        ["04"] = "Апрель",
        ["05"] = "Май",
        ["06"] = "Июнь",
        ["07"] = "Июль",
        ["08"] = "Август",
        ["09"] = "Сентябрь",
        ["10"] = "Октябрь",
        ["11"] = "Ноябрь",
        ["12"] = "Декабрь"
    };

    public Get_date(string d_reverse)
    {
        this.d_reverse = d_reverse;
        day = d_reverse.Substring(8, 2);
        month = d_reverse.Substring(5, 2);
        yaer = d_reverse.Substring(0, 4);
        d_full = day + "." + month + "." + yaer;
        d_briefly = day + "." + month;
    }

    public string Day()
    {
        return d_reverse.Substring(8, 2);
    }
    
    public string Month()
    {
        return d_reverse.Substring(5, 2);
    }

    public string Year()
    {
        return d_reverse.Substring(0, 4);
    }

    public string D_full()
    {
        return day + "." + month + "." + yaer;
    }

    public string D_briefly()
    {
        return day + "." + month;
    }
}
