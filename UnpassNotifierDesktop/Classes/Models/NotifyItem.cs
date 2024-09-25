namespace UnpassNotifierDesktop.Classes;

public class NotifyItem
{
    public string FIO { get; set; }
    public List<UnpassItem> UnpassedList { get; set; } = [];
}