
namespace UnpassNotifier.Classes;

public class Discipline
{
    public string DisciplineName { get; set; }
    public string AttestationDate { get; set; }
    public string TypeControl { get; set; }

    public Discipline()
    {
        
    }
    public Discipline(string disciplineName, string attestationDate, string typeControl)
    {
        DisciplineName = disciplineName;
        AttestationDate = attestationDate;
        TypeControl = typeControl;
    }
}