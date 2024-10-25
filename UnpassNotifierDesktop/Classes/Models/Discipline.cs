namespace UnpassNotifierDesktop.Classes;

public class Discipline
{
    public string DisciplineName { get; set; }
    public string? AttestationDate { get; set; }

    public string TypeControl => DisciplineName.Contains("зачтено") ? "Зачёт" : "Зачёт с оценкой/Экзамен";

    public Discipline()
    {
    }

    public Discipline(string disciplineName)
    {
        DisciplineName = disciplineName;
    }

    public Discipline(string disciplineName, string attestationDate, string typeControl)
    {
        DisciplineName = disciplineName;
        AttestationDate = attestationDate;
        // TypeControl = typeControl;
    }
    
    
    public override string ToString()
    {
        return DisciplineName;
    }
}