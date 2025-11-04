using System;

namespace ReportGeneration.Models
{
    public class Work
    {
        public int Id { get; set; }
        public int IdDiscipline { get; set; }
        public int IdType { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public int Semester { get; set; }

        public Work(int id, int idDiscipline, int idType, DateTime date, string name, int semester)
        {
            Id = id;
            IdDiscipline = idDiscipline;
            IdType = idType;
            Date = date;
            Name = name;
            Semester = semester;
        }
    }
}
