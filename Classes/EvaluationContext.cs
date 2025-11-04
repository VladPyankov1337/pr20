using MySql.Data.MySqlClient;
using ReportGeneration.Classes.Common;
using ReportGeneration.Models;
using System.Collections.Generic;

namespace ReportGeneration.Classes
{
    public class EvaluationContext : Evaluation
    {
        public EvaluationContext(int id, int idWork, int idStudent, string value, string lateness) : base(id, idWork, idStudent, value, lateness) { }

        public static List<EvaluationContext> AllEvaluations()
        {
            List<EvaluationContext> allEvaluations = new List<EvaluationContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDEvaluations = Connection.Query("SELECT * FROM `evaluation`", connection);
            while (BDEvaluations.Read())
            {
                allEvaluations.Add(new EvaluationContext(
                    BDEvaluations.GetInt32(0),
                    BDEvaluations.GetInt32(1),
                    BDEvaluations.GetInt32(2),
                    BDEvaluations.GetString(3),
                    BDEvaluations.GetString(4)
                    ));
            }
            Connection.CloseConnection(connection);
            return allEvaluations;
        }
    }
}
