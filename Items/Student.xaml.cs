using ReportGeneration.Classes;
using ReportGeneration.Pages;
using System.Windows.Controls;

namespace ReportGeneration.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public Student(StudentContext student, Main main)
        {
            InitializeComponent();
        }
    }
}
