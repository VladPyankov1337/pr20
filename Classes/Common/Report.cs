using Microsoft.Win32;
using ReportGeneration.Pages;
using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGeneration.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main Main)
        {
            SaveFileDialog SFD = new SaveFileDialog
            {
                InitialDirectory = @"C:\",
                Filter = "Excel (*.xlsx)|*.xlsx"
            };
            SFD.ShowDialog();
            if (SFD.FileName != "")
            {
                GroupContext Group = Main.AllGroups.Find(x => x.Id == IdGroup);
                var ExcelApp = new Excel.Application();
                try
                {
                    ExcelApp.Visible = false;
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;
                    Worksheet.Name = "Общий отчёт";

                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    Styles(Worksheet.Cells[1, 1], 18);

                    (Worksheet.Cells[3, 1] as Excel.Range).Value = $"Список группы:";
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    Styles(Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);
                    
                    (Worksheet.Cells[4, 1] as Excel.Range).Value = $"ФИО";
                    Styles(Worksheet.Cells[4, 1], 12, XlHAlign.xlHAlignCenter, true);
                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0f;

                    (Worksheet.Cells[4, 2] as Excel.Range).Value = $"Кол-во не сданных практических";
                    Styles(Worksheet.Cells[4, 2], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 3] as Excel.Range).Value = $"Кол-во не сданных теоретических";
                    Styles(Worksheet.Cells[4, 3], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 4] as Excel.Range).Value = $"Отсутствовал на паре";
                    Styles(Worksheet.Cells[4, 4], 12, XlHAlign.xlHAlignCenter, true);

                    (Worksheet.Cells[4, 5] as Excel.Range).Value = $"Опоздал";
                    Styles(Worksheet.Cells[4, 5], 12, XlHAlign.xlHAlignCenter, true);

                    int Height = 5;
                    List<StudentContext> Students = Main.AllStudents.FindAll(x => x.IdGroup == IdGroup);
                    StudentContext TopStudent = null;
                    int MinDebts = int.MaxValue;
                    int MinAttendanceIssues = int.MaxValue;

                    foreach (StudentContext Student in Students)
                    {
                        List<DisciplineContext> StudentDisciplines = Main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);
                        int PracticeCount = 0;
                        int TheoryCount = 0;
                        int AbsenteeismCount = 0;
                        int LateCount = 0;

                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);
                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                EvaluationContext Evaluation = Main.AllEvaluations.Find(x =>
                                    x.IdWork == StudentWork.Id && x.IdStudent == Student.Id);
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2")) || Evaluation == null)
                                {
                                    if (StudentWork.IdType == 1)
                                        PracticeCount++;
                                    else if (StudentWork.IdType == 2)
                                        TheoryCount++;
                                }
                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        AbsenteeismCount++;
                                    else
                                        LateCount++;
                                }
                            }
                        }

                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{Student.LastName} {Student.FirstName}";
                        Styles(Worksheet.Cells[Height, 1], 12, XlHAlign.xlHAlignLeft, true);

                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        Styles(Worksheet.Cells[Height, 2], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                        Styles(Worksheet.Cells[Height, 3], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbsenteeismCount.ToString();
                        Styles(Worksheet.Cells[Height, 4], 12, XlHAlign.xlHAlignCenter, true);

                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                        Styles(Worksheet.Cells[Height, 5], 12, XlHAlign.xlHAlignCenter, true);

                        int TotalDebts = PracticeCount + TheoryCount;
                        int TotalAttendanceIssues = AbsenteeismCount * 2 + LateCount;
                        if (TotalDebts < MinDebts || (TotalDebts == MinDebts && TotalAttendanceIssues < MinAttendanceIssues))
                        {
                            MinDebts = TotalDebts;
                            MinAttendanceIssues = TotalAttendanceIssues;
                            TopStudent = Student;
                        }

                        Height++;
                    }

                    if (TopStudent != null)
                    {
                        (Worksheet.Cells[Height + 1, 1] as Excel.Range).Value = $"Самый успешный студент: {TopStudent.LastName} {TopStudent.FirstName}";
                        Worksheet.Range[Worksheet.Cells[Height + 1, 1], Worksheet.Cells[Height + 1, 5]].Merge();
                        Styles(Worksheet.Cells[Height + 1, 1], 12, Excel.XlHAlign.xlHAlignLeft);
                    }

                    GenerateStudentSheets(Workbook, Students, Main);

                    Workbook.SaveAs(SFD.FileName);
                    Workbook.Close();
                }
                catch (Exception exp)
                {
                }
                ExcelApp.Quit();
            }
        }

        private static void GenerateStudentSheets(Excel.Workbook Workbook, List<StudentContext> Students, Main Main)
        {
            foreach (StudentContext Student in Students)
            {
                Excel.Worksheet StudentSheet = Workbook.Sheets.Add();
                StudentSheet.Name = $"{Student.LastName} {Student.FirstName}".Length > 31
                    ? $"{Student.LastName} {Student.FirstName}".Substring(0, 31)
                    : $"{Student.LastName} {Student.FirstName}";

                (StudentSheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт по студенту: {Student.LastName} {Student.FirstName}";
                StudentSheet.Range[StudentSheet.Cells[1, 1], StudentSheet.Cells[1, 4]].Merge();
                Styles(StudentSheet.Cells[1, 1], 16);

                (StudentSheet.Cells[3, 1] as Excel.Range).Value = "Дисциплина";
                Styles(StudentSheet.Cells[3, 1], 12, XlHAlign.xlHAlignCenter, true);
                (StudentSheet.Cells[3, 1] as Excel.Range).ColumnWidth = 30.0f;

                (StudentSheet.Cells[3, 2] as Excel.Range).Value = "Работа";
                Styles(StudentSheet.Cells[3, 2], 12, XlHAlign.xlHAlignCenter, true);
                (StudentSheet.Cells[3, 2] as Excel.Range).ColumnWidth = 30.0f;

                (StudentSheet.Cells[3, 3] as Excel.Range).Value = "Тип работы";
                Styles(StudentSheet.Cells[3, 3], 12, XlHAlign.xlHAlignCenter, true);

                (StudentSheet.Cells[3, 4] as Excel.Range).Value = "Статус";
                Styles(StudentSheet.Cells[3, 4], 12, XlHAlign.xlHAlignCenter, true);

                int Row = 4;
                List<DisciplineContext> StudentDisciplines = Main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);

                foreach (DisciplineContext Discipline in StudentDisciplines)
                {
                    List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == Discipline.Id);
                    foreach (WorkContext Work in StudentWorks)
                    {
                        EvaluationContext Evaluation = Main.AllEvaluations.Find(x =>
                            x.IdWork == Work.Id && x.IdStudent == Student.Id);

                        (StudentSheet.Cells[Row, 1] as Excel.Range).Value = Discipline.Name;
                        Styles(StudentSheet.Cells[Row, 1], 12, XlHAlign.xlHAlignLeft, true);

                        (StudentSheet.Cells[Row, 2] as Excel.Range).Value = Work.Name;
                        Styles(StudentSheet.Cells[Row, 2], 12, XlHAlign.xlHAlignLeft, true);

                        (StudentSheet.Cells[Row, 3] as Excel.Range).Value = Work.IdType == 1 ? "Практическая" : Work.IdType == 2 ? "Теоретическая" : "Другое";
                        Styles(StudentSheet.Cells[Row, 3], 12, XlHAlign.xlHAlignCenter, true);

                        string Status = (Evaluation != null && Evaluation.Value.Trim() != "" && Evaluation.Value.Trim() != "2") ? "Сдано" : "Не сдано";
                        (StudentSheet.Cells[Row, 4] as Excel.Range).Value = Status;
                        Styles(StudentSheet.Cells[Row, 4], 12, XlHAlign.xlHAlignCenter, true);

                        Row++;
                    }
                }
            }
        }
        public static void Styles(Excel.Range Cell,
                                  int Fontsize,
                                  Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
                                  bool Border = false)
        {
            Cell.Font.Name = "Bahnschrift Light Condensed";
            Cell.Font.Size = Fontsize;
            Cell.HorizontalAlignment = Position;
            Cell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (Border)
            {
                Excel.Borders border = Cell.Borders;
                border.LineStyle = Excel.XlLineStyle.xlDouble;
                border.Weight = XlBorderWeight.xlThin;
            }
        }
    }
}
