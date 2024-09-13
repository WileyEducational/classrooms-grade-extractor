using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace GradeExtractor
{
    public class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.Run();
        }

        void Run()
        {
            var assignments = new Dictionary<string, List<Assignment>>();
            while (true)
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1. Add Assignment from File");
                Console.WriteLine("2. Create Excel Sheet (end with .xlsx)");
                Console.WriteLine("3. Exit");
                Console.Write("Select an option: ");
                var choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        AddAssignmentFromFile(assignments);
                        break;
                    case "2":
                        CreateExcelSheet(assignments);
                        break;
                    case "3":
                        Console.WriteLine("Exiting program.");
                        return;
                    default:
                        Console.WriteLine("Invalid option. Please try again.");
                        break;
                }
            }
        }

        public void AddAssignmentFromFile(Dictionary<string, List<Assignment>> assignments)
        {
            Console.Write("Enter the path to the CSV file: ");
            var filePath = Console.ReadLine();
            var newAssignments = ReadAssignmentsFromClassroomsCSV(filePath);
            foreach (var kvp in newAssignments)
            {
                if (!assignments.ContainsKey(kvp.Key))
                {
                    assignments[kvp.Key] = new List<Assignment>();
                }
                assignments[kvp.Key].AddRange(kvp.Value);
            }
            Console.WriteLine("Assignments added successfully.");
        }

        public void CreateExcelSheet(Dictionary<string, List<Assignment>> assignments)
        {
            Console.Write("Enter the path to save the Excel file: ");
            var excelPath = Console.ReadLine();
            CreateExcelFile(assignments, excelPath);
            Console.WriteLine("Excel file created successfully.");
        }

        public void CreateExcelFile(Dictionary<string, List<Assignment>> assignments, string filePath)
        {
            filePath = filePath.Trim('"');
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Assignments");

                // Add headers
                CreateColumns(worksheet, assignments);

                // Add data
                AddData(worksheet, assignments);

                workbook.SaveAs(filePath);
            }
        }

        void CreateColumns(IXLWorksheet worksheet, Dictionary<string, List<Assignment>> assignments)
        {
            worksheet.Cell(1, 1).Value = "GitHub Username";
            int column = 2;
            List<string> assignmentNames = CreateColumnsForAssignment(assignments);

            foreach (var assignmentName in assignmentNames)
            {
                worksheet.Cell(1, column++).Value = assignmentName;
            }

            worksheet.Cell(1, column).Value = "Pass/Fail";
        }

        void AddData(IXLWorksheet worksheet, Dictionary<string, List<Assignment>> assignments)
        {
            List<string> assignmentNames = CreateColumnsForAssignment(assignments);
            int row = 2;
            foreach (var kvp in assignments)
            {
                var githubUsername = kvp.Key;
                var userAssignments = kvp.Value;

                worksheet.Cell(row, 1).Value = githubUsername;

                var assignmentDict = new Dictionary<string, int>();
                foreach (var assignment in userAssignments)
                {
                    assignmentDict[assignment.AssignmentName] = assignment.PointsAwarded;
                }

                AddUserAssignments(worksheet, row, assignmentNames, assignmentDict);
                row++;
            }
        }

        void AddUserAssignments(IXLWorksheet worksheet, int row, List<string> assignmentNames, Dictionary<string, int> assignmentDict)
        {
            int column = 2;
            bool passed = true;
            foreach (var assignmentName in assignmentNames)
            {
                if (assignmentDict.TryGetValue(assignmentName, out int pointsAwarded))
                {
                    var cell = worksheet.Cell(row, column);
                    cell.Value = pointsAwarded;
                    if (pointsAwarded < 100)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCCCB");
                        cell.Style.Font.FontColor = XLColor.DarkRed;
                        passed = false;
                    }
                    else
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                        cell.Style.Font.FontColor = XLColor.DarkGreen;
                    }
                }
                else
                {
                    var cell = worksheet.Cell(row, column);
                    cell.Value = "N/A";
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCCCB");
                    cell.Style.Font.FontColor = XLColor.DarkRed;
                    passed = false;
                }
                column++;
            }

            var passFailCell = worksheet.Cell(row, column);
            passFailCell.Value = passed ? "Pass" : "Fail";
            passFailCell.Style.Fill.BackgroundColor = passed ? XLColor.LightGreen : XLColor.FromHtml("#FFCCCB");
            passFailCell.Style.Font.FontColor = passed ? XLColor.DarkGreen : XLColor.DarkRed;
        }

        List<string> CreateColumnsForAssignment(Dictionary<string, List<Assignment>> assignments)
        {
            List<string> assignmentNames = new List<string>();
            foreach (var userAssignments in assignments.Values)
            {
                foreach (var assignment in userAssignments)
                {
                    if (!assignmentNames.Contains(assignment.AssignmentName))
                    {
                        assignmentNames.Add(assignment.AssignmentName);
                    }
                }
            }
            return assignmentNames;
        }

        // The key of the dictionary is the GitHub username.
        public Dictionary<string, List<Assignment>> ReadAssignmentsFromClassroomsCSV(string filePath)
        {
            var assignments = new Dictionary<string, List<Assignment>>();

            filePath = filePath.Trim('"');

            using (var reader = new StreamReader(filePath))
            {
                // Skip the header line
                reader.ReadLine();

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    var assignmentName = values[0].Trim('"');
                    var githubUsername = values[3].Trim('"');
                    var pointsAwarded = int.Parse(values[8].Trim('"'));

                    if (!assignments.ContainsKey(githubUsername))
                    {
                        assignments[githubUsername] = new List<Assignment>();
                    }

                    assignments[githubUsername].Add(new Assignment(assignmentName, pointsAwarded));
                }
            }

            return assignments;
        }
    }
}

