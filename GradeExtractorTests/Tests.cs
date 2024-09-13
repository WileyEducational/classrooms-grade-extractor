using NUnit.Framework;
using Moq;
using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace GradeExtractor.Tests
{
    [TestFixture]
    public class ProgramTests
    {
        private Mock<TextReader> _mockConsoleInput;
        private StringWriter _consoleOutput;
        private Program _program;

        [SetUp]
        public void Setup()
        {
            _mockConsoleInput = new Mock<TextReader>();
            _consoleOutput = new StringWriter();
            Console.SetIn(_mockConsoleInput.Object);
            Console.SetOut(_consoleOutput);
            _program = new Program();
        }

        [TearDown]
        public void TearDown()
        {
            _consoleOutput.Dispose();
        }

        [Test]
        public void AddAssignmentFromFile_ShouldAddAssignments()
        {
            // Arrange
            var assignments = new Dictionary<string, List<Assignment>>();
            var csvContent = "AssignmentName, , ,GitHubUsername, , , , ,PointsAwarded\n" +
                             "\"Assignment1\", , ,\"user1\", , , , ,\"100\"\n" +
                             "\"Assignment2\", , ,\"user2\", , , , ,\"90\"";
            var filePath = Path.GetTempFileName();
            File.WriteAllText(filePath, csvContent);
            _mockConsoleInput.SetupSequence(x => x.ReadLine())
                             .Returns(filePath);

            // Act
            _program.AddAssignmentFromFile(assignments);

            // Assert
            Assert.That(assignments.Count, Is.EqualTo(2));
            Assert.That(assignments["user1"].Count, Is.EqualTo(1));
            Assert.That(assignments["user2"].Count, Is.EqualTo(1));
            Assert.That(assignments["user1"][0].AssignmentName, Is.EqualTo("Assignment1"));
            Assert.That(assignments["user1"][0].PointsAwarded, Is.EqualTo(100));
            Assert.That(assignments["user2"][0].AssignmentName, Is.EqualTo("Assignment2"));
            Assert.That(assignments["user2"][0].PointsAwarded, Is.EqualTo(90));

            // Clean up
            File.Delete(filePath);
        }

        [Test]
        public void CreateExcelSheet_ShouldCreateExcelFile()
        {
            // Arrange
            var assignments = new Dictionary<string, List<Assignment>>
            {
                { "user1", new List<Assignment> { new Assignment("Assignment1", 100) } },
                { "user2", new List<Assignment> { new Assignment("Assignment2", 90) } },
                { "user3", new List<Assignment> { new Assignment("Assignment1", 100), new Assignment("Assignment2", 100) } }
            };
            var excelPath = Path.ChangeExtension(Path.GetTempFileName(), "xlsx");
            _mockConsoleInput.SetupSequence(x => x.ReadLine())
                             .Returns(excelPath);

            // Act
            _program.CreateExcelSheet(assignments);

            // Assert
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet("Assignments");
                Assert.That(worksheet.Cell(1, 1).Value, Is.EqualTo("GitHub Username"));
                Assert.That(worksheet.Cell(1, 2).Value, Is.EqualTo("Assignment1"));
                Assert.That(worksheet.Cell(1, 3).Value, Is.EqualTo("Assignment2"));
                Assert.That(worksheet.Cell(1, 4).Value, Is.EqualTo("Pass/Fail"));

                Assert.That(worksheet.Cell(2, 1).Value, Is.EqualTo("user1"));
                Assert.That(worksheet.Cell(2, 2).Value, Is.EqualTo(100));
                Assert.That(worksheet.Cell(2, 3).Value, Is.EqualTo("N/A"));
                Assert.That(worksheet.Cell(2, 4).Value, Is.EqualTo("Fail"));

                Assert.That(worksheet.Cell(3, 1).Value, Is.EqualTo("user2"));
                Assert.That(worksheet.Cell(3, 2).Value, Is.EqualTo("N/A"));
                Assert.That(worksheet.Cell(3, 3).Value, Is.EqualTo(90));
                Assert.That(worksheet.Cell(3, 4).Value, Is.EqualTo("Fail"));

                Assert.That(worksheet.Cell(4, 1).Value, Is.EqualTo("user3"));
                Assert.That(worksheet.Cell(4, 2).Value, Is.EqualTo(100));
                Assert.That(worksheet.Cell(4, 3).Value, Is.EqualTo(100));
                Assert.That(worksheet.Cell(4, 4).Value, Is.EqualTo("Pass"));
            }

            // Clean up
            File.Delete(excelPath);
        }
    }
}
