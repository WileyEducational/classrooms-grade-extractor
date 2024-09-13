using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GradeExtractor
{
    public class Assignment
    {
        public string AssignmentName { get; set; }
        public int PointsAwarded { get; set; }

        public Assignment(string assignmentName, int pointsAwarded)
        {
            AssignmentName = assignmentName;
            PointsAwarded = pointsAwarded;
        }
    }
}
