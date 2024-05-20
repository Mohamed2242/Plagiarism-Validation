using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlagiarismProject
{
	internal class GroupInfo
	{
		
		public List<long> IDs { get; set; }
		public double AvgScore { get; set; }
		public long NumOfComponents { get; set; }
        public List<Edge> ListOfEdges { get; set; }

    }
}
