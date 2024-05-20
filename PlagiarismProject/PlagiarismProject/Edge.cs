using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlagiarismProject
{
	internal class Edge
	{
		public string File1Path { get; set; }
		public string File2Path { get; set; }
		public string File1Link { get; set; }
		public string File2Link { get; set; }
		public long File1Num { get; set; }
		public long File2Num { get; set; }
		public sbyte Similarity { get; set; }
		public double LineMatches { get; set; }
	}
}
