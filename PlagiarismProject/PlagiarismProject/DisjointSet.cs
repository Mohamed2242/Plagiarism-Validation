using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlagiarismProject
{
	internal class DisjointSet
	{

		private Dictionary<long, long> parent = new Dictionary<long, long>();
		private Dictionary<long, int> rank = new Dictionary<long, int>();

		// Creates a new set for a node if it doesn't already exist.
		// O(1)
		public void MakeSet(long node)
		{
			if (!parent.ContainsKey(node))
			{
				parent[node] = node;
				rank[node] = 0;
			}
		}

		// Joins two sets if they are not already part of the same set
		// Complexity is O(log(D)) Where D is number of nodes
		public void Union(long set1, long set2)
		{
			//  Find the root of each set (root1, root2)
			long root1 = Find(set1);
			long root2 = Find(set2);

			// Already in the same set
			if (root1 == root2) return;

			// If the sets are different, it joins them, ensuring the set with lower rank points to the one with higher rank
			if (rank[root1] > rank[root2])
				parent[root2] = root1;
			
			else if (rank[root1] < rank[root2])
				parent[root1] = root2;
			// If they have the same rank, chooses (any) root1 to be the parent and increments its rank
			else if (rank[root1] == rank[root2])
			{
				parent[root2] = root1;
				rank[root1]++;
			}
		}

		// To Determines which set a particular node belongs to
		// Complexity is O(log(D)) Where D is number of nodes
		public long Find(long node)
		{
			if (parent[node] != node)
				parent[node] = Find(parent[node]);
			
			return parent[node];
		}

	}
}
