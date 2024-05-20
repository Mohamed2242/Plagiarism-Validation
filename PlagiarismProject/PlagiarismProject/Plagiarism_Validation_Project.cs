using System.Diagnostics;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Spire.Xls;

namespace PlagiarismProject
{
	internal class Plagiarism_Validation_Project
	{

		public static Dictionary<long, List<KeyValuePair<long, double>>> AdjListFor_Groups_Statistic_File;
		public static HashSet<long> visitedNodesInAdjList;
		public static Dictionary<long, List<long>> groupingCommunities;
		public static Dictionary<long, GroupInfo> filesGroups;
		public static Dictionary<long, double> avgScore;
		public static List<KeyValuePair<long, GroupInfo>> sortedFilesGroups;
		public static long numOfEdges;
		public static double totalScoreOfCurrentCommunity;
		public static long numOfComponents;

		public static Dictionary<long, List<Edge>> AdjListFor_Refined_Matches_File;
		public static Dictionary<long, List<Edge>> groupingEdges;
		public static Dictionary<long, List<Edge>> mstEdges;

		public static void Plagiarism_Validation()
		{
			Stopwatch SW = new Stopwatch();

			AdjListFor_Groups_Statistic_File = new Dictionary<long, List<KeyValuePair<long, double>>>();
			visitedNodesInAdjList = new HashSet<long>();
			groupingCommunities = new Dictionary<long, List<long>>();
			filesGroups = new Dictionary<long, GroupInfo>();
			avgScore = new Dictionary<long, double>();
			sortedFilesGroups = new List<KeyValuePair<long, GroupInfo>>();

			AdjListFor_Refined_Matches_File = new Dictionary<long, List<Edge>>();
			groupingEdges = new Dictionary<long, List<Edge>>();
			mstEdges = new Dictionary<long, List<Edge>>();

			string inputFilePath, StatFile_OutputPath, MST_File_OutputPath;
			while (true)
			{
				Console.WriteLine("Press 1 for Sample test cases: ");
				Console.WriteLine("Press 2 for Complete test cases: ");
				Console.WriteLine("Press 3 for Close ");
				string SampleOrCompleteFile = "", typeOfCompleteIfItIs = "";
				
				sbyte numOfFiles = 0;
				int SampleOrComp = int.Parse(Console.ReadLine());
				if (SampleOrComp == 1)
				{
					numOfFiles = 6;
					SampleOrCompleteFile = "Sample"; // O(1)
					typeOfCompleteIfItIs = @""; // O(1)

				}
				else if (SampleOrComp == 2)
				{
					numOfFiles = 2;
					SampleOrCompleteFile = "Complete"; // O(1)

					Console.WriteLine("Press 1 for Easy test cases: ");
					Console.WriteLine("Press 2 for Medium test cases: ");
					Console.WriteLine("Press 3 for Hard test cases: ");

					int typeOfComp = int.Parse(Console.ReadLine());
					if (typeOfComp == 1)
						typeOfCompleteIfItIs = @"\Easy"; // O(1)
					else if (typeOfComp == 2)
						typeOfCompleteIfItIs = @"\Medium"; // O(1)
					else if (typeOfComp == 3)
						typeOfCompleteIfItIs = @"\Hard"; // O(1)
				}
				else if (SampleOrComp == 3)
				{
					break;
				}
				for (sbyte i = 1; i <= numOfFiles; i++)
				{
					AdjListFor_Groups_Statistic_File.Clear();
					visitedNodesInAdjList.Clear();
					groupingCommunities.Clear();
					filesGroups.Clear();
					avgScore.Clear();
					sortedFilesGroups.Clear();
					numOfEdges = 0;
					totalScoreOfCurrentCommunity = 0;
					numOfComponents = 0;

					AdjListFor_Refined_Matches_File.Clear();
					groupingEdges.Clear();
					mstEdges.Clear();

					sbyte fileNum = i; // O(1)

					// O(1)
					inputFilePath = $@"E:\Year 3\Algo\Project\PlagiarismProject\Test Cases\{SampleOrCompleteFile}{typeOfCompleteIfItIs}\{fileNum}-Input.xlsx";
					StatFile_OutputPath = $@"E:\Year 3\Algo\Project\PlagiarismProject\Answers\{SampleOrCompleteFile}{typeOfCompleteIfItIs}\{fileNum}-StatFileAnswer.xlsx";
					MST_File_OutputPath = $@"E:\Year 3\Algo\Project\PlagiarismProject\Answers\{SampleOrCompleteFile}{typeOfCompleteIfItIs}\{fileNum}-mst_fileAnswer.xlsx";

					ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

					// Exact(NK) Whwere N is number of (pairs) rows and K is the avg length of the string
					Read_Input_File(inputFilePath);

					SW.Restart(); // O(1)

					// Complexity is O(D^2 + D*N)   D is number of nodes and N is number of pairs (edges)
					GenerateStat();

					// Complexity is O(G log(G) + Summation on the groups (Mc log(Mc))) Where G is number of groups Mc is the number of files in each group
					Create_Stat_File(StatFile_OutputPath, fileNum, typeOfCompleteIfItIs);

					SW.Stop(); // O(1)
					Console.WriteLine("Time Taken in Stat File in ms: " + SW.ElapsedMilliseconds); // O(1)

					SW.Restart(); // O(1)

					 // Complexity is O(Summation on the groups (Nc log(Nc))) Where Nc is the number of pairs (edges) in each group
					GenerateMST();

					// Complexity is Theta of(Summation on the groups (Nc))   Where Nc is the number of pairs in each group
					Create_MST_File(MST_File_OutputPath);
					SW.Stop(); // O(1)
					Console.WriteLine("Time Taken in MST File in ms: " + SW.ElapsedMilliseconds);
				}
			}
		}

		// Exact(NK) Whwere N is number of (pairs) rows and K is the avg length of the string
		private static void Read_Input_File(string inputFilePath)
		{

			// Load the existing Excel file
			using (ExcelPackage package = new ExcelPackage(new FileInfo(inputFilePath)))
			{
				// Assuming We're reading from the first sheet
				ExcelWorksheet sheet = package.Workbook.Worksheets[0];

				// Start from row 2 (headers are in 1st row)
				int numberOfRows = sheet.Dimension.Rows;
				// Exact(NK) Whwere N is number of (pairs) rows and K is the avg length of the string
				for (int row = 2; row <= numberOfRows; row++) 
				{

					string file1Path = sheet.Cells[row, 1].Value?.ToString();
					string file2Path = sheet.Cells[row, 2].Value?.ToString();

					string file1Link, file2Link;
					if (sheet.Cells[row, 1].Hyperlink != null || sheet.Cells[row, 2].Hyperlink != null)
					{
						file1Link = sheet.Cells[row, 1].Hyperlink.OriginalString;
						file2Link = sheet.Cells[row, 2].Hyperlink.OriginalString;
					}
					else
					{
						file1Link = file1Path; 
						file2Link = file2Path; 
					}
					string lineMatchesTxtValue = sheet.Cells[row, 3].Text;

					long file1Num = 0, file2Num = 0;
					sbyte file1Similarity = 0, file2Similarity = 0, similarity = 0;
					double lineMatches = 0; 
					if (file1Path != null && file2Path != null && lineMatchesTxtValue != null) 
					{
						// Extract the file number from the path of the file
						file1Num = long.Parse(Regex.Match(file1Path, @"\d+").Value);
						file2Num = long.Parse(Regex.Match(file2Path, @"\d+").Value);
						// Extract the numeric part (percentage) from the similarity string
						file1Similarity = sbyte.Parse(Regex.Match(file1Path, @"\((\d+(\.\d+)?)%\)").Groups[1].Value); 
						file2Similarity = sbyte.Parse(Regex.Match(file2Path, @"\((\d+(\.\d+)?)%\)").Groups[1].Value); 

						lineMatches = double.Parse(lineMatchesTxtValue);

						// For the Groups Statistic File
						if (!AdjListFor_Groups_Statistic_File.ContainsKey(file1Num)) 
							AdjListFor_Groups_Statistic_File[file1Num] = new List<KeyValuePair<long, double>>(); 
						AdjListFor_Groups_Statistic_File[file1Num].Add(new KeyValuePair<long, double>(file2Num, file1Similarity)); 
						if (!AdjListFor_Groups_Statistic_File.ContainsKey(file2Num))
							AdjListFor_Groups_Statistic_File[file2Num] = new List<KeyValuePair<long, double>>();
						AdjListFor_Groups_Statistic_File[file2Num].Add(new KeyValuePair<long, double>(file1Num, file2Similarity)); 


						// For the Refined Matches File
						similarity = file1Similarity > file2Similarity ? file1Similarity : file2Similarity;

						if (!AdjListFor_Refined_Matches_File.ContainsKey(file1Num))
							AdjListFor_Refined_Matches_File[file1Num] = new List<Edge>(); 

						if (!AdjListFor_Refined_Matches_File.ContainsKey(file2Num))
							AdjListFor_Refined_Matches_File[file2Num] = new List<Edge>();


						Edge edge = new Edge
						{
							File1Path = file1Path,
							File2Path = file2Path,
							File1Link = file1Link,
							File2Link = file2Link,
							File1Num = file1Num,
							File2Num = file2Num,
							Similarity = similarity,
							LineMatches = lineMatches
						};

						AdjListFor_Refined_Matches_File[file1Num].Add(edge);
						AdjListFor_Refined_Matches_File[file2Num].Add(edge);
					}
				}
			}
		}



		// Complexity is O(D^2 + D*N)   D is number of nodes and N is number of pairs (edges)
		private static void GenerateStat()
		{
			int groupId = 0;
			foreach (var vertex in AdjListFor_Groups_Statistic_File.Keys)
			{
				numOfComponents = 0;
				//this will be executed with number of groups
				if (!visitedNodesInAdjList.Contains(vertex))
				{
					// Initialize list of long for this groupID if it’s not exist
					if (!groupingCommunities.ContainsKey(groupId)) 
						groupingCommunities[groupId] = new List<long>();
					// Initialize list of edges for this groupID if it’s not exist
					if (!groupingEdges.ContainsKey(groupId))
						groupingEdges[groupId] = new List<Edge>();

					numOfEdges = 0;
					totalScoreOfCurrentCommunity = 0;

					// O(D + N) Where D is number of nodes and N is number of pairs (edges)
					DFS(vertex, groupId);

					// Calculate the avgScore of the current group
					avgScore[groupId] = Math.Round((totalScoreOfCurrentCommunity / numOfEdges), 1);
					// O(1)
					filesGroups[groupId] = new GroupInfo
					{
						IDs = groupingCommunities[groupId],
						AvgScore = avgScore[groupId],
						NumOfComponents = ++numOfComponents, 
						ListOfEdges = groupingEdges[groupId] 
					};
					groupId++;
				}
			}
		}

		// O(D + N) Where D is number of nodes and N is number of pairs (edges)
		private static void DFS(long startNode, long GID)
		{
			visitedNodesInAdjList.Add(startNode);
			// Add this node to it’s group
			groupingCommunities[GID].Add(startNode);
			
			if (AdjListFor_Groups_Statistic_File.ContainsKey(startNode))
			{
				foreach (var neighbor in AdjListFor_Groups_Statistic_File[startNode])
				{
					long neighborNode = neighbor.Key;
					totalScoreOfCurrentCommunity += neighbor.Value;
					numOfEdges++;
					if (!visitedNodesInAdjList.Contains(neighborNode))
					{
						// Add the list of edges of this node to it’s group, we need this in the MST part
						groupingEdges[GID].AddRange(AdjListFor_Refined_Matches_File[startNode]);
						numOfComponents++;
						DFS(neighborNode, GID);
					}
				}
			}
			return;
		}

		// Complexity is O(G log(G) + Summation on the groups (Mc log(Mc))) Where G is number of groups Mc is the number of files in each group
		private static void Create_Stat_File(string statFile_OutputPath, sbyte fileNum, string typeOfCompleteIfItIs)
		{
			// O(G log(G))  where G is number of groups
			sortedFilesGroups = filesGroups.OrderByDescending(kv => kv.Value.AvgScore).ToList();

			// Create a new workbook
			using (ExcelPackage package = new ExcelPackage())
			{
				// Assuming We're reading from the first sheet
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Stat"); 

				// Write headers
				worksheet.Cells[1, 1].Value = "Component Index"; 
				worksheet.Cells[1, 2].Value = "Vertices"; 
				worksheet.Cells[1, 3].Value = "Average Similarity"; 
				worksheet.Cells[1, 4].Value = "Component Count"; 
				int j = 0;

				int max1 = 16;
				int max2 = 9;
				int max3 = 18;
				int max4 = 16;
				// Write data
				// Complexity is Theta of(Summation on the groups (Mc Log(Mc))) Where Mc is the number of files in each group
				foreach (var kvp in sortedFilesGroups)
				{

					GroupInfo groupInfo = kvp.Value;

					// Component index
					worksheet.Cells[j + 2, 1].Value = (j + 1);

					// Vertices
					// O(M log(M)) where M is number of items in the IDs list
					groupInfo.IDs.Sort();
					worksheet.Cells[j + 2, 2].Value = string.Join(", ", groupInfo.IDs);
					if (groupInfo.IDs.Count() > max2)
						max2 = groupInfo.IDs.Count();

					// Average similarity
					worksheet.Cells[j + 2, 3].Value = groupInfo.AvgScore;

					// Component Count
					worksheet.Cells[j + 2, 4].Value = groupInfo.NumOfComponents;
					j++;
				}

				worksheet.Column(1).Width = max1;
				if (fileNum == 1 && typeOfCompleteIfItIs == @"")
					worksheet.Column(2).Width = max2 * 11;
				if (fileNum == 3 && typeOfCompleteIfItIs == @"")
					worksheet.Column(2).Width = max2 * 3.5;
				if ((fileNum == 2 || fileNum == 5 || fileNum == 6) && typeOfCompleteIfItIs == @"")
					worksheet.Column(2).Width = max2 * 1.7;
				if (typeOfCompleteIfItIs == @"\Easy" || typeOfCompleteIfItIs == @"\Medium")
					worksheet.Column(2).Width = max2 * 4;
				if (typeOfCompleteIfItIs == @"\Hard")
					worksheet.Column(2).Width = max2 * 2.6;
				worksheet.Column(3).Width = max3;
				worksheet.Column(4).Width = max4;

				// Save the workbook to the specified location
				package.SaveAs(new FileInfo(statFile_OutputPath)); // O(1)
			}
		}

		// Complexity is O(Summation on the groups (Nc log(Nc))) Where Nc is the number of pairs (edges) in each group
		private static Dictionary<long, List<Edge>> GenerateMST()
		{
			// To detect cycles
			DisjointSet disjointSet = new DisjointSet();

			//Dictionary<long, List<Edge>> mstEdges = new Dictionary<long, List<Edge>>();
			Dictionary<long, List<Edge>> sortedEdgesInGroups = new Dictionary<long, List<Edge>>();

			// Sort sortedGroups by Similarity in descending order to build mst with edges that have the largest similarities
			// Complexity is O(Summation on the groups(Nc log(Nc))) Where Nc is the number of pairs(edges) in each group
			foreach (var kv in sortedFilesGroups)
				sortedEdgesInGroups[kv.Key] = kv.Value.ListOfEdges.OrderByDescending(e => e.Similarity).ThenByDescending(ee=>ee.LineMatches).ToList();

			// Initialize disjoint sets for all nodes (ensures that every node starts in its own set.)
			// O(D) Where D is number of nodes
			foreach (var node in AdjListFor_Refined_Matches_File.Keys)
				disjointSet.MakeSet(node);

			// Build the Maximum Spanning Tree (MST) by Kruskal's algorithm
			// O(N log(D))  Where N is number of pairs (edges) and D is number of nodes
			foreach (var groupID in sortedEdgesInGroups)
			{
				foreach (var edge in groupID.Value)
				{
					long groupEdgesID = groupID.Key;
					// Initialize list of edges for this groupID if it’s not exist
					if (!mstEdges.ContainsKey(groupEdgesID))
						mstEdges[groupEdgesID] = new List<Edge>();
					// Checks if adding the edge would create a cycle
					// by comparing the sets containing the source and destination nodes (File1Num, File2Num)
					if (disjointSet.Find(edge.File1Num) != disjointSet.Find(edge.File2Num))
					{
						// Join the sets containing the source and destination nodes (File1Num, File2Num)
						disjointSet.Union(edge.File1Num, edge.File2Num);
						// Add this edge to it’s group in mstEdges
						mstEdges[groupEdgesID].Add(edge);
					}
					// If the sets are the same, it would create a cycle, so the edge is not added
					// So mstEdges is a dictionary that contain groups, on each group the edges that have the lareest similarities and doesn’t create a cycle
				}
			}

			// Sort the mstEdges dictionary to sort the edges in each group with lineMatches
			// Complexity is O(Summation on the groups(Nc log(Nc))) Where Nc is the number of pairs(edges) in each group
			foreach (var kv in mstEdges)
				mstEdges[kv.Key] = kv.Value.OrderByDescending(e => e.LineMatches).ToList();

			return mstEdges;
		}

		// Complexity is Theta of(Summation on the groups (Nc)) Where Nc is the number of pairs
		private static void Create_MST_File(string MST_OutputPath)
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				// Assuming We're reading from the first sheet
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MST");

				// Write header
				worksheet.Cells	[1, 1].Value = "File 1";
				worksheet.Cells[1, 2].Value = "File 2";
				worksheet.Cells[1, 3].Value = "Line Matches";
				int jj = 0;

				// Write edge data
				// Complexity is Theta of(Summation on the groups (Nc))   Where Nc is the number of pairs 
				foreach (var groupedEdges in mstEdges)
				{
					foreach (var edge in groupedEdges.Value)
					{

						worksheet.Cells[jj + 2, 1].Hyperlink = new Uri(edge.File1Link);
						worksheet.Cells[jj + 2, 1].Value = edge.File1Path;
						worksheet.Cells[jj + 2, 1].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

						worksheet.Cells[jj + 2, 2].Hyperlink = new Uri(edge.File2Link);
						worksheet.Cells[jj + 2, 2].Value = edge.File2Path;
						worksheet.Cells[jj + 2, 2].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

						// Set the Line Matches value
						worksheet.Cells[jj + 2, 3].Value = edge.LineMatches;
						jj++;
					}
				}
				for (int col = 1; col <= worksheet.Dimension.Columns; col++)
					worksheet.Column(col).AutoFit();

				// Save the workbook to the specified location
				package.SaveAs(new FileInfo(MST_OutputPath));
			}
		}
	}
}
