using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {

        var numBees = new List<int> { 500 };
        var numEliteSites = new List<int> { 2, 5, 15, 25 };
        var numBestSites = new List<int> { 4, 10, 30, 50 };
        var numEliteBees = new List<int> { 4, 10, 30, 50 };
        var numBestBees = new List<int> { 2, 5, 15, 25 };
        var maxIterations = new List<int> { 1000 };
        var criteriaTypes = new List<string> { "liczba iteracji" };
        var neighborhoodTypes = new List<string> { "swap", "insert", "inverse" };


        var dataFiles = new Dictionary<string, double[,]>();
        string[] fileNames = { "Dane_TSP_48.xlsx", "Dane_TSP_76.xlsx", "Dane_TSP_127.xlsx" };

        foreach (var fileName in fileNames)
        {
            var data = ReadExcel(fileName);
            if (data != null)
            {
                dataFiles[fileName] = data;
                Console.WriteLine($"Wczytano dane z pliku: {fileName}");
            }
            else
            {
                Console.WriteLine($"Błąd wczytywania danych z pliku: {fileName}");
            }
        }


        int totalCases = numBees.Count * numEliteSites.Count * numBestSites.Count *
                         numEliteBees.Count * numBestBees.Count * maxIterations.Count *
                         criteriaTypes.Count * neighborhoodTypes.Count * fileNames.Length * 3;
        int completedCases = 0;


        var resultsFileName = "BeeAlgorithmResults.xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            foreach (var file in dataFiles)
            {
                string sheetName = Path.GetFileNameWithoutExtension(file.Key);
                var worksheet = package.Workbook.Worksheets.Add(sheetName);
                worksheet.Cells[1, 1].Value = "Liczba pszczół zwiadowczych";
                worksheet.Cells[1, 2].Value = "Liczba elitarnych stanowisk";
                worksheet.Cells[1, 3].Value = "Liczba najlepszych stanowisk";
                worksheet.Cells[1, 4].Value = "Liczba pszczół elitarnych";
                worksheet.Cells[1, 5].Value = "Liczba pszczół najlepszych";
                worksheet.Cells[1, 6].Value = "Maksymalna liczba iteracji";
                worksheet.Cells[1, 7].Value = "Rodzaj kryterium";
                worksheet.Cells[1, 8].Value = "Rodzaj sąsiedztwa";
                worksheet.Cells[1, 9].Value = "Najkrótszy dystans";
                worksheet.Cells[1, 10].Value = "Czas (milisekundy)";
                worksheet.Cells[1, 11].Value = "Najlepsza ścieżka";

                int row = 2;

                foreach (var bees in numBees)
                {
                    foreach (var eliteSites in numEliteSites)
                    {
                        foreach (var bestSites in numBestSites)
                        {
                            foreach (var eliteBees in numEliteBees)
                            {
                                foreach (var bestBees in numBestBees)
                                {
                                    foreach (var iterations in maxIterations)
                                    {
                                        foreach (var criteria in criteriaTypes)
                                        {
                                            foreach (var neighborhood in neighborhoodTypes)
                                            {
                                                for (int i = 0; i < 2; i++) 
                                                {
                                                    var data = file.Value;
                                                    var initialCombination = GenerateRandomCombination(data.GetLength(0));

                                                    var result = BeesAlgorithm(
                                                        data,
                                                        bees,
                                                        eliteSites,
                                                        bestSites,
                                                        eliteBees,
                                                        bestBees,
                                                        iterations,
                                                        criteria,
                                                        neighborhood
                                                    );

                                                    worksheet.Cells[row, 1].Value = bees;
                                                    worksheet.Cells[row, 2].Value = eliteSites;
                                                    worksheet.Cells[row, 3].Value = bestSites;
                                                    worksheet.Cells[row, 4].Value = eliteBees;
                                                    worksheet.Cells[row, 5].Value = bestBees;
                                                    worksheet.Cells[row, 6].Value = iterations;
                                                    worksheet.Cells[row, 7].Value = criteria;
                                                    worksheet.Cells[row, 8].Value = neighborhood;
                                                    worksheet.Cells[row, 9].Value = result.BestDistance;
                                                    worksheet.Cells[row, 10].Value = result.TimeInMilliseconds;
                                                    worksheet.Cells[row, 11].Value = string.Join(", ", result.BestCombination);

                                                    row++;


                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            FileInfo fileInfo = new FileInfo(resultsFileName);
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }
            package.SaveAs(fileInfo);
        }

        Console.WriteLine($"Wyniki zapisane w pliku {resultsFileName}");
    }


    static double[,] ReadExcel(string fileName)
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                double[,] distanceMatrix = new double[rowCount - 1, colCount - 1];

                for (int row = 2; row <= rowCount; row++)
                {
                    for (int col = 2; col <= colCount; col++)
                    {
                        if (double.TryParse(worksheet.Cells[row, col].Text, out double value))
                        {
                            distanceMatrix[row - 2, col - 2] = value;
                        }
                        else
                        {
                            Console.WriteLine($"Nieprawidłowa wartość w komórce ({row}, {col}).");
                            return null;
                        }
                    }
                }
                return distanceMatrix;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Błąd podczas odczytu pliku: " + ex.Message);
            return null;
        }
    }

    static List<int> GenerateRandomCombination(int cityCount)
    {
        var combination = new List<int>();
        var random = new Random();
        var availableCities = new HashSet<int>();

        for (int i = 1; i <= cityCount; i++)
            availableCities.Add(i);

        while (availableCities.Count > 0)
        {
            var city = random.Next(1, cityCount + 1);
            if (availableCities.Remove(city))
                combination.Add(city);
        }

        return combination;
    }

    static (double BestDistance, List<int> BestCombination, double TimeInMilliseconds) BeesAlgorithm(
        double[,] cities,
        int numBees,
        int numEliteSites,
        int numBestSites,
        int numEliteBees,
        int numBestBees,
        int maxIterations,
        string criteria,
        string neighborhoodType)
    {
        var startTime = DateTime.Now;
        int n = cities.GetLength(0);


        var scoutBees = new List<List<int>>();
        for (int i = 0; i < numBees; i++)
        {
            scoutBees.Add(GenerateRandomCombination(n));
        }

        double bestDistance = double.MaxValue;
        List<int> bestCombination = null;
        int iterationsWithoutImprovement = 0;

        for (int iteration = 0; iteration < maxIterations; iteration++)
        {
            var evaluatedSites = new List<(double Distance, List<int> Combination)>();

            foreach (var scout in scoutBees)
            {
                double distance = CalculateDistance(cities, scout);
                evaluatedSites.Add((distance, new List<int>(scout)));
            }

            evaluatedSites.Sort((a, b) => a.Distance.CompareTo(b.Distance));

            var eliteSites = evaluatedSites.GetRange(0, numEliteSites);
            var bestSites = evaluatedSites.GetRange(numEliteSites, Math.Min(numBestSites, evaluatedSites.Count - numEliteSites));

            scoutBees.Clear();


            foreach (var (distance, combination) in eliteSites)
            {
                for (int i = 0; i < numEliteBees; i++)
                {
                    scoutBees.Add(GenerateNeighbor(combination, neighborhoodType));
                }
            }


            foreach (var (distance, combination) in bestSites)
            {
                
                for (int i = 0; i < numBestBees; i++)
                {
                    scoutBees.Add(GenerateNeighbor(combination, neighborhoodType));
                }
            }


            while (scoutBees.Count < numBees)
            {
                scoutBees.Add(GenerateRandomCombination(n));
            }


            var currentBest = evaluatedSites[0];
            if (currentBest.Distance < bestDistance)
            {
                bestDistance = currentBest.Distance;
                bestCombination = currentBest.Combination;
                iterationsWithoutImprovement = 0;
            }
            else
            {
                iterationsWithoutImprovement++;
            }

            Console.WriteLine($"Iteration: {iteration}, Iterations without improvement: {iterationsWithoutImprovement}");

            if (criteria == "liczba iteracji bez poprawy" && iterationsWithoutImprovement >= maxIterations)
                break;

            if (criteria == "liczba iteracji" && iteration >= maxIterations - 1)
                break;
        }

        double elapsedTime = (DateTime.Now - startTime).TotalMilliseconds;
        return (bestDistance, bestCombination, elapsedTime);
    }

    static List<int> GenerateNeighbor(List<int> combination, string type)
    {
        var neighbor = new List<int>(combination);
        var random = new Random();

        int i = random.Next(neighbor.Count);
        int j = random.Next(neighbor.Count);

        if (type == "swap")
        {
            (neighbor[i], neighbor[j]) = (neighbor[j], neighbor[i]);
        }
        else if (type == "insert")
        {
            int city = neighbor[j];
            neighbor.RemoveAt(j);
            neighbor.Insert(i, city);
        }
        else if (type == "inverse")
        {
            if (i > j) (i, j) = (j, i);
            neighbor.Reverse(i, j - i + 1);
        }

        return neighbor;
    }

    static double CalculateDistance(double[,] cities, List<int> combination)
    {
        double distance = 0;
        for (int i = 0; i < combination.Count - 1; i++)
            distance += cities[combination[i] - 1, combination[i + 1] - 1];
        distance += cities[combination[^1] - 1, combination[0] - 1];
        return distance;
    }
}
