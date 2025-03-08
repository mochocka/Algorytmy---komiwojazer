using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {

        var tabuLengths = new List<int> { 10, 50, 100, 500 };
        var criteriaLimits = new List<int> { 100, 500, 1000, 5000 };
        var neighborhoodTypes = new List<string> { "swap", "insert", "inverse" };
        var criteriaTypes = new List<string> { "liczba iteracji bez poprawy", "liczba iteracji" };


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


        var resultsFileName = "Results.xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            foreach (var file in dataFiles)
            {
                string sheetName = Path.GetFileNameWithoutExtension(file.Key);
                var worksheet = package.Workbook.Worksheets.Add(sheetName);
                worksheet.Cells[1, 1].Value = "Długość listy tabu";
                worksheet.Cells[1, 2].Value = "Limit kryterium";
                worksheet.Cells[1, 3].Value = "Rodzaj sąsiedztwa";
                worksheet.Cells[1, 4].Value = "Rodzaj kryterium";
                worksheet.Cells[1, 5].Value = "Najkrótszy dystans";
                worksheet.Cells[1, 6].Value = "Czas (milisekundy)";
                worksheet.Cells[1, 7].Value = "Najlepsza ścieżka";

                int row = 2;


                foreach (var tabuLength in tabuLengths)
                {
                    foreach (var criteriaLimit in criteriaLimits)
                    {
                        foreach (var neighborhoodType in neighborhoodTypes)
                        {
                            foreach (var criteriaType in criteriaTypes)
                            {
                                for (int i = 0; i < 5 ; i++)
                                {

                                    var data = file.Value;
                                    var initialCombination = GenerateRandomCombination(data.GetLength(0));


                                    var result = TabuSearch(data, initialCombination, criteriaType, criteriaLimit, tabuLength, neighborhoodType);


                                    worksheet.Cells[row, 1].Value = tabuLength;
                                    worksheet.Cells[row, 2].Value = criteriaLimit;
                                    worksheet.Cells[row, 3].Value = neighborhoodType;
                                    worksheet.Cells[row, 4].Value = criteriaType;
                                    worksheet.Cells[row, 5].Value = result.BestDistance;
                                    worksheet.Cells[row, 6].Value = result.TimeInMilliseconds;
                                    worksheet.Cells[row, 7].Value = string.Join(", ", result.BestCombination);
                                    row++;
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

    static (double BestDistance, List<int> BestCombination, double TimeInMilliseconds) TabuSearch(
        double[,] cities,
        List<int> initialCombination,
        string criterion,
        int criterionLimit,
        int tabuLength,
        string neighborhoodType)
    {
        var startTime = DateTime.Now;
        int n = cities.GetLength(0);
        var tabuList = new int[n + 1, n + 1]; 
        var bestDistance = CalculateDistance(cities, initialCombination);
        var bestCombination = new List<int>(initialCombination);
        var currentCombination = new List<int>(bestCombination);
        int iterationsWithoutImprovement = 0;

        while (true)
        {
            double bestOption = double.MaxValue;
            List<int> bestOptionCombination = null;
            int[] move = null;

            // Sąsiedztwo
            for (int i = 0; i < n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if (i != j)
                    {
                        var newCombination = GenerateNeighbor(currentCombination, i, j, neighborhoodType);
                        double newDistance = CalculateDistance(cities, newCombination);

                        if ((tabuList[i, j] == 0 && newDistance <= bestOption) ||
                            (tabuList[i, j] > 0 && newDistance < bestDistance))
                        {
                            bestOption = newDistance;
                            bestOptionCombination = newCombination;
                            move = new[] { i, j };
                        }
                    }
                }
            }

            if (bestOptionCombination == null)
                break;

            currentCombination = bestOptionCombination;
            if (bestDistance > bestOption)
            {
                bestDistance = bestOption;
                bestCombination = new List<int>(currentCombination);
                iterationsWithoutImprovement = 0;
            }
            else
            {
                iterationsWithoutImprovement++;
            }

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    if (tabuList[i, j] > 0)
                        tabuList[i, j]--;
                }
            }

            if (move != null)
            {
                tabuList[move[0], move[1]] = tabuLength;
                tabuList[move[1], move[0]] = tabuLength;
            }

            if (criterion == "liczba iteracji bez poprawy" && iterationsWithoutImprovement >= criterionLimit)
                break;

            if (criterion == "liczba iteracji" && iterationsWithoutImprovement >= criterionLimit)
                break;
        }

        double elapsedTime = (DateTime.Now - startTime).TotalMilliseconds;
        return (bestDistance, bestCombination, elapsedTime);
    }

    static List<int> GenerateNeighbor(List<int> combination, int i, int j, string type)
    {
        var neighbor = new List<int>(combination);
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
