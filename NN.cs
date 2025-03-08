using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml; 

class TravelingSalesman
{
    static void Main(string[] args)
    {
        string inputFileName = "Dane_TSP_127.xlsx"; 
        string outputFileName = "Wyniki_TSP.xlsx"; 

        if (!File.Exists(inputFileName))
        {
            Console.WriteLine($"Plik {inputFileName} nie istnieje w bieżącym katalogu.");
            return;
        }

        var distanceMatrix = ReadExcel(inputFileName);

        if (distanceMatrix == null)
        {
            Console.WriteLine("Nie udało się wczytać danych z pliku Excel.");
            return;
        }


        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        using (var package = new ExcelPackage())
        {
            var summarySheet = package.Workbook.Worksheets.Add("Podsumowanie");

            summarySheet.Cells[1, 1].Value = "Czas wykonania algorytmu (ms)";
            summarySheet.Cells[3, 1].Value = "Miasto startowe";
            summarySheet.Cells[3, 2].Value = "Suma odległości";
            summarySheet.Cells[3, 3].Value = "Trasa";

            int summaryRow = 4;

            for (int startCity = 1; startCity <= distanceMatrix.GetLength(0); startCity++)
            {
                Console.WriteLine($"Przetwarzanie dla miasta startowego: {startCity}");

                var result = NearestNeighborAlgorithm(distanceMatrix, startCity);

                summarySheet.Cells[summaryRow, 1].Value = startCity;
                summarySheet.Cells[summaryRow, 2].Value = result.Item2;
                summarySheet.Cells[summaryRow, 3].Value = string.Join(", ", result.Item1);
                summaryRow++;
            }

            stopwatch.Stop();
            double elapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds;


            summarySheet.Cells[2, 1].Value = elapsedMilliseconds;

            Console.WriteLine($"Czas wykonania algorytmu: {elapsedMilliseconds} ms.");

            try
            {
                package.SaveAs(new FileInfo(outputFileName));
                Console.WriteLine($"Wyniki zapisano do pliku {outputFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Błąd podczas zapisywania pliku: " + ex.Message);
            }
        }
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
            Console.WriteLine("Błąd podczas odczytu pliku Excel: " + ex.Message);
            return null;
        }
    }

    static Tuple<List<int>, double> NearestNeighborAlgorithm(double[,] distanceMatrix, int startCity)
    {
        int cityCount = distanceMatrix.GetLength(0);
        bool[] visited = new bool[cityCount];
        List<int> path = new List<int>();
        double totalDistance = 0;

        int currentCity = startCity;
        visited[currentCity - 1] = true;
        path.Add(currentCity);

        for (int step = 0; step < cityCount - 1; step++)
        {
            double shortestDistance = double.MaxValue;
            int nextCity = -1;

            for (int i = 0; i < cityCount; i++)
            {
                if (!visited[i] && distanceMatrix[currentCity - 1, i] < shortestDistance)
                {
                    shortestDistance = distanceMatrix[currentCity - 1, i];
                    nextCity = i + 1;
                }
            }

            if (nextCity == -1)
            {
                Console.WriteLine("Nie znaleziono następnego miasta. Algorytm zakończony przedwcześnie.");
                return Tuple.Create(path, totalDistance);
            }

            visited[nextCity - 1] = true;
            path.Add(nextCity);
            totalDistance += shortestDistance;
            currentCity = nextCity;
        }

        if (currentCity >= 1 && currentCity <= distanceMatrix.GetLength(0) &&
            path[0] >= 1 && path[0] <= distanceMatrix.GetLength(1))
        {
            totalDistance += distanceMatrix[currentCity - 1, path[0] - 1]; 
            path.Add(path[0]); 
        }
        else
        {
            Console.WriteLine("Nieprawidłowe indeksy w macierzy odległości podczas powrotu do miasta początkowego.");
        }

        return Tuple.Create(path, totalDistance);
    }
}
