using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        Stopwatch stopwatch = Stopwatch.StartNew(); 

        string inputFileName = "Dane_TSP_127.xlsx";
        string outputFileName = "Wyniki_127.xlsx";

        if (!File.Exists(inputFileName))
        {
            Console.WriteLine($"Plik {inputFileName} nie istnieje.");
            return;
        }

        var distanceMatrix = ReadExcel(inputFileName);

        if (distanceMatrix == null)
        {
            Console.WriteLine("Nie udało się wczytać danych z pliku.");
            return;
        }

        var populationSizes = new List<int> {100, 500, 1000, 2500};
        var crossoverProbabilities = new List<double> { 0.8 };
        var mutationProbabilities = new List<double> { 0.2 };
        var maxGenerationsList = new List<int> {  500, 1000, 2500, 5000};
        var maxNoImprovementList = new List<int> { 150, 500, 750, 1000};

        var crossoverMethods = new List<string> { "PMX", "OX"};
        var selectionMethods = new List<string> { "Tournament", "Ranking", "Scaling"};

        var results = new List<(int MaxGenerations, int MaxNoImprovement, string CrossoverMethod, string SelectionMethod, 
                                int PopulationSize, double CrossoverProbability, double MutationProbability,
                                double BestResult, string Route, long ExecutionTimeMs)>();


        int totalIterations = populationSizes.Count * crossoverProbabilities.Count * mutationProbabilities.Count
                              * maxGenerationsList.Count * maxNoImprovementList.Count
                              * crossoverMethods.Count * selectionMethods.Count;
        int currentIteration = 0;

        int repeatCount = 0; 

        while (repeatCount < 3) 
        {
            repeatCount++; 
            foreach (var populationSize in populationSizes)
            {
                foreach (var crossoverProbability in crossoverProbabilities)
                {
                    foreach (var mutationProbability in mutationProbabilities)
                    {
                        foreach (var maxGenerations in maxGenerationsList)
                        {
                            foreach (var maxNoImprovement in maxNoImprovementList)
                            {
                                foreach (var crossoverMethod in crossoverMethods)
                                {
                                    foreach (var selectionMethod in selectionMethods)
                                    {

                                        Stopwatch iterationStopwatch = Stopwatch.StartNew();

                                       
                                        double progress = (double)currentIteration / totalIterations * 100;

                                        TimeSpan elapsedTime = stopwatch.Elapsed;

                                        
                                        Console.SetCursorPosition(0, 0); 
                                        Console.WriteLine($"Postęp: {progress:F2}%");
                                        Console.WriteLine($"Czas: {elapsedTime:hh\\:mm\\:ss}");
                                        Console.WriteLine($"Powtórzenie {repeatCount} z 1");

                                        var result = GeneticAlgorithm(distanceMatrix, populationSize, crossoverProbability, mutationProbability,
                                                                      maxGenerations, maxNoImprovement, crossoverMethod, selectionMethod);

                                        iterationStopwatch.Stop(); 


                                        string route = string.Join(", ", result.Item1.Select(city => city + 1)); 
                                        results.Add((maxGenerations, maxNoImprovement, crossoverMethod, selectionMethod, 
                                                     populationSize, crossoverProbability, mutationProbability, result.Item2, route,
                                                     iterationStopwatch.ElapsedMilliseconds));

                                        currentIteration++;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            SaveResultsToExcel(results, outputFileName);
            Console.WriteLine($"Wyniki zapisane w pliku {outputFileName}");
        }

        stopwatch.Stop(); 
        Console.WriteLine($"Czas wykonania: {stopwatch.Elapsed:hh\\:mm\\:ss}");
    }





    static Tuple<List<int>, double> GeneticAlgorithm(double[,] distanceMatrix, int populationSize, double crossoverProbability, double mutationProbability, int maxGenerations, int maxNoImprovement, string crossoverMethod, string selectionMethod)
    {
        int citiesCount = distanceMatrix.GetLength(0);
        var random = new Random();

        List<List<int>> population = GenerateInitialPopulation(populationSize, citiesCount, random);
        double bestFitness = double.MaxValue;
        List<int> bestSolution = null;
        int generationsWithoutImprovement = 0;

        for (int generation = 0; generation < maxGenerations; generation++)
        {
            population = SelectParents(population, distanceMatrix, selectionMethod, random);

            if (random.NextDouble() < crossoverProbability)
                population = ApplyCrossover(population, crossoverMethod, random);

            if (random.NextDouble() < mutationProbability)
                population = ApplyMutation(population, random);

            var currentBest = EvaluatePopulation(population, distanceMatrix);

            if (currentBest.Item2 < bestFitness)
            {
                bestFitness = currentBest.Item2;
                bestSolution = currentBest.Item1;
                generationsWithoutImprovement = 0;
            }
            else
            {
                generationsWithoutImprovement++;
            }

            if (generationsWithoutImprovement >= maxNoImprovement)
                break;
        }

        return Tuple.Create(bestSolution, bestFitness);
    }

    static List<List<int>> GenerateInitialPopulation(int size, int citiesCount, Random random)
    {
        var population = new List<List<int>>();
        for (int i = 0; i < size; i++)
        {
            var individual = Enumerable.Range(0, citiesCount).ToList();
            individual = individual.OrderBy(_ => random.Next()).ToList();
            population.Add(individual);
        }
        return population;
    }

    static List<List<int>> SelectParents(List<List<int>> population, double[,] distanceMatrix, string method, Random random)
    {
        if (method == "Scaling")
        {
            var fitnesses = population.Select(ind => EvaluateIndividual(ind, distanceMatrix)).ToList();

            double averageFitness = fitnesses.Average();
            double maxFitness = fitnesses.Max();

            double shift = 2 * averageFitness - maxFitness;

            var scaledFitnesses = fitnesses.Select(f => Math.Max(0, f - shift)).ToList();

            double totalScaledFitness = scaledFitnesses.Sum();
            var probabilities = scaledFitnesses.Select(f => f / totalScaledFitness).ToList();

            return population.Select(_ => RouletteWheelSelection(population, probabilities, random)).ToList();
        }
        else if (method == "Ranking")
        {
            var ranked = population.Select(ind => Tuple.Create(ind, EvaluateIndividual(ind, distanceMatrix)))
                                   .OrderBy(t => t.Item2) 
                                   .ToList();


            int populationSize = ranked.Count;
            var probabilities = Enumerable.Range(1, populationSize)
                                          .Select(i => i / (double)populationSize) 
                                          .ToList();

            return ranked.Select(_ => RouletteWheelSelection(ranked.Select(r => r.Item1).ToList(), probabilities, random)).ToList();
        }
        else if (method == "Tournament")
        {
            int tournamentSize = 3;
            return TournamentSelection(population, distanceMatrix, tournamentSize, random);
        }

        return population;
    }

    static List<List<int>> TournamentSelection(List<List<int>> population, double[,] distanceMatrix, int tournamentSize, Random random)
    {
        var selectedParents = new List<List<int>>();

        for (int i = 0; i < population.Count; i++)
        {
            var tournamentParticipants = new List<List<int>>();

            for (int j = 0; j < tournamentSize; j++)
            {
                int randomIndex = random.Next(population.Count);
                tournamentParticipants.Add(population[randomIndex]);
            }

            var bestParticipant = tournamentParticipants
                .OrderBy(individual => EvaluateIndividual(individual, distanceMatrix))
                .First();

            selectedParents.Add(bestParticipant);
        }

        return selectedParents;
    }


    static List<int> RouletteWheelSelection(List<List<int>> population, List<double> probabilities, Random random)
    {
        double r = random.NextDouble();
        double cumulative = 0;

        for (int i = 0; i < probabilities.Count; i++)
        {
            cumulative += probabilities[i];
            if (r <= cumulative)
                return new List<int>(population[i]);
        }

        return new List<int>(population[random.Next(population.Count)]);
    }

    static List<List<int>> ApplyCrossover(List<List<int>> population, string method, Random random)
    {
        var newPopulation = new List<List<int>>();
        for (int i = 0; i < population.Count; i += 2)
        {
            if (i + 1 >= population.Count)
                break;

            var parent1 = population[i];
            var parent2 = population[i + 1];

            if (method == "PMX")
            {
                var offspring = PartiallyMappedCrossover(parent1, parent2, random);
                newPopulation.Add(offspring.Item1);
                newPopulation.Add(offspring.Item2);
            }
            else if (method == "OX")
            {
                var offspring = OrderCrossover(parent1, parent2, random);
                newPopulation.Add(offspring.Item1);
                newPopulation.Add(offspring.Item2);
            }
        }

        return newPopulation;
    }

    static Tuple<List<int>, List<int>> PartiallyMappedCrossover(List<int> parent1, List<int> parent2, Random random)
    {
        int size = parent1.Count;
        int start = random.Next(size);
        int end = random.Next(start, size);

        var offspring1 = new int[size];
        var offspring2 = new int[size];
        Array.Fill(offspring1, -1);
        Array.Fill(offspring2, -1);

        for (int i = start; i <= end; i++)
        {
            offspring1[i] = parent1[i];
            offspring2[i] = parent2[i];
        }

        FillOffspring(offspring1, parent2, start, end);
        FillOffspring(offspring2, parent1, start, end);

        return Tuple.Create(offspring1.ToList(), offspring2.ToList());
    }

    static Tuple<List<int>, List<int>> OrderCrossover(List<int> parent1, List<int> parent2, Random random)
    {
        int size = parent1.Count;
        int start = random.Next(size);
        int end = random.Next(start, size);

        var offspring1 = new int[size];
        var offspring2 = new int[size];
        Array.Fill(offspring1, -1);
        Array.Fill(offspring2, -1);

        for (int i = start; i <= end; i++)
        {
            offspring1[i] = parent1[i];
            offspring2[i] = parent2[i];
        }

        int currentIndex1 = (end + 1) % size;
        int currentIndex2 = (end + 1) % size;

        foreach (var gene in parent2.Skip(end + 1).Concat(parent2.Take(end + 1)))
        {
            if (!offspring1.Contains(gene))
            {
                offspring1[currentIndex1] = gene;
                currentIndex1 = (currentIndex1 + 1) % size;
            }
        }

        foreach (var gene in parent1.Skip(end + 1).Concat(parent1.Take(end + 1)))
        {
            if (!offspring2.Contains(gene))
            {
                offspring2[currentIndex2] = gene;
                currentIndex2 = (currentIndex2 + 1) % size;
            }
        }

        return Tuple.Create(offspring1.ToList(), offspring2.ToList());
    }

    static void FillOffspring(int[] offspring, List<int> parent, int start, int end)
    {
        for (int i = 0; i < parent.Count; i++)
        {
            if (offspring.Contains(parent[i]))
                continue;

            for (int j = 0; j < offspring.Length; j++)
            {
                if (offspring[j] == -1)
                {
                    offspring[j] = parent[i];
                    break;
                }
            }
        }
    }

    static List<List<int>> ApplyMutation(List<List<int>> population, Random random)
    {
        foreach (var individual in population)
        {
            int index1 = random.Next(individual.Count);
            int index2 = random.Next(individual.Count);
            (individual[index1], individual[index2]) = (individual[index2], individual[index1]);
        }
        return population;
    }

    static Tuple<List<int>, double> EvaluatePopulation(List<List<int>> population, double[,] distanceMatrix)
    {
        var bestIndividual = population.OrderBy(ind => EvaluateIndividual(ind, distanceMatrix)).First();
        double bestFitness = EvaluateIndividual(bestIndividual, distanceMatrix);
        return Tuple.Create(bestIndividual, bestFitness);
    }

    static double EvaluateIndividual(List<int> individual, double[,] distanceMatrix)
    {
        double totalDistance = 0.0;
        for (int i = 0; i < individual.Count - 1; i++)
        {
            totalDistance += distanceMatrix[individual[i], individual[i + 1]];
        }
        totalDistance += distanceMatrix[individual.Last(), individual.First()];
        return totalDistance;
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

    static void SaveResultsToExcel(List<(int MaxGenerations, int MaxNoImprovement, string CrossoverMethod, string SelectionMethod, 
                                    int PopulationSize, double CrossoverProbability, double MutationProbability,
                                    double BestResult, string Route, long ExecutionTimeMs)> results, string fileName)
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(fileName)))
            {
                var worksheet = package.Workbook.Worksheets[0] ?? package.Workbook.Worksheets.Add("Wyniki");

                if (worksheet.Dimension == null)
                {
                    worksheet.Cells[1, 1].Value = "Liczba pokoleń";
                    worksheet.Cells[1, 2].Value = "Liczba iteracji bez poprawy";
                    worksheet.Cells[1, 3].Value = "Metoda krzyżowania";
                    worksheet.Cells[1, 4].Value = "Metoda selekcji";
                    worksheet.Cells[1, 5].Value = "Wielkość populacji";
                    worksheet.Cells[1, 6].Value = "Prawdopodobieństwo krzyżowania";
                    worksheet.Cells[1, 7].Value = "Prawdopodobieństwo mutacji";
                    worksheet.Cells[1, 8].Value = "Najkrótszy dystans";
                    worksheet.Cells[1, 9].Value = "Trasa";
                    worksheet.Cells[1, 10].Value = "Czas wykonania (ms)";
                }

                for (int i = 0; i < results.Count; i++)
                {
                    var result = results[i];
                    worksheet.Cells[i + 2, 1].Value = result.MaxGenerations;
                    worksheet.Cells[i + 2, 2].Value = result.MaxNoImprovement;
                    worksheet.Cells[i + 2, 3].Value = result.CrossoverMethod;
                    worksheet.Cells[i + 2, 4].Value = result.SelectionMethod;
                    worksheet.Cells[i + 2, 5].Value = result.PopulationSize;
                    worksheet.Cells[i + 2, 6].Value = result.CrossoverProbability;
                    worksheet.Cells[i + 2, 7].Value = result.MutationProbability;
                    worksheet.Cells[i + 2, 8].Value = result.BestResult;
                    worksheet.Cells[i + 2, 9].Value = result.Route;
                    worksheet.Cells[i + 2, 10].Value = result.ExecutionTimeMs;
                }

                // Zapis do pliku
                package.Save();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Błąd podczas zapisu pliku: " + ex.Message);
        }
    }

}


