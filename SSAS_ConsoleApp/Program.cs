using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.AdomdClient;

namespace SSAS_ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Data Source=localhost;Integrated Security=SSPI";

            // Create the server object and connect to the instance..
            Server server = new Server();
            server.Connect(connectionString);

            // Get count of all databases in the connected instance..
            Console.WriteLine("ServerName: " + server.ConnectionInfo.Server.ToString());
            Console.WriteLine("Total databases: " + server.Databases.Count);
            Console.WriteLine();

            // Get list of all databases in the instance..
            Console.WriteLine("List of all databases: ");
            foreach (var database in server.Databases){
                Console.WriteLine(database.ToString());
            }
            Console.WriteLine();

            // Get list of all cubes in a database..
            Console.WriteLine("List of all cubes in database: RCX");
            foreach(var cube in server.Databases.GetByName("RCX").Cubes)
            {
                Console.WriteLine(cube.ToString());
            }
            Console.WriteLine();

            // Get list of all dimensions in a database..
            Console.WriteLine("List of all dimensions in database: RCX");
            foreach (var dimension in server.Databases.GetByName("RCX").Dimensions)
            {
                Console.WriteLine(dimension.ToString());
            }
            Console.WriteLine();

            // Get list of all dimensions in a cube..
            Console.WriteLine("List of all dimensions in cube: Expositions");
            foreach (var dimension in server.Databases.GetByName("RCX").Cubes.GetByName("Expositions").Dimensions)
            {
                Console.WriteLine(dimension.ToString());
            }
            Console.WriteLine();

            // Get list of all measures in a cube..
            Console.WriteLine("List of all dimensions in cube: Expositions");
            foreach (var measureGroup in server.Databases.GetByName("RCX").Cubes.GetByName("Expositions").MeasureGroups)
            {
                Console.WriteLine(measureGroup.ToString());
            }
            Console.WriteLine();

            // Get cube properties...
            Console.WriteLine("Properties for cube: Expositions");
            Console.WriteLine("Last Processed: " + server.Databases.GetByName("Analysis Services Tutorial").Cubes.GetByName("Analysis Services Tutorial").LastProcessed);
            Console.WriteLine();

            //Process a cube
            server.Databases.GetByName("Analysis Services Tutorial").Cubes.GetByName("Analysis Services Tutorial").Process(ProcessType.ProcessFull);
            Console.WriteLine("Last Processed: " + server.Databases.GetByName("Analysis Services Tutorial").Cubes.GetByName("Analysis Services Tutorial").LastProcessed);

            Console.WriteLine();
            Console.WriteLine("Trying to read data from cube");
            GetData(connectionString);




            Console.ReadKey();

        }

        static void GetData(string connectionString)
        {
            AdomdConnection connection = new AdomdConnection(connectionString + ";Catalog=Analysis Services Tutorial");
            connection.Open();

            //string queryCommand = @" SELECT NON EMPTY { [Measures].[Internet Sales Count] } ON COLUMNS FROM [Analysis Services Tutorial] ";
            string queryCommand = @" SELECT NON EMPTY { [Measures].[Internet Sales Count] } ON COLUMNS, NON EMPTY { ([Order Date].[Calendar Year].[Calendar Year].ALLMEMBERS ) }  ON ROWS FROM [Analysis Services Tutorial] ";

            AdomdCommand command = new AdomdCommand(queryCommand, connection);

            AdomdDataReader dataReader = command.ExecuteReader();

            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0].ToString() + " | " + dataReader[1].ToString());
            }

            dataReader.Close();
            connection.Close();
        }

    }
}
