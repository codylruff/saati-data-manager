/*
 * C#
 * User: CRuff
 * Date: 2/17/2019
 * Time: 12:19 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DM_Lib;

namespace DM_CLI
{
    class Program
    {
        public static SpecManager manager = new SpecManager();
        static void Main()
        {
            // Parse args
            try
            {
            	Console.WriteLine("|--------------------------------------|");
            	Console.WriteLine("|         SAATI Spec Manager           |");
            	Console.WriteLine("|--------------------------------------|");
            	Console.Write("Enter a material ID :");
            	string input = Console.ReadLine();
                StartViewer(input);       
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: {0}", ex.Message);
                Console.ReadLine();
            }
            
        }

        public static void StartViewer(string material_id)
        {
            // Create a spec
            Console.WriteLine("---------------------------------------");
            Console.WriteLine("Specifications for {0} :", material_id);
            Console.WriteLine("---------------------------------------");
            int retVal = manager.LoadStandard(material_id);
            if ( retVal == -1) NewMaterialDialogue(material_id);
            manager.LoadSpecification(retVal == 0 ? material_id : manager.CorrectId);
            Console.WriteLine("---------------------------------------");
            manager.PrintSpecification();
            Console.WriteLine("---------------------------------------");
            Console.WriteLine("Press enter to exit . . .");
            Console.ReadLine();
        }

        public static void NewMaterialDialogue(string material_id)
        {
            Console.WriteLine("Material : {0}, Does not exist. Would you like to create it? (y/n)", material_id);

            if (Console.ReadLine().ToLower() == "y")
            {
                Console.WriteLine("Please enter a material type :\n");
                PrintMaterialTypes();
                Console.Write("Material Type : ");
                string spec_type = Console.ReadLine().ToLower();
                manager.CreateNewMaterial(material_id, spec_type);
            }
            else
            {
                ExitProgram(0);
            }
                
        }

        public static void PrintMaterialTypes()
        {
            foreach(string material in manager.MaterialsList)
            {
                Console.WriteLine(material);
            }
        }

        public static void Usage()
        {
            Console.WriteLine("Help:");
            Console.WriteLine();
            Console.WriteLine("DM_CLI.exe [options] 'material id'");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("-h   Show usage information");
            Console.WriteLine();
            Console.WriteLine("-config   Initialize spec configuration");
            Console.WriteLine();
        }

        public static void ExitProgram(int code)
        {
            Environment.Exit(code);
        }
    }
}