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
        static void Main(string[] args)
        {
            // Parse args
            try
            {
            	Console.WriteLine("|--------------------------------------|");
        		Console.WriteLine("|         SAATI Spec Manager           |");
        		Console.WriteLine("|--------------------------------------|");
            	MaterialInputDialog();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: {0}", ex.Message);
                Console.ReadLine();
            }
            
        }
        
        public static void MaterialInputDialog()
        {
        	Console.Write("Enter a material ID :");
        	string input = Console.ReadLine();
            StartViewer(input);    
        }

        public static void StartViewer(string material_id)
        {
            // Create a spec
            Console.WriteLine("---------------------------------------");
            Console.WriteLine("Specifications for {0} :", material_id);
            Console.WriteLine("---------------------------------------");
            try
            {
            	
            	int retVal = manager.LoadStandard(material_id);
            	Console.WriteLine(retVal);
            	if ( retVal == -1) NewMaterialDialogue(material_id);
            	manager.LoadSpecification(retVal == 0 ? material_id : manager.CorrectId);
            	
            }catch
            {
            	NewMaterialDialogue(material_id);            
            }
            finally
            {
            	
            	Console.WriteLine("---------------------------------------");
	            manager.PrintSpecification();
	            Console.WriteLine("---------------------------------------");
	            
	            Console.WriteLine("Would you like to see another specification? (y/n)");
	            if(Console.ReadLine().ToString().ToLower() == "y")
	            {
	            	manager.Specs.Reset();
	            	MaterialInputDialog();
	            }else
	            {
	            	Console.WriteLine("Press enter to exit . . .");
	            	Console.ReadLine();
	            }
            }
        }

        public static void NewMaterialDialogue(string material_id)
        {
        	ISpec spec;
            Console.WriteLine("Material : {0}, Does not exist. Would you like to create it? (y/n)", material_id);

            if (Console.ReadLine().ToLower() == "y")
            {
                Console.WriteLine("Please enter a material type :\n");
                PrintMaterialTypes();
                Console.Write("Material Type : ");
                string spec_type = Console.ReadLine().ToLower();
                Console.WriteLine("---------------------------------------");
                spec = manager.CreateNewMaterial(material_id, spec_type);
                Console.WriteLine(spec.ToString());
                Console.WriteLine("---------------------------------------");
                Console.WriteLine("Would you like to safe this specification? (y/n)");
                
                if (Console.ReadLine().ToLower() == "y")
                {
                	spec.ToString();
                	manager.CommitSpecificationRecord(Factory.CreateRecordFromSpec(spec), "standard_specifications");
                	Console.WriteLine("Specification has been saved.");
                }
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