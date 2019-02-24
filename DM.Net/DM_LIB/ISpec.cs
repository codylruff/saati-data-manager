/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:06 AM
 * 
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DM_Lib
{
    // Interface for version control
    // of standardized process
    // specifications
    
    public interface ISpec
    {
        string MaterialId { get; set; }
        DateTime TimeStamp { get; set; }
        string SpecType { get; }
        bool IsDefault { get; set; }
        ISpec ParentSpec { get; set; }
        string Revision { get; set;}

        void SetDefaultProperties();
        string ToString();
    }
}
