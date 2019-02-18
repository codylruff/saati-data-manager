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
    //Public Property Get MaterialId() As String: End Property ' sap code, style # etc...
    //Public Property Get Properties() As Dictionary: End Property ' Must contain MaterialId
    //Public Property Get SpecType() As String: End Property
    //Public Property Get ParentSpec() As ISpec: End Property
    //Public Property Let JsonText(value As String): End Property
    //Public Property Get JsonText() As String: End Property ' If = vbNullString -> ObjectToJson()
    //Public Sub JsonToObject(jsonText As String) : End Sub ' Map json to the spec
    //Public Function ObjectToJson() : As String End Sub ' Store spec Properties dictionary as json
    //Public Sub SetDefaultProperties() : End Sub ' Use create "standard" specification
    
    public interface ISpec
    {
        string MaterialId { get; set; }
        DateTime TimeStamp { get; set; }
        string SpecType { get; }
        bool IsDefault { get; set; }
        ISpec ParentSpec { get; set; }

        void SetDefaultProperties();
        string ToString();
    }
}
