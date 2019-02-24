/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:10 AM
 * 
 * 
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DM_Lib
{
    public class SpecsCollection : IEnumerable
    {
        private List<ISpec> _specsCollection;
        private ISpec _defaultSpec;
        public string SpecType { get; private set; }
        // For easy access to the default spec
        // default spec can only be written from data access
        // TODO: This is broken ***
        public ISpec DefaultSpec 
        {
            get { return _defaultSpec; }
            set
            {
            	_specsCollection.Add(value);
                _defaultSpec = value;
               
            }
        }

        public SpecsCollection()
        {
            _specsCollection = new List<ISpec>();
        }

        public void Add(ISpec spec)
        {
            _specsCollection.Add(spec);
        }
        
        public ISpec SpecByMaterialId(string material_id)
        {
            foreach (ISpec spec in _specsCollection)
            {
                if (spec.MaterialId == material_id)
                    return spec;
            }
            // returns null if no spec is found.
            Debug.Print("No spec found with this id");
            return null;
        }

        public IEnumerator GetEnumerator()
        {
            // Return the array object's IEnumerator.
            return _specsCollection.GetEnumerator();
        }
    }
}