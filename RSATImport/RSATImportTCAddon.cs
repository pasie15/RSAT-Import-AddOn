using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Tricentis.TCAddOns;

//The namespace needs to match the default namespace of the assembly
namespace RSATImportAddOn
{
    public class RSATImportTCAddon : TCAddOn
    {
        public override string UniqueName => "RSATImportAddOn";

        public override string DisplayedName => "RSAT Import AddOn";
    }
}