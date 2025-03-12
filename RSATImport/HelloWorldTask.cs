using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;

namespace RSATImportAddon.Tasks
{
    public class HelloWorldTask : TCAddOnTask
    {
        public override string Name => "Hello World";

        public override bool IsTaskPossible(TCObject obj) => true;

        public override Type ApplicableType => typeof(TCObject);

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            taskContext.ShowMessageBox("Hello World", "Hello World from the context menu");
            return objectToExecuteOn;
        }
    }
}
