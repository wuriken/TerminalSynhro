using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;

namespace TerminalSynhro
{
    [Guid("96E63303-951D-428F-A58A-9EB6E405C233")]
    internal interface ISycnhro
    {
        [DispId(1)]

        bool CheckConnectionWithTerminal();

        bool GetInvoicesFromTerminal();

        bool LoadDataToTerminal();
    }

    [Guid("1B1709D3-9E60-439C-8260-566716DBB0CA"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISynchroEvents
    {
    }

    [Guid("42CABCD9-BFE0-4B4D-BDE1-2D6D80D55164"), ClassInterface(ClassInterfaceType.None),
     ComSourceInterfaces(typeof (ISynchroEvents))]
    public class Synchro : ISycnhro
    {
        public string ErrorMessages { get; private set; }

        public bool CheckConnectionWithTerminal()
        {
            return new DirectoryInfo(MConvert.PathToRootTerminalFolder).Exists;
        }

        public bool GetInvoicesFromTerminal()
        {
            try
            {
                return MConvert.GetInvoicesFromTerminal();
            }
            catch (Exception ex)
            {
                ErrorMessages = ex.Message;
                return false;
            }
        }

        public bool LoadDataToTerminal()
        {
            return MConvert.LoadDataToTerminal();
        }


    }
}
