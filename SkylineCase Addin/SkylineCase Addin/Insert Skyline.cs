// Skyline Case Addin | stru.ca | Copyright Neural Systems Inc

#region Namespaces
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
#endregion

namespace Skylineˉcaseˉaddin
{
    internal sealed class Insertˉskylins
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("96bd2f5f-05ca-4228-93ee-236bf40d75a6");
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Insertˉskylins"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Insertˉskylins(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static Insertˉskylins Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public DTE Dte;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, DTE Dte)
        {
            // Verify the current thread is the UI thread - the call to AddCommand in Insert_Skyline's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new Insertˉskylins(package, commandService);        
            Instance.Dte = Dte;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            EnvDTE.Document doc = this.Dte.ActiveDocument;
            if (doc == null)
            {
                return;
            }
            dynamic textSelection = doc.Selection;
            if (textSelection == null)
            {
                return;
            }
            textSelection.Insert("ˉ");
        }
    }
}
