using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace LocateFolder
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class LocateFolderCmd
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("4b1b60f4-3cf6-42c9-b1e5-dc2931e5cb3a");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly AsyncPackage package;

    /// <summary>
    /// Initializes a new instance of the <see cref="LocateFolderCmd"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    private LocateFolderCmd(AsyncPackage package, OleMenuCommandService commandService)
    {
      this.package = package ?? throw new ArgumentNullException(nameof(package));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandID = new CommandID(CommandSet, CommandId);
      var menuItem = new MenuCommand(this.ExecuteAsync, menuCommandID);
      commandService.AddCommand(menuItem);
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static LocateFolderCmd Instance
    {
      get;
      private set;
    }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
    {
      get
      {
        return this.package;
      }
    }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static async Task InitializeAsync(AsyncPackage package)
    {
      // Verify the current thread is the UI thread - the call to AddCommand in LocateFolderCmd's constructor requires
      // the UI thread.
      ThreadHelper.ThrowIfNotOnUIThread();

      OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
      Instance = new LocateFolderCmd(package, commandService);
    }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private async void ExecuteAsync(object sender, EventArgs e)
    {

      ThreadHelper.ThrowIfNotOnUIThread();
      var dte = (DTE2) await this.ServiceProvider.GetServiceAsync(typeof(DTE));
      var uiHierachy = (UIHierarchy) dte.Windows.Item("{3AE79031-E1BC-11D0-8F78-00A0C9110057}").Object;
      if (uiHierachy.SelectedItems is object[] selectedItems)
      {
        LocateFile.FilesOrFolders(from t in selectedItems
                                  where (t as UIHierarchyItem)?.Object is ProjectItem
                                  select ((ProjectItem)((UIHierarchyItem)t).Object).FileNames[1]);
      }
    }
  }
}
