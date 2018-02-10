using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using ABPHelper.Extensions;
using ABPHelper.Models.HelperModels;
using ABPHelper.Models.TemplateModels;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using RazorEngine.Templating;
using Engine = RazorEngine.Engine;

namespace ABPHelper.Helper
{
    public class AddNewBusinessHelper : HelperBase<AddNewBusinessModel>
    {
        private Project _appProj;

        private Project _webProj;

        private readonly string _appName;

        private StatusBar _statusBar;

        private int _totalSteps;

        private int _steps; 

        public AddNewBusinessHelper(IServiceProvider serviceProvider) : base(serviceProvider)
        {
            _appName = Dte.Solution.Properties.Item("Name").Value.ToString();

            _statusBar = Dte.StatusBar;
        }
        #region Get Solution's projects
        public static DTE2 GetActiveIDE()
        {
            // Get an instance of currently running Visual Studio IDE.
            DTE2 dte2 = Package.GetGlobalService(typeof(DTE)) as DTE2;
            return dte2;
        }
        public static IList<Project> Projects()
        {
            Projects projects = GetActiveIDE().Solution.Projects;
            List<Project> list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                
                if (project == null)
                {
                    continue;
                }
               
                if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(project));
                }
                else
                {
                    list.Add(project);
                }
            }
            return list;
        }
        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                {
                    continue;
                }

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(subProject));
                }
                else
                {
                    list.Add(subProject);
                }
            }
            return list;
        }
        #endregion

        public override bool CanExecute(AddNewBusinessModel parameter)
        {
            var projectList = Projects();
            foreach (var project in projectList)
            {
                var m = Regex.Match(project.Name, @"(.+)\.Application");
                if (m.Success)
                {
                    _appProj = project;
                   
                }
                if (Regex.IsMatch(project.Name, @"(.+)\.Web"))
                {
                    if (!project.Name.Contains(".WebApi"))
                        _webProj = project;
                }
                if (_appProj != null && _webProj != null) break;
            }

            if (_appProj == null)
            {
                Utils.MessageBox("Cannot find the Application project. Please ensure that your are in the ABP solution.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }
            if (_webProj == null)
            {
                Utils.MessageBox("Cannot find the Web project. Please ensure that your are in the ABP solution.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return false;
            }
            return true;
        }

        public override void Execute(AddNewBusinessModel parameter)
        {
            try
            {
                _totalSteps = parameter.ViewFiles.Count() * 2 + 2;
                _steps = 1;

                var folder = AddDeepFolder(_appProj.ProjectItems, parameter.ServiceFolder);
                AddDeepFolder(folder.ProjectItems, "Dto");
                CreateServiceFile(parameter, folder);
                CreateServiceInterfaceFile(parameter, folder);
                
                folder = AddAbpDeepFolder(_webProj.ProjectItems, parameter.ViewFolder);

                CreateViewFiles(parameter, folder);

                Utils.MessageBox("Done!");
            }
            catch (Exception e)
            {
                Utils.MessageBox("Generation failed.\r\nException: {0}", MessageBoxButton.OK, MessageBoxImage.Exclamation, e.Message);
            }
            finally
            {
                _statusBar.Progress(false);
            }
        }

        private void CreateViewFiles(AddNewBusinessModel parameter, ProjectItem folder)
        {
            foreach (var viewFileViewModel in parameter.ViewFiles)
            {
                var model = new ViewFileModel
                {
                    BusinessName = parameter.BusinessName,
                    Namespace = GetNamespace(parameter.ViewFolder),
                    FileName = viewFileViewModel.FileName,
                    IsPopup = viewFileViewModel.IsPopup,
                    ViewFolder = parameter.ViewFolder,
                    ViewFiles = parameter.ViewFiles
                };
                foreach (var ext in new[] { ".cshtml", ".js" })
                {
                    var fileName = viewFileViewModel.FileName + ext;
                    _statusBar.Progress(true, $"Generating view file: {fileName}", _steps++, _totalSteps);
                    if (FindProjectItem(folder, fileName, ItemType.PhysicalFile) != null) continue;
                    string content = Engine.Razor.RunCompile(ext == ".cshtml" ? "CshtmlTemplate" : "JsTemplate", typeof(ViewFileModel), model);
                    CreateAndAddFile(folder, fileName, content);
                }
            }
        }

        private string GetNamespace(string viewFolder)
        {
            return string.Join(".", viewFolder.Split('\\').Select(s => s.LowerFirstChar()));
        }


        private void CreateServiceFile(AddNewBusinessModel parameter, ProjectItem folder)
        {
            var fileName = parameter.ServiceName + ".cs";
            _statusBar.Progress(true, $"Generating service file: {fileName}", _steps++, _totalSteps);
            if (FindProjectItem(folder, fileName, ItemType.PhysicalFile) != null) return;
            var model = new ServiceFileModel
            {
                AppName = _appName,
                Namespace = GetNamespace(parameter),
                InterfaceName = parameter.ServiceInterfaceName,
                ServiceName = parameter.ServiceName
            };
            string content = Engine.Razor.RunCompile("ServiceFileTemplate", typeof(ServiceFileModel), model);
            CreateAndAddFile(folder, fileName, content);
        }

        private void CreateServiceInterfaceFile(AddNewBusinessModel parameter, ProjectItem folder)
        {
            var fileName = parameter.ServiceInterfaceName + ".cs";
            _statusBar.Progress(true, $"Generating interface file: {fileName}", _steps++, _totalSteps);
            if (FindProjectItem(folder, fileName, ItemType.PhysicalFile) != null) return;
            var model = new ServiceInterfaceFileModel
            {
                Namespace = GetNamespace(parameter),
                InterfaceName = parameter.ServiceInterfaceName
            };
            string content = Engine.Razor.RunCompile("ServiceInterfaceFileTemplate", typeof(ServiceInterfaceFileModel), model);
            CreateAndAddFile(folder, fileName, content);
        }

        private string GetNamespace(AddNewBusinessModel parameter)
        {
            var str = parameter.ServiceFolder.Replace('\\', '.');
            return $"{_appName}.{str}";
        }

        private ProjectItem AddDeepFolder(ProjectItems parentItems, string deepFolder)
        {
            ProjectItem addedFolder = null;

                foreach (var folder in deepFolder.Split('\\'))
                {
                    var projectItem = FindProjectItem(parentItems, folder, ItemType.PhysicalFolder);
                    addedFolder = projectItem ?? parentItems.AddFolder(folder);
                    parentItems = addedFolder.ProjectItems;
                }
            return addedFolder;
        }

        /// <summary>
        ///getting the Full project path inorder to get the path of the ABP.dll
        ///in ABP.dll version number 1.0.1.5 it's uses the path 
        ///C:\Sestek.CallSteering\SestekCallSteering\Development\Sestek.CallSteering\Sestek.CallSteering.Web\App\Main\Views
        ///for the views.
        ///while in other versions it uses 
        ///C:\Sestek.CallSteering\SestekCallSteering\Development\Sestek.CallSteering\Sestek.CallSteering.Web\App\Main\views
        ///The difference in "View" world specialy in the first letter V
        /// </summary>
        /// <param name="parentItems"></param>
        /// <param name="deepFolder"></param>
        /// <returns>ProjectItem</returns>
        private ProjectItem AddAbpDeepFolder(ProjectItems parentItems, string deepFolder)
        {
            ProjectItem addedFolder = null;
            var projectPath = parentItems.ContainingProject.FullName;
            projectPath = projectPath.Remove(projectPath.IndexOf(".Web"));
            var tmpIndex = findIndex(projectPath, '\\');
            projectPath = projectPath.Remove(tmpIndex);

            projectPath += "packages\\";

            List<string> getFiles = Directory.GetDirectories(projectPath).Where(x=>x.Contains("Abp")).ToList();

            foreach (var file in getFiles)
            {
                //Checking ABP version
                var fileName = Path.GetFileName(file);
                var versionNumber = fileName.Substring(4);
                if (versionNumber == "1.0.1.5")
                {
                    deepFolder = deepFolder.Remove(deepFolder.IndexOf("views"), 5).Insert(deepFolder.IndexOf("views"), "Views");
                    break;
                }
            }

            foreach (var folder in deepFolder.Split('\\'))
            {
               
                var projectItem = FindProjectItem(parentItems, folder, ItemType.PhysicalFolder);
                addedFolder = projectItem ?? parentItems.AddFolder(folder);
                parentItems = addedFolder.ProjectItems;
            }   
            return addedFolder;
        }

        private int findIndex(string searchPattern,char c)
        {
            for (var i = searchPattern.Length -1; i >= 0; i--)
            {
                if (searchPattern[i] == c) return i+1;
            }

            return 0;
        }
    }
}