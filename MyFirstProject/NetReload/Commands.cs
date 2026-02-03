using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(NetReload.Commands))]

namespace NetReload
{

	public class Commands
	{

		// NRL command - Short alias for RELOAD
		[CommandMethod("NRL")]
		public static void Nrl()
		{
			Reload();
		}

		// RELOAD command - Hot reload project without Visual Studio.
		[CommandMethod("RELOAD")]
		public static void Reload()
		{
			Autodesk.AutoCAD.ApplicationServices.Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
			Editor ed = doc.Editor;

			try
			{
				// 1. Determine Project Path
				// Since this is a standalone reloader, we point it to the main project we want to reload.
				// You can change this path to reload different projects.
				string projectDir = @"c:\Google\Autocad 2026_API\Autocad 2026_API\MyFirstProject";

				if (string.IsNullOrEmpty(projectDir) || !Directory.Exists(projectDir))
				{
					ed.WriteMessage($"\nCould not find project directory: {projectDir}. *Cancel*");
					return;
				}

				string? csprojFile = Directory.GetFiles(projectDir, "*.csproj").FirstOrDefault();
				if (csprojFile == null)
				{
					ed.WriteMessage("\nCould not find .csproj file. *Cancel*");
					return;
				}

				string projectName = Path.GetFileNameWithoutExtension(csprojFile);

				// Generate unique assembly name to avoid "Assembly already loaded" error
				string randomName = Path.GetRandomFileName().Replace(".", "");
				string uniqueAssemblyName = $"{projectName}_{randomName}";

				ed.WriteMessage($"\nReloading project: {projectName} (New Assembly: {uniqueAssemblyName})...");

				// 2. Run dotnet build with unique AssemblyName
				string configuration = "Debug";

				ProcessStartInfo psi = new ProcessStartInfo
				{
					FileName = "dotnet",
					Arguments = $"build -c {configuration} /p:AssemblyName={uniqueAssemblyName}",
					WorkingDirectory = projectDir,
					RedirectStandardOutput = true,
					RedirectStandardError = true,
					UseShellExecute = false,
					CreateNoWindow = true
				};

				using (Process? process = Process.Start(psi))
				{
					if (process == null)
					{
						ed.WriteMessage("\nFailed to start dotnet process.");
						return;
					}

					StringBuilder outputBuilder = new StringBuilder();
					StringBuilder errorBuilder = new StringBuilder();

					process.OutputDataReceived += (sender, args) =>
					{
						if (args.Data != null) outputBuilder.AppendLine(args.Data);
					};
					process.ErrorDataReceived += (sender, args) =>
					{
						if (args.Data != null) errorBuilder.AppendLine(args.Data);
					};

					process.BeginOutputReadLine();
					process.BeginErrorReadLine();

					process.WaitForExit();

					string output = outputBuilder.ToString();
					string error = errorBuilder.ToString();

					if (process.ExitCode != 0)
					{
						ed.WriteMessage($"\nBuild Failed:\n{output}\n{error}");
						return;
					}
				}

				// 3. Locate Output DLL
				string binDir = Path.Combine(projectDir, "bin", configuration);
				var dllFiles = Directory.GetFiles(binDir, $"{uniqueAssemblyName}.dll", SearchOption.AllDirectories)
					.Select(f => new FileInfo(f))
					.OrderByDescending(f => f.LastWriteTime)
					.ToList();

				if (dllFiles.Count == 0)
				{
					ed.WriteMessage($"\nCould not find built DLL ({uniqueAssemblyName}.dll). *Cancel*");
					return;
				}

				FileInfo sourceDll = dllFiles.First();

				// 4. Copy to NetReload directory (inside the target project's bin)
				string netReloadDir = Path.Combine(projectDir, "bin", "NetReload");
				if (!Directory.Exists(netReloadDir))
				{
					Directory.CreateDirectory(netReloadDir);
				}

				string destDllPath = Path.Combine(netReloadDir, sourceDll.Name);
				string destPdbPath = Path.ChangeExtension(destDllPath, "pdb");

				File.Copy(sourceDll.FullName, destDllPath, true);

				string sourcePdb = Path.ChangeExtension(sourceDll.FullName, "pdb");
				if (File.Exists(sourcePdb))
				{
					File.Copy(sourcePdb, destPdbPath, true);
				}

				// Clean up the artifact from bin
				try
				{
					File.Delete(sourceDll.FullName);
					if (File.Exists(sourcePdb)) File.Delete(sourcePdb);
				}
				catch { /* Ignore cleanup errors */ }

				// 5. Load the new DLL
				Assembly.LoadFrom(destDllPath);
				ed.WriteMessage($"\nNETRELOAD complete. Loaded: {uniqueAssemblyName}.dll");

				// 6. Auto-execute CIVIL_RIBBON command to show the Ribbon
				doc.SendStringToExecute("CIVIL_RIBBON ", true, false, false);

			}
			catch (System.Exception ex)
			{
				ed.WriteMessage($"\nError: {ex.Message}");
			}
		}
	}
}
