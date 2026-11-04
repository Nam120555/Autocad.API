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
		// Theo dõi số lần reload để dọn rác
		private static int _reloadCount = 0;

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

			Stopwatch sw = Stopwatch.StartNew();

			try
			{
				// 1. Determine Project Path
				string projectDir = @"D:\Autocad 2026_API\Autocad 2026_API\MyFirstProject";

				string? executingDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
				string? currentSearch = executingDir;
				while (!string.IsNullOrEmpty(currentSearch) && Directory.Exists(currentSearch))
				{
					string[] csProjFiles = Directory.GetFiles(currentSearch, "*.csproj");
					string? targetCsproj = csProjFiles.FirstOrDefault(f => !Path.GetFileName(f).Equals("NetReload.csproj", StringComparison.OrdinalIgnoreCase));
					if (targetCsproj != null)
					{
						projectDir = currentSearch;
						break;
					}
					DirectoryInfo? parent = Directory.GetParent(currentSearch);
					if (parent == null) break;
					currentSearch = parent.FullName;
				}

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
				//    TỐI ƯU: Dùng --no-restore để bỏ qua NuGet restore (rất tốn thời gian)
				//    NuGet đã được restore ở lần build đầu tiên, các lần sau không cần nữa.
				string configuration = "Debug";

				ProcessStartInfo psi = new ProcessStartInfo
				{
					FileName = "dotnet",
					// --no-restore: Bỏ qua bước restore NuGet (tiết kiệm ~30-60 giây)
					// --nologo: Bỏ banner thừa
					// -v q: Verbose = quiet, giảm output
					Arguments = $"build -c {configuration} /p:AssemblyName={uniqueAssemblyName} --no-restore --nologo -v q",
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
						// Nếu lỗi do thiếu restore, tự động retry với restore
						if (output.Contains("NETSDK1004") || output.Contains("restore") || error.Contains("NETSDK1004"))
						{
							ed.WriteMessage("\n[NRL] NuGet chưa restore, đang restore lần đầu...");
							psi.Arguments = $"build -c {configuration} /p:AssemblyName={uniqueAssemblyName} --nologo -v q";
							
							using (Process? retryProcess = Process.Start(psi))
							{
								if (retryProcess == null) { ed.WriteMessage("\nFailed to start retry process."); return; }
								
								outputBuilder.Clear();
								errorBuilder.Clear();
								
								retryProcess.OutputDataReceived += (sender, args) => { if (args.Data != null) outputBuilder.AppendLine(args.Data); };
								retryProcess.ErrorDataReceived += (sender, args) => { if (args.Data != null) errorBuilder.AppendLine(args.Data); };
								retryProcess.BeginOutputReadLine();
								retryProcess.BeginErrorReadLine();
								retryProcess.WaitForExit();
								
								output = outputBuilder.ToString();
								error = errorBuilder.ToString();
								
								if (retryProcess.ExitCode != 0)
								{
									ed.WriteMessage($"\nBuild Failed:\n{output}\n{error}");
									return;
								}
							}
						}
						else
						{
							ed.WriteMessage($"\nBuild Failed:\n{output}\n{error}");
							return;
						}
					}
				}

				double buildTime = sw.Elapsed.TotalSeconds;
				ed.WriteMessage($"\n[NRL] Build OK ({buildTime:F1}s)");

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
				else
				{
					// Xóa tất cả DLL/PDB cũ để tránh tích lũy rác (memory leak)
					// CHỈ xóa khi đã reload > 5 lần để tránh xóa quá thường xuyên
					_reloadCount++;
					if (_reloadCount % 5 == 0)
					{
						try
						{
							int cleaned = 0;
							foreach (string oldFile in Directory.GetFiles(netReloadDir, "MyFirstProject_*.*"))
							{
								try { File.Delete(oldFile); cleaned++; } catch { }
							}
							if (cleaned > 0)
								ed.WriteMessage($"\n[NRL] Đã dọn {cleaned} file rác cũ.");
						}
						catch { }
					}
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

				sw.Stop();
				ed.WriteMessage($"\n╔══════════════════════════════════════════╗");
				ed.WriteMessage($"\n║  NETRELOAD hoàn tất trong {sw.Elapsed.TotalSeconds:F1} giây       ║");
				ed.WriteMessage($"\n║  Assembly: {uniqueAssemblyName.Substring(0, Math.Min(30, uniqueAssemblyName.Length))}...       ║");
				ed.WriteMessage($"\n╚══════════════════════════════════════════╝");

				// Lưu ý: KHÔNG gọi show_menu ở đây vì AutoLoader.Initialize() 
				// trong assembly mới sẽ tự động gọi ShowMenu + ApplyAliases

			}
			catch (System.Exception ex)
			{
				ed.WriteMessage($"\nError: {ex.Message}");
			}
		}
	}
}
